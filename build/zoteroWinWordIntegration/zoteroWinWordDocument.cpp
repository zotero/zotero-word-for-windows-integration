/*
    ***** BEGIN LICENSE BLOCK *****
	
	Copyright (c) 2009  Zotero
	                    Center for History and New Media
						George Mason University, Fairfax, Virginia, USA
						http://zotero.org
	
	Zotero is free software: you can redistribute it and/or modify
	it under the terms of the GNU Affero General Public License as published by
	the Free Software Foundation, either version 3 of the License, or
	(at your option) any later version.
	
	Zotero is distributed in the hope that it will be useful,
	but WITHOUT ANY WARRANTY; without even the implied warranty of
	MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
	GNU Affero General Public License for more details.
	
	You should have received a copy of the GNU Affero General Public License
	along with Zotero.  If not, see <http://www.gnu.org/licenses/>.
    
    Permission is granted to link statically the libraries included with
    a stock copy of Microsoft Windows. This library may not be linked, 
    directly or indirectly, with any other proprietary code.
    
    ***** END LICENSE BLOCK *****
*/

#include "nsCOMPtr.h"
#include "nsIConsoleService.h"
#include "nsServiceManagerUtils.h"

#include "zoteroWinWordDocument.h"
#include "zoteroWinWordField.h"
#include "zoteroWinWordBookmark.h"
#include "zoteroException.h"

static COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

/* Global filter, since this applies on a per-application basis
 * Need to refcount */
COleMessageFilter *filter = NULL;
short filterRefs = 0;

/* Implementation file */
NS_IMPL_ISUPPORTS1(zoteroWinWordDocument, zoteroIntegrationDocument)

zoteroWinWordDocument::zoteroWinWordDocument()
{
	initFilter();
	initFromActiveObject();
}

zoteroWinWordDocument::zoteroWinWordDocument(const PRUnichar *docName) {
	initFilter();
	// attach to a specific document
	// Get a BindCtx.
	IBindCtx *pbc;
	HRESULT hr = CreateBindCtx(0, &pbc);
	if(FAILED(hr)) {
		ZOTERO_THROW_EXCEPTION("Could not get a BindCTX");
		throw NS_ERROR_FAILURE;
	}
	
	// Get running-object table.
	IRunningObjectTable *prot;
	hr = pbc->GetRunningObjectTable(&prot);
	if(FAILED(hr)) {
		ZOTERO_THROW_EXCEPTION("Could not get Running Object Table");
		pbc->Release();
		throw NS_ERROR_FAILURE;
	}

	IEnumMoniker *pem;
	hr = prot->EnumRunning(&pem);
	if(FAILED(hr)) {
		ZOTERO_THROW_EXCEPTION("Could not get moniker enumerator");
		prot->Release();
		pbc->Release();
		throw NS_ERROR_FAILURE;
	}
	
	// Start at the beginning.
	pem->Reset();
	
	ULONG fetched;
	IMoniker *pmon;
	IDispatch *pDisp = NULL;
	int n = 0;
	while(pem->Next(1, &pmon, &fetched) == S_OK) {
		LPOLESTR pName;
		pmon->GetDisplayName(pbc, NULL, &pName);

		if(wcscmp(pName, docName) == 0) {
			hr = pmon->BindToObject(pbc, NULL, IID_IDispatch, (void **)&pDisp);
			pmon->Release();
			break;
		}

		pmon->Release();
	}
	
	pmon->Release();
	prot->Release();
	pbc->Release();
	
	if(pDisp == NULL) {
		nsCOMPtr<nsIConsoleService> consoleService = do_GetService(NS_CONSOLESERVICE_CONTRACTID);
		consoleService->LogStringMessage(L"Zotero WinWord Integration: getDocument() could not find document in ROT");
		initFromActiveObject();
	} else {
		// attach our class to the running Word instance
		comDoc.AttachDispatch(pDisp);
		
		// get some useful things
		comApp = comDoc.get_Application();
		comProperties = comDoc.get_CustomDocumentProperties();
		currentScreenUpdatingStatus = true;
	}
}

zoteroWinWordDocument::~zoteroWinWordDocument() {
	filterRefs--;
	if(filterRefs <= 0 && filter != NULL) {
		filter->Revoke();
		delete filter;
		filter = NULL;
	}
}

/* short displayAlert (in wstring dialogText, in unsigned short icon, in unsigned short buttons); */
NS_IMETHODIMP zoteroWinWordDocument::DisplayAlert(const PRUnichar *dialogText, PRUint16 icon, PRUint16 buttons, PRInt16 *_retval)
{
	UINT nType = MB_SYSTEMMODAL;

	if(icon == zoteroIntegrationDocument::DIALOG_ICON_STOP) {
		nType |= MB_ICONSTOP;
	} else if(icon == zoteroIntegrationDocument::DIALOG_ICON_NOTICE) {
		nType |= MB_ICONINFORMATION;
	} else if(icon == zoteroIntegrationDocument::DIALOG_ICON_CAUTION) {
		nType |= MB_ICONEXCLAMATION;
	}

	if(buttons == zoteroIntegrationDocument::DIALOG_BUTTONS_OK) {
		nType |= MB_OK;
	} else if(buttons == zoteroIntegrationDocument::DIALOG_BUTTONS_OK_CANCEL) {
		nType |= MB_OKCANCEL;
	} else if(buttons == zoteroIntegrationDocument::DIALOG_BUTTONS_YES_NO) {
		nType |= MB_YESNO;
	} else if(buttons == zoteroIntegrationDocument::DIALOG_BUTTONS_YES_NO_CANCEL) {
		nType |= MB_YESNOCANCEL;
	}
	
	int buttonClicked = AfxMessageBox(dialogText, nType);
	
	if(_retval != NULL) {
		if(buttons == zoteroIntegrationDocument::DIALOG_BUTTONS_OK) {
			*_retval = 0;
		} else if(buttons == zoteroIntegrationDocument::DIALOG_BUTTONS_OK_CANCEL) {
			if(buttonClicked == IDOK) {
				*_retval = 1;
			} else {
				*_retval = 0;
			}
		} else if(buttons == zoteroIntegrationDocument::DIALOG_BUTTONS_YES_NO) {
			if(buttonClicked == IDYES) {
				*_retval = 1;
			} else {
				*_retval = 0;
			}
		} else if(buttons == zoteroIntegrationDocument::DIALOG_BUTTONS_YES_NO_CANCEL) {
			if(buttonClicked == IDYES) {
				*_retval = 2;
			} else if(buttonClicked == IDNO) {
				*_retval = 1;
			} else {
				*_retval = 0;
			}
		}
	}

	return NS_OK;
}

/* void activate (); */
NS_IMETHODIMP zoteroWinWordDocument::Activate()
{
	// not necessary on Windows
	return NS_OK;
}

/* boolean canInsertField (in string fieldType); */
NS_IMETHODIMP zoteroWinWordDocument::CanInsertField(const char *fieldType, PRBool *_retval)
{
	ZOTERO_EXCEPTION_CATCHER_START
	CSelection selection = comApp.get_Selection();
	long position = selection.get_StoryType();
	
	*_retval = PR_FALSE;
	if((strcmp(fieldType, "Bookmark") != 0 && (position == 2 || position == 3))		// not a bookmark, and in footnote or endnote
		|| position == 1) {															// or in main text
		*_retval = PR_TRUE;
	}
	return NS_OK;
	ZOTERO_EXCEPTION_CATCHER_END
}

/* zoteroIntegrationField cursorInField (in string fieldType); */
NS_IMETHODIMP zoteroWinWordDocument::CursorInField(const char *fieldType, zoteroIntegrationField **_retval)
{
	ZOTERO_EXCEPTION_CATCHER_START
	*_retval = NULL;
	CSelection selection = comApp.get_Selection();
	
	if(strcmp(fieldType, "Field") == 0) {
		CFields rangeFields = selection.get_Fields();
		long selectionStart = selection.get_Start();
		long selectionEnd = selection.get_End();
		
		long rangeFieldCount = rangeFields.get_Count();
		if(!rangeFieldCount) {
			// otherwise, we will need to create a range that would span any field
			CRange selectionRange = selection.get_Range();
			CRange range = selectionRange.get_Duplicate();
			CRange rangeEnd = selectionRange.get_Duplicate();
			range = range.GoToPrevious(7);							// go to previous field
			rangeEnd = rangeEnd.GoToNext(7);						// go to next field

			// might only be one field in the entire doc
			long rangeEndIndex = rangeEnd.get_Start();
			if(range.get_Start() == rangeEndIndex) {
				range = range.GoToPrevious(3);		// move a line back
			}
			if(rangeEndIndex < selectionEnd) {
				rangeEndIndex = selectionEnd;		// span at least to selection end
			}

			// make range span from previous field to next field
			range.put_End(rangeEndIndex);

			// check all fields to see if they are in the selection
			rangeFields = range.get_Fields();
			rangeFieldCount = rangeFields.get_Count();
		}

		for(long i=0; i<rangeFieldCount; i++) {
			CField testField = rangeFields.Item(i+1);
			CRange testFieldCode, testFieldResult;
			
			// for whatever reason, the result can be undefined on a field
			// (I've seen it with my own eyes!)
			try {
				testFieldCode = testField.get_Code();
				testFieldResult = testField.get_Result();
			} catch(COleDispatchException *e) {
				e->Delete();
				continue;
			}

			CString testFieldCodeText = testFieldCode.get_Text();
			long testFieldStart = testFieldCode.get_Start();
			long testFieldEnd = testFieldResult.get_End();
			
			// if there is no overlap, continue
			if((testFieldStart > selectionStart && testFieldEnd > selectionEnd
					&& testFieldStart > selectionEnd && testFieldEnd > selectionEnd)
					|| (testFieldStart < selectionStart && testFieldEnd < selectionEnd
					&& testFieldStart < selectionEnd && testFieldEnd < selectionEnd)) continue;
			
			// otherwise, check for an appropriate code
			if(wcsncmp(testFieldCodeText, FIELD_PREFIX, FIELD_PREFIX.GetLength()) == 0
					|| wcsncmp(testFieldCodeText, BACKUP_FIELD_PREFIX, BACKUP_FIELD_PREFIX.GetLength()) == 0) {
				*_retval = new zoteroWinWordField(this, testField);
				AddRef();
				(*_retval)->AddRef();
				break;
			}
		}
	
		return NS_OK;
	}
	
	if(strcmp(fieldType, "Bookmark") == 0) {
		CBookmarks bookmarks = selection.get_Bookmarks();
		long count = bookmarks.get_Count();
		for(long i=0; i<count; i++) {
			CBookmark0 testBookmark = bookmarks.Item(i+1);
			CString testBookmarkCode = testBookmark.get_Name();

			if(wcsncmp(testBookmarkCode, BOOKMARK_REFERENCE_PROPERTY, BOOKMARK_REFERENCE_PROPERTY.GetLength()) == 0 ||
					wcsncmp(testBookmarkCode, BACKUP_BOOKMARK_REFERENCE_PROPERTY, BACKUP_BOOKMARK_REFERENCE_PROPERTY.GetLength()) == 0) {
				*_retval = new zoteroWinWordBookmark(this, testBookmark);
				AddRef();
				(*_retval)->AddRef();
				break;
			}
		}

		return NS_OK;
	}

	return NS_ERROR_NOT_IMPLEMENTED;
	ZOTERO_EXCEPTION_CATCHER_END
}

/* wstring getDocumentData (); */
NS_IMETHODIMP zoteroWinWordDocument::GetDocumentData(PRUnichar **_retval)
{
	ZOTERO_EXCEPTION_CATCHER_START
	CStringW comString = getProperty(PREFS_PROPERTY);
	if(comString == "") {
		CStringW comString = getProperty(BACKUP_PREFS_PROPERTY);
	}
	long length = comString.GetLength();
	*_retval = (PRUnichar *) NS_Alloc((length+1) * sizeof(PRUnichar));
	lstrcpyn(*_retval, comString, length+1);
	return NS_OK;
	ZOTERO_EXCEPTION_CATCHER_END
}

/* void setDocumentData (in wstring data); */
NS_IMETHODIMP zoteroWinWordDocument::SetDocumentData(const PRUnichar *data)
{
	ZOTERO_EXCEPTION_CATCHER_START
	CString dataString(data);
	setProperty(PREFS_PROPERTY, dataString);
	return NS_OK;
	ZOTERO_EXCEPTION_CATCHER_END
}

/* zoteroIntegrationField insertField (in string fieldType, in unsigned short noteType); */
NS_IMETHODIMP zoteroWinWordDocument::InsertField(const char *fieldType, PRUint16 noteType, zoteroIntegrationField **_retval)
{
	ZOTERO_EXCEPTION_CATCHER_START
	setScreenUpdatingStatus(false);

	CSelection selection = comApp.get_Selection();
	
	CRange tempRange = selection.get_Range();
	CRange insertRange = tempRange.get_Duplicate();
	
	if(insertRange.get_StoryType() == 1) {
		// if inserting a note citation in the main story, we need to make a new note
		if(noteType == zoteroIntegrationDocument::NOTE_FOOTNOTE) {
			CFootnotes notes = comDoc.get_Footnotes();
			CFootnote note = notes.Add(insertRange, covOptional, covOptional);
			// move cursor back to main text
			insertRange.Select();
			// now inserting field into note
			insertRange = note.get_Range();
		} else if(noteType == zoteroIntegrationDocument::NOTE_ENDNOTE) {
			CEndnotes notes = comDoc.get_Endnotes();
			CEndnote note = notes.Add(insertRange, covOptional, covOptional);
			// move cursor back to main text
			insertRange.Select();
			// now inserting field into note
			insertRange = note.get_Range();
		}
	}

	// make field
	nsresult success = makeNewField(fieldType, insertRange, _retval);
	if(success == NS_OK) {
		(*_retval)->SetCode(L"");
	}
	
	return success;
	ZOTERO_EXCEPTION_CATCHER_END
}

/* nsISimpleEnumerator getFields (in string fieldType); */
NS_IMETHODIMP zoteroWinWordDocument::GetFields(const char *fieldType, nsISimpleEnumerator **_retval)
{
	ZOTERO_EXCEPTION_CATCHER_START
	if(strcmp(fieldType, "Field") == 0) {
		*_retval = new zoteroWinWordFieldEnumerator(this);
	} else if(strcmp(fieldType, "Bookmark") == 0) {
		*_retval = new zoteroWinWordBookmarkEnumerator(this);
	} else {
		return NS_ERROR_NOT_IMPLEMENTED;
	}

	AddRef();
	(*_retval)->AddRef();
	return NS_OK;
	ZOTERO_EXCEPTION_CATCHER_END
}

/* void convert (in nsISimpleEnumerator fields, in string toFieldType, in nsISimpleEnumerator toNoteType); */
NS_IMETHODIMP zoteroWinWordDocument::Convert(nsISimpleEnumerator *fields, const char *toFieldType, PRUint16 *toNoteType, PRUint32 count)
{
	ZOTERO_EXCEPTION_CATCHER_START
	setScreenUpdatingStatus(false);

	long i = 0;
	PRBool moreElements;
	nsISupports *xpcomField;
	zoteroWinWordField *field;
	zoteroWinWordBookmark *bookmark;

	while(fields->HasMoreElements(&moreElements) == NS_OK && moreElements) {
		fields->GetNext((nsISupports **)&xpcomField);
		if(xpcomField->QueryInterface(zoteroIntegrationField::COMTypeInfo<int>::kIID, (void **)&field) != NS_OK) {
			return NS_ERROR_CANNOT_CONVERT_DATA;
		}
		xpcomField->Release();
		
		bookmark = dynamic_cast<zoteroWinWordBookmark*>(field);
		bool convertToBookmark = bookmark == NULL && strcmp(toFieldType, "Bookmark") == 0;
		bool convertToField = bookmark != NULL && strcmp(toFieldType, "Field") == 0;
		if(convertToBookmark || convertToField) {
			zoteroWinWordField *newField;

			// save and remove field code
			CRange insertRange = field->getFieldRange();
			wchar_t *oldCode;
			field->GetCode(&oldCode);
				
			// create new field
			makeNewField(toFieldType, insertRange, (zoteroIntegrationField **)&newField);

			// restore code
			newField->SetCode(oldCode);
			NS_Free(oldCode);

			// free old field and use new one
			field->Release();
			field = newField;
		}
		field->convertToNoteType(toNoteType[i]);

		field->Release();
		i++;
	}
	
	return NS_OK;
	ZOTERO_EXCEPTION_CATCHER_END
}

/* void setBibliographyStyle (in long firstLineIndent, in long bodyIndent, in unsigned long lineSpacing, in unsigned long entrySpacing, [array, size_is (tabStopCount)] in long tabStops, in unsigned long tabStopCount); */
NS_IMETHODIMP zoteroWinWordDocument::SetBibliographyStyle(PRInt32 firstLineIndent, PRInt32 bodyIndent, PRUint32 lineSpacing, PRUint32 entrySpacing, PRInt32 *tabStops, PRUint32 tabStopCount)
{
	ZOTERO_EXCEPTION_CATCHER_START
	// get bibliography style
	CStyles styles = comDoc.get_Styles();
	CStyle style;
	try {
		style = styles.Item(BIBLIOGRAPHY_STYLE);
	} catch(COleDispatchException* e) {
		style = styles.Add(BIBLIOGRAPHY_STYLE,
			COleVariant((short) 1) /* wdStyleTypeParagraph */);
		e->Delete();
	}
	CParagraphFormat paraFormat = style.get_ParagraphFormat();
	paraFormat.put_FirstLineIndent(((float) firstLineIndent)/20);
	paraFormat.put_LeftIndent(((float) bodyIndent)/20);
	paraFormat.put_LineSpacing(((float) lineSpacing)/20);
	paraFormat.put_SpaceAfter(((float) entrySpacing)/20);
	CTabStops comTabStops = paraFormat.get_TabStops();
	comTabStops.ClearAll();

	for(PRUint32 i=0; i<tabStopCount; i++) {
		comTabStops.Add(((float) tabStops[i])/20,
			COleVariant((short) 0) /* wdAlignTabLeft */,
			COleVariant((short) 0) /* wdTabLeaderSpaces */);
	}

    return NS_OK;
	ZOTERO_EXCEPTION_CATCHER_END
}

/* void cleanup (); */
NS_IMETHODIMP zoteroWinWordDocument::Cleanup()
{
	setScreenUpdatingStatus(true);
	return NS_OK;
}

/* End of implementation class template. */

CString zoteroWinWordDocument::getProperty(CString propertyName)
{
	int i = 0;
	CString propertyValue;
	CString comPropertyName;
	CCustomProperty comProperty;
	
	// get all existing properties corresponding to this propertyName
	while(true) {
		i++;
		comPropertyName.Format(_T("%s_%d"), propertyName, i);

		try {
			comProperty = comProperties.Item(comPropertyName);
			propertyValue += comProperty.get_Value();
		} catch(COleDispatchException* e) {
			e->Delete();
			break;
		}
	}
	
	return propertyValue;
}

void zoteroWinWordDocument::setProperty(CString propertyName, CString propertyValue)
{
	long i = 0;
	long propertyValueLength = propertyValue.GetLength();
	CString subpropertyName, subpropertyValue;
	CCustomProperty comProperty;
	
	// change existing properties
	while(true) {
		subpropertyName.Format(_T("%s_%i"), propertyName, (i+1));

		if(propertyValueLength > (i*MAX_PROPERTY_LENGTH)) {
			// if more to add, add it
			subpropertyValue = propertyValue.Mid(i*MAX_PROPERTY_LENGTH, MAX_PROPERTY_LENGTH);
			try {
				comProperty = comProperties.Item(subpropertyName);
				comProperty.put_Value(subpropertyValue);
			} catch(COleDispatchException* e) {
				e->Delete();
				comProperties.Add(subpropertyName, false, 4, subpropertyValue);
			}
		} else {
			// if no more to add, delete extraneous properties until there are no more
			try {
				comProperty = comProperties.Item(subpropertyName);
				comProperty.Delete();
			} catch(COleDispatchException* e) {
				e->Delete();
				break;
			}
		}

		i++;
	}
}

CString zoteroWinWordDocument::getRandomString(int length) {
	// seed random number generator
	static bool rndGeneratorSeeded = false;
	if(rndGeneratorSeeded == false) {
		srand((unsigned)time(0));
		rndGeneratorSeeded = true;
	}
	
	// generator random string of desired length
	CString randString = L"";
	static const CString characters = L"abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
	for(int i=0; i<length; i++) {
		randString += characters.GetAt(rand() % 62);
	}
	return randString;
}

/* make a new field at insertRange, without setting the field code */
nsresult zoteroWinWordDocument::makeNewField(const char *fieldType, CRange insertRange, zoteroIntegrationField **_retval) {
	if(strcmp(fieldType, "Field") == 0) {
		CFields comFields = comDoc.get_Fields();
		// creating the field as a wdFieldQuote type, since creating as a wdFieldAddin field screws up the ranges
		CField field = comFields.Add(insertRange, COleVariant((short) 35), COleVariant(FIELD_PLACEHOLDER), COleVariant((short)true));

		*_retval = new zoteroWinWordField(this, field);
	} else if(strcmp(fieldType, "Bookmark") == 0) {
		CBookmarks comBookmarks = comDoc.get_Bookmarks();
		insertRange.put_Text(FIELD_PLACEHOLDER);
		CBookmark0 bookmark = comBookmarks.Add(BOOKMARK_REFERENCE_PROPERTY+getRandomString(12), insertRange);
		
		*_retval = new zoteroWinWordBookmark(this, bookmark);
	} else {
		return NS_ERROR_NOT_IMPLEMENTED;
	}
	
	AddRef();
	(*_retval)->AddRef();
	return NS_OK;
}

/* turn on or off screenUpdating */
void zoteroWinWordDocument::setScreenUpdatingStatus(bool status) {
	if(status != currentScreenUpdatingStatus) {
		comApp.put_ScreenUpdating(status);
		currentScreenUpdatingStatus = status;
	}
}

/* attaches this document to the active document of the first launched WinWord instance */
void zoteroWinWordDocument::initFromActiveObject() {
	IUnknown *pUnk = NULL;
	IDispatch *pDisp = NULL;
	// get the class ID for Word
	CLSID clsid;
	CLSIDFromProgID(L"Word.Application", &clsid); 
	// get running Word instance
	GetActiveObject(clsid, NULL, (IUnknown**)&pUnk);
	if(pUnk == NULL) {
		ZOTERO_THROW_EXCEPTION("Could not find a running Word instance.");
		throw NS_ERROR_FAILURE;
	}
	pUnk->QueryInterface(IID_IDispatch, (void **)&pDisp);
	// attach our class to the running Word instance
	comApp.AttachDispatch(pDisp);
	pUnk->Release();
	
	// get some useful things
	comDoc = comApp.get_ActiveDocument();
	comProperties = comDoc.get_CustomDocumentProperties();
	currentScreenUpdatingStatus = true;
}

/**
 * Initialize a COleMessageFilter to deal with Word being busy
 */
void zoteroWinWordDocument::initFilter() {
	if(filter == NULL) {
		filter = new COleMessageFilter();
		filter->EnableBusyDialog(false);
		filter->EnableNotRespondingDialog(false);
		filter->SetRetryReply(250);
		filter->Register();
		filterRefs = 1;
	} else {
		filterRefs++;
	}
}