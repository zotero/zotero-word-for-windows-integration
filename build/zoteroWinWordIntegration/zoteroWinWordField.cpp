/*
    ***** BEGIN LICENSE BLOCK *****
	
	Copyright (c) 2009  Zotero
	                    Center for History and New Media
						George Mason University, Fairfax, Virginia, USA
						http://zotero.org
	
	This program is free software: you can redistribute it and/or modify
	it under the terms of the GNU General Public License as published by
	the Free Software Foundation, either version 3 of the License, or
	(at your option) any later version.
	
	This program is distributed in the hope that it will be useful,
	but WITHOUT ANY WARRANTY; without even the implied warranty of
	MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
	GNU General Public License for more details.
	
	You should have received a copy of the GNU General Public License
	along with this program.  If not, see <http://www.gnu.org/licenses/>.
    
    Permission is granted to link statically the libraries included with
    a stock copy of Microsoft Windows. This library may not be linked, 
    directly or indirectly, with any other proprietary code.
    
    ***** END LICENSE BLOCK *****
*/

#include "zoteroWinWordDocument.h"
#include "zoteroWinWordField.h"
#include "zoteroWinWordBookmark.h"
#include "zoteroException.h"

static COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
static COleVariant covTrue((short)TRUE), covFalse((short)FALSE);
/* Implementation file */
NS_IMPL_ISUPPORTS1(zoteroWinWordField, zoteroIntegrationField)
NS_IMPL_ISUPPORTS1(zoteroWinWordEnumerator, nsISimpleEnumerator)

zoteroWinWordField::zoteroWinWordField() {}
zoteroWinWordField::zoteroWinWordField(zoteroWinWordDocument *aDoc, CField field)
{
	comField = field;
	doc = aDoc;
	init(true);
}

zoteroWinWordField::zoteroWinWordField(zoteroWinWordDocument *aDoc, CField field, CRange codeRange, CString code)
{
	comField = field;
	doc = aDoc;
	comCodeRange = codeRange;
	rawCode = code;
	init(false);
}

zoteroWinWordField::~zoteroWinWordField()
{
	doc->Release();
}

/* void delete (); */
NS_IMETHODIMP zoteroWinWordField::Delete()
{
	ZOTERO_EXCEPTION_CATCHER_START
	if(isWholeNote()) {
		// if the note contains only this field, delete the note
		if(comFootnote) {
			comFootnote.Delete();
		} else if(comEndnote) {
			comEndnote.Delete();
		}
	} else {
		// delete the field
		deleteField();
	}

	return NS_OK;
	ZOTERO_EXCEPTION_CATCHER_END
}

/* void select (); */
NS_IMETHODIMP zoteroWinWordField::Select()
{
	ZOTERO_EXCEPTION_CATCHER_START
	doc->setScreenUpdatingStatus(true);
	comTextRange.Select();
	return NS_OK;
	ZOTERO_EXCEPTION_CATCHER_END
}

/* void removeCode (); */
NS_IMETHODIMP zoteroWinWordField::RemoveCode()
{
	ZOTERO_EXCEPTION_CATCHER_START
	comField.Unlink();
	return NS_OK;
	ZOTERO_EXCEPTION_CATCHER_END
}

/* void setText (in wstring text, in boolean isRich); */
NS_IMETHODIMP zoteroWinWordField::SetText(const PRUnichar *text, PRBool isRich)
{
	ZOTERO_EXCEPTION_CATCHER_START
	doc->setScreenUpdatingStatus(false);
	
	// get current font info
	CFont0 comFont = comTextRange.get_Font();
	CString fontName = comFont.get_Name();
	float fontSize = comFont.get_Size();

	if(isRich) {
		// get a temp file
		TCHAR tempPath[MAX_PATH+1], tempFile[MAX_PATH+1];
		GetTempPath(MAX_PATH, tempPath);
		GetTempFileName(tempPath, _T("ZOTERO"), 0, tempFile);

		// allocate a buffer larger than necessary and convert from wide chars
		int length = lstrlen(text)*2;
		char *writeBuffer = new char[length+1];
		int nBytes = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK | WC_SEPCHARS, text,
			-1, writeBuffer, length, NULL, NULL);

		// open and write file
		HANDLE tempFileHandle = CreateFile(tempFile, GENERIC_WRITE, FILE_SHARE_READ, NULL,
			CREATE_ALWAYS, FILE_ATTRIBUTE_TEMPORARY, NULL);
		DWORD nWritten;
		WriteFile(tempFileHandle, writeBuffer, nBytes-1, &nWritten, NULL);
		SetEndOfFile(tempFileHandle);
		CloseHandle(tempFileHandle);
		delete[] writeBuffer;

		// read from file into range
		comTextRange.put_Text(L"");
		comTextRange.InsertFile(tempFile, &covOptional, &covFalse, &covFalse, &covFalse);
		CFile::Remove(tempFile);

		// put font back on
		init(true);
		comFont = comTextRange.get_Font();

		// need to delete the return that gets added at the end, but only if there are no
		// returns within the text to be inserted
		if(!wcsstr(text, L"\\\r") && !wcsstr(text, L"\\par") && !wcsstr(text, L"\\\n")) {
			CRange toDelete = comTextRange.get_Duplicate();
			toDelete.Collapse(0);
			toDelete.MoveStart(1, -1);
			CString test = toDelete.get_Text();
			if(toDelete.get_Text() != L"\x0d") {
				toDelete.Collapse(0);
				toDelete.MoveEnd(1, 1);
				test = toDelete.get_Text();
			}
			toDelete.put_Text(L"");
		}

		if(wcsncmp(code, BIBLIOGRAPHY_CODE, BIBLIOGRAPHY_CODE.GetLength()) == 0) {
			// set style on bibliography
			try {
				comTextRange.put_Style(L"Bibliography");
			} catch(...) {}
		}
	} else {
		CFont0 comFont = comTextRange.get_Font();
		comFont.Reset();
		comTextRange.put_Text(text);
	}
	comFont.put_Name(fontName);
	comFont.put_Size(fontSize);
	return NS_OK;
	ZOTERO_EXCEPTION_CATCHER_END
}

/* wstring getCode (); */
NS_IMETHODIMP zoteroWinWordField::GetCode(PRUnichar **_retval)
{
	ZOTERO_EXCEPTION_CATCHER_START
	if(code.IsEmpty()) return NS_ERROR_FAILURE;

	long size = code.GetLength()+1;
	*_retval = (PRUnichar *) NS_Alloc(size * sizeof(PRUnichar));
	lstrcpyn(*_retval, (LPCTSTR) code, size);
	return NS_OK;
	ZOTERO_EXCEPTION_CATCHER_END
}

/* void setCode (in wstring code); */
NS_IMETHODIMP zoteroWinWordField::SetCode(const PRUnichar *aCode)
{
	ZOTERO_EXCEPTION_CATCHER_START
	code = aCode;
	rawCode = FIELD_PREFIX+code+L' ';
	comCodeRange.put_Text(rawCode);
	return NS_OK;
	ZOTERO_EXCEPTION_CATCHER_END
}

/* unsigned short getNoteIndex (); */
NS_IMETHODIMP zoteroWinWordField::GetNoteIndex(PRUint32 *_retval)
{
	ZOTERO_EXCEPTION_CATCHER_START
	loadOffsets();
	if(comFootnote != NULL) {
		*_retval = comFootnote.get_Index();
	} else if(comEndnote != NULL) {
		*_retval = comEndnote.get_Index();
	} else {
		*_retval = 0;
	}
	return NS_OK;
	ZOTERO_EXCEPTION_CATCHER_END
}

/* boolean equals (in zoteroIntegrationField field); */
NS_IMETHODIMP zoteroWinWordField::Equals(zoteroIntegrationField *field, PRBool *_retval)
{
	ZOTERO_EXCEPTION_CATCHER_START
	zoteroWinWordField *winWordField = static_cast<zoteroWinWordField*>(field);
	*_retval = comTextRange.IsEqual(winWordField->comTextRange);
	return NS_OK;
	ZOTERO_EXCEPTION_CATCHER_END
}

/* End of implementation class template. */

bool zoteroWinWordField::operator<(zoteroWinWordField &other) {
	loadOffsets();
	other.loadOffsets();

	if(offset1 < other.offset1) {
		return true;
	} else if(offset1 == other.offset1) {
		return offset2 < other.offset2;
	}
	return false;
}

/* Load offsets for comparisons, if not already loaded */
void zoteroWinWordField::loadOffsets()
{
	if(offset1 != -1) {
		return;
	}

	comFootnote = NULL;
	comEndnote = NULL;
	long storyType = comTextRange.get_StoryType();
	if(storyType == 2) {
		// in a footnote
		CFootnotes comFootnotes = comTextRange.get_Footnotes();
		comFootnote = comFootnotes.Item(1);
	} else if(storyType == 3) {
		// in an endnote
		CEndnotes comEndnotes = comTextRange.get_Endnotes();
		comEndnote = comEndnotes.Item(1);
	}

	if(comFootnote != NULL || comEndnote != NULL) {
		CRange tempRange;
		if(comFootnote != NULL) {
			tempRange = comFootnote.get_Reference();
		} else {
			tempRange = comEndnote.get_Reference();
		}
		offset1 = tempRange.get_Start();
		offset2 = comTextRange.get_Start();
	} else {
		offset1 = comTextRange.get_Start();
		offset2 = 0;
	}
}

void zoteroWinWordField::convertToNoteType(PRUint16 toNoteType)
{
	loadOffsets();
		
	if(comFootnote != NULL && toNoteType == zoteroIntegrationDocument::NOTE_ENDNOTE) {
		// footnote to endnote
		CRange comRange = comFootnote.get_Range();
		CFootnotes comFootnotes = comRange.get_Footnotes();
		comFootnotes.Convert();
	} else if(comEndnote != NULL && toNoteType == zoteroIntegrationDocument::NOTE_FOOTNOTE) {
		// endnote to footnote
		CRange comRange = comEndnote.get_Range();
		CEndnotes comEndnotes = comRange.get_Endnotes();
		comEndnotes.Convert();
	} else if((comFootnote != NULL || comEndnote != NULL) && !toNoteType) {
		// footnote or endnote to in-text
		if(isWholeNote()) { // replace reference with citation if this is the only one in the note
			CRange comRefRange, comNoteRange;
			if(comFootnote) {
				comRefRange = comFootnote.get_Reference();
				comNoteRange = comFootnote.get_Range();
			} else {
				comRefRange = comEndnote.get_Reference();
				comNoteRange = comFootnote.get_Range();
			}
			CRange comRange = comRefRange.get_Duplicate();
			comRange.Collapse(0);
			comRange.put_FormattedText(comNoteRange);
			comRefRange.put_Text(L"");
			loadFromRange(comRange);
		}
	} else if((comEndnote == NULL && comFootnote == NULL) && toNoteType) {
		// in-text to footnote or endnote
		// get document
		CRange comRange = comTextRange.get_Duplicate();
		comRange.Collapse(0);
		comRange.Move(1, 1);
		
		// create a new note and get its range
		if(toNoteType == zoteroIntegrationDocument::NOTE_FOOTNOTE) {
			CFootnotes notes = doc->comDoc.get_Footnotes();
			CFootnote note = notes.Add(comRange, covOptional, covOptional);
			comRange = note.get_Range();
		} else if(toNoteType == zoteroIntegrationDocument::NOTE_ENDNOTE) {
			CEndnotes notes = doc->comDoc.get_Endnotes();
			CEndnote note = notes.Add(comRange, covOptional, covOptional);
			comRange = note.get_Range();
		}
		
		// put formatted text in the range
		comRange.put_FormattedText(getFieldRange());
		deleteField();

		loadFromRange(comRange);
	}
}

bool zoteroWinWordField::isWholeNote()
{
	loadOffsets();
	if(comFootnote != NULL || comEndnote != NULL) {
		CRange noteRange;

		if(comFootnote != NULL) {
			noteRange = comFootnote.get_Range();
		} else if(comEndnote != NULL) {
			noteRange = comEndnote.get_Range();
		}
		
		// if this field is the entire content of this footnote, delete the footnote
		CRange testRange = getFieldRange();
		return noteRange.IsEqual(testRange) != 0;
	}
	return false;
}

void zoteroWinWordField::init(bool needCode)
{
	offset1 = -1;

	// get ranges
	if(needCode) {
		comCodeRange = comField.get_Code();
	}
	comTextRange = comField.get_Result();

	// get rawCode
	if(rawCode.IsEmpty()) {
		rawCode = comCodeRange.get_Text();
	}
	
	// get code
	CStringW prefix;
	if(wcsncmp(rawCode, FIELD_PREFIX, FIELD_PREFIX.GetLength()) == 0) {
		prefix = FIELD_PREFIX;
	} else if(wcsncmp(rawCode, BACKUP_FIELD_PREFIX, BACKUP_FIELD_PREFIX.GetLength()) == 0) {
		prefix = BACKUP_FIELD_PREFIX;
	}
	
	if(!prefix.IsEmpty()) {
		long codeLength = rawCode.GetLength();
		if(rawCode[codeLength-1] == ' ') {
			long prefixLength = prefix.GetLength();
			code = rawCode.Mid(prefixLength, codeLength-prefixLength-1);
		} else {
			code = rawCode.Mid(prefix.GetLength());
		}
	}
}

void zoteroWinWordField::loadFromRange(CRange comRange)
{
	CFields comFields = comRange.get_Fields();
	comField = comFields.Item(1);
	init(true);
}

CRange zoteroWinWordField::getFieldRange()
{
	CRange testRange = comCodeRange.get_Duplicate();
	testRange.MoveStart(1, -1);
	testRange.put_End(comTextRange.get_End()+1);
	return testRange;
}

void zoteroWinWordField::deleteField()
{
	comField.Delete();
}

zoteroWinWordEnumerator::~zoteroWinWordEnumerator() {
	for(short i=0; i<3; i++) {
		if(fieldItem[i] != NULL) delete fieldItem[i];
	}
	doc->Release();
}

NS_IMETHODIMP zoteroWinWordEnumerator::GetNext(nsISupports **_retval) {
	ZOTERO_EXCEPTION_CATCHER_START
	// figure out which item comes next
	short minItem = -1;
	for(short i=0; i<3; i++) {
		if(fieldItem[i] != NULL && (minItem == -1 || *fieldItem[i] < *fieldItem[minItem])) {
			minItem = i;
		}
	}
	
	// prepare for return
	*_retval = fieldItem[minItem];
	(*_retval)->AddRef();
	
	// add a new one
	fetchNextItem(minItem);

	return NS_OK;
	ZOTERO_EXCEPTION_CATCHER_END
}

NS_IMETHODIMP zoteroWinWordEnumerator::HasMoreElements(PRBool *_retval) {
	ZOTERO_EXCEPTION_CATCHER_START
	*_retval = false;
	for(short i=0; i<3; i++) {
		if(fieldItem[i] != NULL) {
			*_retval = true;
			break;
		}
	}
	return NS_OK;
	ZOTERO_EXCEPTION_CATCHER_END
}

zoteroWinWordFieldEnumerator::zoteroWinWordFieldEnumerator() {}
zoteroWinWordFieldEnumerator::zoteroWinWordFieldEnumerator(zoteroWinWordDocument *aDoc) {
	doc = aDoc;
	CStoryRanges comStoryRanges = doc->comDoc.get_StoryRanges();
	CRange comStoryRange;
	CFields comFields;
	LPUNKNOWN pUnk;
	for(short i=0; i<3; i++) {
		try {
			comStoryRange = comStoryRanges.Item(i+1);
		} catch(COleDispatchException* e) {
			e->Delete();
			fieldItem[i] = NULL;
			continue;
		}
		comFields = comStoryRange.get_Fields();

		// generate enum
		pUnk = comFields.get__NewEnum();
		pUnk->QueryInterface(IID_IEnumVARIANT, (void **)&fieldEnum[i]);
		pUnk->Release();
		
		// get first zoteroWinWordField
		fetchNextItem(i);
	}
}

zoteroWinWordFieldEnumerator::~zoteroWinWordFieldEnumerator() {
	for(short i=0; i<3; i++) {
		if(fieldEnum[i] != NULL) fieldEnum[i]->Release();
	}
}

void zoteroWinWordFieldEnumerator::fetchNextItem(short i) {
	VARIANT var;
	if(fieldEnum[i]->Next(1, &var, NULL) == S_OK) {
		CField field = var.pdispVal;
		CRange codeRange = field.get_Code();
		CString codeString = codeRange.get_Text();
		if(wcsncmp(codeString, FIELD_PREFIX, FIELD_PREFIX.GetLength()) == 0
				|| wcsncmp(codeString, BACKUP_FIELD_PREFIX, BACKUP_FIELD_PREFIX.GetLength()) == 0) {
			doc->AddRef();
			fieldItem[i] = new zoteroWinWordField(doc, field, codeRange, codeString);
		} else {
			fetchNextItem(i);
		}
	} else {
		fieldItem[i] = NULL;
	}
}