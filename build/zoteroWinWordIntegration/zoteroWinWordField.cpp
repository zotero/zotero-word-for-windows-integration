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

static COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
/* Implementation file */
NS_IMPL_ISUPPORTS1(zoteroWinWordField, zoteroIntegrationField)
NS_IMPL_ISUPPORTS1(zoteroWinWordEnumerator, nsISimpleEnumerator)

zoteroWinWordField::zoteroWinWordField() {}
zoteroWinWordField::zoteroWinWordField(zoteroWinWordDocument *aDoc, CField field)
{
	try {
		comField = field;
		doc = aDoc;
		init();
	} catch(...) {}
}

zoteroWinWordField::~zoteroWinWordField()
{
	doc->Release();
}

/* void delete (); */
NS_IMETHODIMP zoteroWinWordField::Delete()
{
	try {
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
	} catch(...) {
		return NS_ERROR_FAILURE;
	}
}

/* void select (); */
NS_IMETHODIMP zoteroWinWordField::Select()
{
	try {
		comTextRange.Select();
		return NS_OK;
	} catch(...) {
		return NS_ERROR_FAILURE;
	}
}

/* void removeCode (); */
NS_IMETHODIMP zoteroWinWordField::RemoveCode()
{
	try {
		comField.Unlink();
		return NS_OK;
	} catch(...) {
		return NS_ERROR_FAILURE;
	}
}

/* void setText (in wstring text, in boolean isRich); */
NS_IMETHODIMP zoteroWinWordField::SetText(const PRUnichar *text, PRBool isRich)
{
	try {
		if(isRich) {
			// get current font info
			CFont0 comFont = comTextRange.get_Font();
			CString fontName = comFont.get_Name();
			float fontSize = comFont.get_Size();
			
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
			comTextRange.InsertFile(tempFile, &covOptional, &covOptional, &covOptional, &covOptional);
			CFile::Remove(tempFile);

			// put font back on
			init();
			comFont = comTextRange.get_Font();
			comFont.put_Name(fontName);
			comFont.put_Size(fontSize);

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
		} else {
			CFont0 comFont = comTextRange.get_Font();
			comFont.Reset();
			comTextRange.put_Text(text);
		}
		return NS_OK;
	} catch(...) {
		return NS_ERROR_FAILURE;
	}
}

/* wstring getCode (); */
NS_IMETHODIMP zoteroWinWordField::GetCode(PRUnichar **_retval)
{
	try {
		CStringW comString = comCodeRange.get_Text();
		CStringW prefix;
		if(wcsncmp(comString, FIELD_PREFIX, FIELD_PREFIX.GetLength()) == 0) {
			prefix = FIELD_PREFIX;
		} else if(wcsncmp(comString, BACKUP_FIELD_PREFIX, BACKUP_FIELD_PREFIX.GetLength()) == 0) {
			prefix = BACKUP_FIELD_PREFIX;
		} else {
			return NS_ERROR_FAILURE;
		}
		OutputDebugString(prefix);

		long length = comString.GetLength()-prefix.GetLength();
		*_retval = (PRUnichar *) NS_Alloc((length+1) * sizeof(PRUnichar));
		lstrcpyn(*_retval, ((LPCTSTR)comString)+(prefix.GetLength()), length+1);
		OutputDebugString(*_retval);
		return NS_OK;
	} catch(...) {
		return NS_ERROR_FAILURE;
	}
}

/* void setCode (in wstring code); */
NS_IMETHODIMP zoteroWinWordField::SetCode(const PRUnichar *code)
{
	try {
		comCodeRange.put_Text(FIELD_PREFIX+code);
		return NS_OK;
	} catch(...) {
		return NS_ERROR_FAILURE;
	}
}

/* unsigned short getNoteIndex (); */
NS_IMETHODIMP zoteroWinWordField::GetNoteIndex(PRUint32 *_retval)
{
	try {
		loadOffsets();
		if(comFootnote != NULL) {
			*_retval = comFootnote.get_Index();
		} else if(comEndnote != NULL) {
			*_retval = comEndnote.get_Index();
		} else {
			*_retval = 0;
		}
		return NS_OK;
	} catch(...) {
		return NS_ERROR_FAILURE;
	}
}

/* boolean equals (in zoteroIntegrationField field); */
NS_IMETHODIMP zoteroWinWordField::Equals(zoteroIntegrationField *field, PRBool *_retval)
{
	try {
		zoteroWinWordField *winWordField = dynamic_cast<zoteroWinWordField*>(field);
		*_retval = comTextRange.IsEqual(winWordField->comTextRange);
		return NS_OK;
	} catch(...) {
		return NS_ERROR_FAILURE;
	}
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

void zoteroWinWordField::init()
{
	offset1 = -1;
	comCodeRange = comField.get_Code();
	comTextRange = comField.get_Result();
}

void zoteroWinWordField::loadFromRange(CRange comRange)
{
	CFields comFields = comRange.get_Fields();
	comField = comFields.Item(1);
	init();
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
	try {
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
	} catch(...) {
		return NS_ERROR_FAILURE;
	}
}

NS_IMETHODIMP zoteroWinWordEnumerator::HasMoreElements(PRBool *_retval) {
	try {
		*_retval = false;
		for(short i=0; i<3; i++) {
			if(fieldItem[i] != NULL) {
				*_retval = true;
				break;
			}
		}
		return NS_OK;
	} catch(...) {
		return NS_ERROR_FAILURE;
	}
}

zoteroWinWordFieldEnumerator::zoteroWinWordFieldEnumerator() {}
zoteroWinWordFieldEnumerator::zoteroWinWordFieldEnumerator(zoteroWinWordDocument *aDoc) {
	try {
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
	} catch(...) {}
}

zoteroWinWordFieldEnumerator::~zoteroWinWordFieldEnumerator() {
	for(short i=0; i<3; i++) {
		if(fieldEnum[i] != NULL) fieldEnum[i]->Release();
	}
}

void zoteroWinWordFieldEnumerator::fetchNextItem(short i) {
	VARIANT var;
	if(fieldEnum[i]->Next(1, &var, NULL) == S_OK) {
		doc->AddRef();
		fieldItem[i] = new zoteroWinWordField(doc, var.pdispVal);
	} else {
		fieldItem[i] = NULL;
	}
}