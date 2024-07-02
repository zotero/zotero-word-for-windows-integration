/*
    ***** BEGIN LICENSE BLOCK *****
	
	Copyright (c) 2009-2012  Zotero
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

#include "zoteroWinWordIntegration.h"

static COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
static COleVariant covTrue((short)TRUE), covFalse((short)FALSE);
static wchar_t* FIELD_PREFIXES[] = {L" ADDIN ZOTERO_", L" CSL_", NULL};
static wchar_t* BOOKMARK_PREFIXES[] = {L"ZOTERO_", L"CSL_", NULL};

statusCode isWholeNote(field_t* field, bool* returnValue);
statusCode setTextAndNoteLocations(field_t* field);

// Allocates a field structure based on a CField, optionally checking to make
// sure that the field code actually matches a Zotero field.
statusCode initField(document_t *doc, CField comField, short noteType,
					 bool ignoreCode, field_t **returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	CRange comCodeRange = comField.get_Code();
	
	field_t *field = NULL;
	if(ignoreCode) {
		field = new field_t;
		field->code = NULL;
	} else {
		// Read code
		ENSURE_OK(prepareReadFieldCode(doc));
		CString rawCode = comCodeRange.get_Text();
		
		if(rawCode) {
			// Check that this field is a valid Zotero field
			for(unsigned short i=0; FIELD_PREFIXES[i] != NULL; i++) {
				int location = rawCode.Find(FIELD_PREFIXES[i]);
				if(location != -1) {
					field = new field_t;
					
					// If field code is all caps, make sure text object isn't in
					// all caps mode
					CString upperCaseCode = rawCode;
					upperCaseCode.MakeUpper();
					if(rawCode == upperCaseCode) {
						try {
							CFont0 comFont = comCodeRange.get_Font();
							if(comFont.get_AllCaps()) {
								comFont.put_AllCaps(false);
								rawCode = comCodeRange.get_Text();
							}
						} catch(CException* e) {
							e->Delete();
						}
					}
					
					// Get code
					int rawCodeLength = rawCode.GetLength();
					int codeStart = location + ((int) wcslen(FIELD_PREFIXES[i]));
					int codeLength = rawCodeLength - codeStart;
					
					// Sometimes there is a space at the end of the code, which
					// we ignore here
					if(rawCode.GetAt(rawCodeLength-1) == L' ') {
						codeLength--;
					}

					CString code = rawCode.Mid(codeStart, codeLength);
					field->code = _wcsdup(code);
					
					break;
				}
			}
		}
	}
	
	if(field) {
		field->text = NULL;
		field->bookmarkName = NULL;
		field->comBookmark = NULL;
		field->adjacent = false;
		
		field->doc = doc;
		field->comField = comField;
		field->comCodeRange = comCodeRange;
		field->comContentRange = comField.get_Result();
		field->noteType = noteType;
		setTextAndNoteLocations(field);
		
		// Add field to field list for document
		addValueToList(field, &(doc->allocatedFieldsStart),
					   &(doc->allocatedFieldsEnd));
		
		*returnValue = field;
		return STATUS_OK;
	}
	
	*returnValue = NULL;
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Allocates a field structure based on a WordBookmark
statusCode initBookmark(document_t *doc, CBookmark0 comBookmark, short noteType,
						bool ignoreCode, field_t **returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	field_t *field = NULL;
	
	// Read code
	CString bookmarkName = comBookmark.get_Name();
	
	if(ignoreCode) {
		field = new field_t;
		field->code = NULL;
	} else {
		// Check that this field is a valid Zotero field
		for(unsigned short i=0; BOOKMARK_PREFIXES[i] != NULL; i++) {
			int location = bookmarkName.Find(BOOKMARK_PREFIXES[i]);
			if(location != -1) {
				// Get code from preferences
				CString propertyValue;
				ENSURE_OK(getProperty(doc, bookmarkName, &propertyValue));
				
				// Check that preferences code is valid
				for(unsigned short i=0; BOOKMARK_PREFIXES[i] != NULL; i++) {
					int location = propertyValue.Find(BOOKMARK_PREFIXES[i]);
					if(location != -1) {
						field = new field_t;

						// Get code
						int rawCodeLength = propertyValue.GetLength();
						int codeStart = location + ((int) wcslen(BOOKMARK_PREFIXES[i]));
						int codeLength = rawCodeLength - codeStart;

						CString code = propertyValue.Mid(codeStart, codeLength);
						field->code = _wcsdup(code);
						break;
					}
				}
				
				if(field) break;
			}
		}
		
	}
	
	if(field) {
		field->text = NULL;
		field->comField = NULL;
		field->adjacent = false;
		
		field->bookmarkName = _wcsdup(bookmarkName);
		field->doc = doc;
		field->comBookmark = comBookmark;
		
		// Get the field contents
		field->comContentRange = comBookmark.get_Range();
		field->comCodeRange = field->comContentRange;
		field->noteType = noteType;
		setTextAndNoteLocations(field);
		
		// Add field to field list for document
		addValueToList(field, &(doc->allocatedFieldsStart),
					   &(doc->allocatedFieldsEnd));
		
		*returnValue = field;
		return STATUS_OK;
	}
	
	*returnValue = NULL;
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Frees a field struct. This does not free the corresponding field list node.
void __stdcall freeField(field_t* field) {
	if(field->text) delete field->text;
	if(field->code) delete field->code;
	if(field->bookmarkName) free(field->bookmarkName);
	delete field;
}

// Deletes this field from the document
statusCode __stdcall deleteField(field_t* field) {
	HANDLE_EXCEPTIONS_BEGIN
	bool wholeNote;
	ENSURE_OK(isWholeNote(field, &wholeNote));
	CSelection selection = field->doc->comWindow.get_Selection();
	CRange selectionRange = selection.get_Range();
	if(wholeNote) {
		// If the note contains only this field, delete the note
		if(field->noteType == NOTE_FOOTNOTE) {
			CRange noteRange = field->comFootnote.get_Range();
			// If we remove a note and the selection cursor is in it, the document ends up without a valid selection
			// so we put the selection right after the note in the reference.
			if (selection.get_Start() >= noteRange.get_Start() && selection.get_End() <= noteRange.get_End()) {
				CRange refRange = field->comFootnote.get_Reference();
				CRange dupRange = refRange.get_Duplicate();
				dupRange.Collapse(0);
				dupRange.Select();
				field->doc->insertTextIntoNote = field->noteType;
			}
			field->comFootnote.Delete();
		} else if(field->noteType == NOTE_ENDNOTE) {
			CRange noteRange = field->comEndnote.get_Range();
			if (selection.get_Start() >= noteRange.get_Start() && selection.get_End() <= noteRange.get_End()) {
				CRange refRange = field->comEndnote.get_Reference();
				CRange dupRange = refRange.get_Duplicate();
				dupRange.Collapse(0);
				dupRange.Select();
				field->doc->insertTextIntoNote = field->noteType;
			}
			field->comEndnote.Delete();
		}
	} else if(field->comBookmark) {
		field->comContentRange.put_Text(L"");
	} else {
		field->comField.Delete();
	}

	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Selects this field
statusCode __stdcall selectField(field_t* field) {
	HANDLE_EXCEPTIONS_BEGIN
	setScreenUpdatingStatus(field->doc, true);
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Removes a field code
statusCode __stdcall removeCode(field_t* field) {
	HANDLE_EXCEPTIONS_BEGIN
	if(field->comBookmark) {
		field->comBookmark.Delete();
		setProperty(field->doc, field->bookmarkName, L"");
	} else {
		field->comField.Unlink();
	}
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Gets text inside this field. DO NOT FREE THE RETURN VALUE!
statusCode __stdcall getText(field_t* field, wchar_t** returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	CString fieldText = field->comContentRange.get_Text();
	*returnValue = field->text = _wcsdup(fieldText);
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Sets text of this field
statusCode __stdcall setText(field_t* field, const wchar_t string[], bool isRich) {
	HANDLE_EXCEPTIONS_BEGIN
	setScreenUpdatingStatus(field->doc, false);
	
	// Get current font info
	CFont0 comFont = field->comContentRange.get_Font();
	CString fontName = comFont.get_Name();
	float fontSize = comFont.get_Size();

	// Check if we need to restore cursor position after insert (bookmarks only)
	bool restoreSelectionToEnd = false;
	if(field->comBookmark) {
		CSelection comSelection = (field->doc)->comWindow.get_Selection();
		CRange comSelectionRange = comSelection.get_Range();
		CRange comTestRange = field->comContentRange.get_Duplicate();
		comTestRange.Collapse(0 /*wdCollapseEnd*/);
		if(comSelectionRange.IsEqual(comTestRange)) {
			restoreSelectionToEnd = true;
		}
	}

	if(isRich) {
		// Get a temp file

		char* utf8String;
		int nBytes;
		if(field->comBookmark && wcslen(string) > 6) {
			// InsertFile method will clobber the bookmark, so add it to the RTF
			CString insertString;
			insertString.Format(L"{\\rtf {\\bkmkstart %s}{%s{\\bkmkend %s}}", field->bookmarkName, string+6, field->bookmarkName);
			nBytes = WideCharToMultiByte(CP_UTF8, 0, insertString, -1, NULL, 0, NULL, NULL);
			utf8String = new char[nBytes];
			WideCharToMultiByte(CP_UTF8, 0, insertString, -1, utf8String, nBytes, NULL, NULL);
		} else {
			// Convert to UTF-8
			nBytes = WideCharToMultiByte(CP_UTF8, 0, string, -1, NULL, 0, NULL, NULL);
			utf8String = new char[nBytes];
			WideCharToMultiByte(CP_UTF8, 0, string, -1, utf8String, nBytes, NULL, NULL);
		}

		// Open and write file
		DWORD nWritten;
		HANDLE tempFileHandle = getTemporaryFile();
		WriteFile(tempFileHandle, utf8String, nBytes-1, &nWritten, NULL);
		SetEndOfFile(tempFileHandle);
		delete[] utf8String;

		// Read from file into range
		if(!(field->comBookmark) && (field->doc)->wordVersion >= 15) {
			// In Word 2013, text does not get inserted into ranges
			field->comContentRange.put_Text(L"  ");
			CRange toDelete = field->comContentRange.get_Duplicate();
			toDelete.Collapse(0);
			CRange comDupRange = field->comContentRange.get_Duplicate();
			comDupRange.MoveEnd(1, -1);
			comDupRange.InsertFile(getTemporaryFilePath(), &covOptional, &covFalse, &covFalse, &covFalse);
			toDelete.MoveStart(1, -1);
			toDelete.put_Text(L"");
		} else {
			field->comContentRange.put_Text(L"");
			field->comContentRange.InsertFile(getTemporaryFilePath(), &covOptional, &covFalse, &covFalse, &covFalse);

			if(field->comBookmark) {
				field->comContentRange = field->comBookmark.get_Range();
				field->comCodeRange = field->comContentRange;
			} else {
				field->comContentRange = field->comField.get_Result();
			}
		}

		// Put font back on
		comFont = field->comContentRange.get_Font();

		// Need to delete the return that gets added at the end, but only if there are no
		// returns within the text to be inserted
		if(!wcsstr(string, L"\\\r") && !wcsstr(string, L"\\par") && !wcsstr(string, L"\\\n")) {
			CRange toDelete = field->comContentRange.get_Duplicate();
			toDelete.Collapse(0);
			toDelete.MoveStart(1, -1);
			if(toDelete.get_Text() != L"\x0d") {
				toDelete.Collapse(0);
				toDelete.MoveEnd(1, 1);
			}
			toDelete.put_Text(L"");
		}

		if(wcsncmp(field->code, L"BIBL", 4) == 0) {
			setStyle(field->doc, &field->comContentRange, BIBLIOGRAPHY_STYLE_ENUM, BIBLIOGRAPHY_STYLE_NAME);
		}
	} else {
		CFont0 comFont = field->comContentRange.get_Font();
		comFont.Reset();
		field->comContentRange.put_Text(string);
		if(field->comBookmark) {
			// Setting the text of the bookmark erases it, so we need to recreate it
			CBookmarks comBookmarks = field->doc->comDoc.get_Bookmarks();
			field->comBookmark = comBookmarks.Add(field->bookmarkName, field->comContentRange);
			field->comContentRange = field->comBookmark.get_Range();
		}
	}

	// Restore the selection to the end of a bookmark
	if(restoreSelectionToEnd) {
		CRange comRange = field->comContentRange.get_Duplicate();
		comRange.Collapse(0 /*wdCollapseEnd*/);
		comRange.Select();
	}

	comFont.put_Name(fontName);
	comFont.put_Size(fontSize);
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Sets the field code
statusCode __stdcall setCode(field_t *field, const wchar_t code[]) {
	HANDLE_EXCEPTIONS_BEGIN
	if(field->comBookmark) {
		CString rawCode;
		rawCode.Format(L"%s%s", BOOKMARK_PREFIXES[0], code);
		ENSURE_OK(setProperty(field->doc, field->bookmarkName, rawCode));
	} else {
		CString rawCode;
		rawCode.Format(L"%s%s ", FIELD_PREFIXES[0], code);
		field->comCodeRange.put_Text(rawCode);
	}
	
	// Store code in struct
	if(field->code) free(field->code);
	field->code = _wcsdup(code);
	
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Returns the index of the note in which this field resides
statusCode __stdcall getNoteIndex(field_t* field, unsigned long *returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	if(field->noteType == NOTE_FOOTNOTE) {
		*returnValue = field->comFootnote.get_Index();
	} else if(field->noteType == NOTE_ENDNOTE){ 
		*returnValue = field->comEndnote.get_Index();
	} else {
		*returnValue = 0;
	}
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Returns whether the field is adjacent to the next field
statusCode __stdcall isAdjacentToNextField(field_t* field, bool *returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	*returnValue = field->adjacent;
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Compares two fields to determine which comes before which
statusCode compareFields(field_t* a, field_t* b, short *returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	if(a->textLocation < b->textLocation) {
		*returnValue = -1;
		return STATUS_OK;
	} else if(b->textLocation < a->textLocation) {
		*returnValue = 1;
		return STATUS_OK;
	}
	
	// Compare positions inside a footnote
	if(a->noteType && b->noteType) {
		if(a->noteLocation < b->noteLocation) {
			*returnValue = -1;
			return STATUS_OK;
		} else if(b->noteLocation < a->noteLocation) {
			*returnValue = 1;
			return STATUS_OK;
		}
	}
	
	*returnValue = 0;
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Converts this field to a different note type. The implementation should not expect to use the
// field structure for anything else after this happens.
statusCode convertToNoteType(field_t* field, short toNoteType) {
	HANDLE_EXCEPTIONS_BEGIN
	if(field->noteType == NOTE_FOOTNOTE && toNoteType == NOTE_ENDNOTE) {
		// Footnote to endnote
		CRange comRange = field->comFootnote.get_Range();
		CFootnotes comFootnotes = comRange.get_Footnotes();
		comFootnotes.Convert();
	} else if(field->noteType == NOTE_ENDNOTE && toNoteType == NOTE_FOOTNOTE) {
		// Endnote to footnote
		CRange comRange = field->comEndnote.get_Range();
		CEndnotes comEndnotes = comRange.get_Endnotes();
		comEndnotes.Convert();
	} else {
		CRange comRange;
		if(field->noteType && !toNoteType) {
			// Footnote or endnote to in-text
			bool wholeNote;
			ENSURE_OK(isWholeNote(field, &wholeNote));
			if(wholeNote) { // Replace reference with citation if this is the only one in the note
				CRange comRefRange, comNoteRange;
				if(field->noteType == NOTE_FOOTNOTE) {
					comRefRange = field->comFootnote.get_Reference();
					comNoteRange = field->comFootnote.get_Range();
				} else {
					comRefRange = field->comEndnote.get_Reference();
					comNoteRange = field->comEndnote.get_Range();
				}
				comRange = comRefRange.get_Duplicate();
				comRange.Collapse(0);
				comRange.put_FormattedText(comNoteRange);
				comRefRange.put_Text(L"");
			}
		} else if(!field->noteType && toNoteType) {
			// In-text to footnote or endnote
			// Get document
			comRange = field->comContentRange.get_Duplicate();
			comRange.Collapse(0);
			comRange.Move(1, 1);
		
			// Create a new note and get its range
			if(toNoteType == NOTE_FOOTNOTE) {
				CFootnotes notes = (field->doc)->comDoc.get_Footnotes();
				CFootnote note = notes.Add(comRange, covOptional, covOptional);
				comRange = note.get_Range();
			} else if(toNoteType == NOTE_ENDNOTE) {
				CEndnotes notes = (field->doc)->comDoc.get_Endnotes();
				CEndnote note = notes.Add(comRange, covOptional, covOptional);
				comRange = note.get_Range();
			}
		
			// Put formatted text in the range
			CRange fieldRange;
			ENSURE_OK(getFieldRange(field, &fieldRange));
			comRange.put_FormattedText(fieldRange);
			deleteField(field);
		}

		// If a bookmark, re-create the mark
		if(field->bookmarkName) {
			CBookmarks comBookmarks = field->doc->comDoc.get_Bookmarks();
			field->comBookmark = comBookmarks.Add(field->bookmarkName, comRange);
		}
	}
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

statusCode isWholeNote(field_t* field, bool* returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	if(field->noteType) {
		CRange noteRange;

		if(field->noteType == NOTE_FOOTNOTE) {
			noteRange = field->comFootnote.get_Range();
		} else if(field->noteType == NOTE_ENDNOTE) {
			noteRange = field->comEndnote.get_Range();
		}
		
		CRange testRange;
		ENSURE_OK(getFieldRange(field, &testRange));
		*returnValue = noteRange.IsEqual(testRange) != 0;
	} else {
		*returnValue = false;
	}
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Get a range encompassing both the code range and the content range
statusCode getFieldRange(field_t* field, CRange* testRange) {
	HANDLE_EXCEPTIONS_BEGIN
	if(field->comBookmark) {
		*testRange = field->comContentRange.get_Duplicate();
	} else {
		*testRange = field->comCodeRange.get_Duplicate();
		testRange->MoveStart(1, -1);
		testRange->put_End(field->comContentRange.get_End()+1);
	}
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Sets noteType, textLocation, noteLocation, comFootnote, and comEndnote
statusCode setTextAndNoteLocations(field_t* field) {
	HANDLE_EXCEPTIONS_BEGIN
	if(field->noteType == -1) {
		long storyType = field->comCodeRange.get_StoryType();
		if(storyType == 2) {
			field->noteType = NOTE_FOOTNOTE;
		} else if(storyType == 3) {
			field->noteType = NOTE_ENDNOTE;
		} else {
			field->noteType = 0;
		}
	}

	if(field->noteType == NOTE_FOOTNOTE) {
		CFootnotes comFootnotes = field->comContentRange.get_Footnotes();
		field->comFootnote = comFootnotes.Item(1);
		field->comEndnote = NULL;
		CRange comNoteReference = field->comFootnote.get_Reference();
		field->textLocation = comNoteReference.get_Start();
		field->noteLocation = field->comCodeRange.get_Start();
	} else if(field->noteType == NOTE_ENDNOTE){ 
		CFootnotes comEndnotes = field->comContentRange.get_Endnotes();
		field->comEndnote = comEndnotes.Item(1);
		field->comFootnote = NULL;
		CRange comNoteReference = field->comEndnote.get_Reference();
		field->textLocation = comNoteReference.get_Start();
		field->noteLocation = field->comCodeRange.get_Start();
	} else {
		field->comFootnote = NULL;
		field->comEndnote = NULL;
		field->textLocation = field->comCodeRange.get_Start();
		field->noteLocation = field->textLocation;
	}
	
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}
 