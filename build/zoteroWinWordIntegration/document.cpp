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
COleMessageFilter *comFilter = NULL;
unsigned int comFilterRefs = 0;

static short LINK_PLACEHOLDER_LENGTH = 6;
static wchar_t *IMPORT_LINK_URL = L"https://www.zotero.org/";
static wchar_t *IMPORT_ITEM_PREFIX = L"ITEM CSL_CITATION ";
static wchar_t *IMPORT_BIBL_PREFIX = L"BIBL ";
static wchar_t *IMPORT_DOC_PREFS_PREFIX = L"DOCUMENT_PREFERENCES ";
// ZOTERO_EXPORTED_DOCUMENT is for legacy/old beta support and may be removed at a later date.
static wchar_t *EXPORTED_DOCUMENT_MARKER[] = { L"ZOTERO_TRANSFER_DOCUMENT", L"ZOTERO_EXPORTED_DOCUMENT", NULL };

statusCode addNextFieldToList(document_t* doc, field_t* currentFields[], listNode_t** fieldListStart, listNode_t** fieldListEnd,
                              short* noteType, bool* noMoreFields);
statusCode getNextFieldFromEnumerator(document_t* doc, IEnumVARIANT* enumerator, short noteType, field_t** returnValue);
statusCode getNextBookmark(document_t* doc, CBookmarks comBookmarks, unsigned long* bookmarkIndex,
	                       const unsigned long bookmarkCount, short noteType, field_t** returnValue);
void freeFieldList(listNode_t* fieldList, bool freeFields);

statusCode __stdcall getDocument(const wchar_t documentName[], document_t** returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	// Initialize OLE message filter
	if(comFilter == NULL) {
		comFilter = new COleMessageFilter();
		comFilter->EnableBusyDialog(false);
		comFilter->EnableNotRespondingDialog(false);
		comFilter->SetRetryReply(250);
		comFilter->Register();
	}
	comFilterRefs++;

	document_t* doc = new document_t();
	doc->comDoc = NULL;

	if(documentName != NULL) {
		// Try to attach to a specific document by finding it in the Running Object Table.
		// Convert path to document to a UNC path
		DWORD bufferLength = 1024;
		wchar_t buffer[1024];
		UNIVERSAL_NAME_INFO* unameInfo = (UNIVERSAL_NAME_INFO*) &buffer;
		wchar_t *canonicalPath;
		if(WNetGetUniversalName(documentName, UNIVERSAL_NAME_INFO_LEVEL, unameInfo, &bufferLength) == NO_ERROR) {
			canonicalPath = unameInfo->lpUniversalName;
		} else {
			wcscpy_s(buffer, bufferLength, documentName);
			canonicalPath = buffer;
		}

		// Get a BindCtx.
		IBindCtx *pbc;
		HRESULT hr = CreateBindCtx(0, &pbc);
		if(FAILED(hr)) {
			DIE(L"Could not get a BindCTX");
		}
	
		// Get running-object table.
		IRunningObjectTable *prot;
		hr = pbc->GetRunningObjectTable(&prot);
		if(FAILED(hr)) {
			pbc->Release();
			DIE(L"Could not get Running Object Table");
		}

		IEnumMoniker *pem;
		hr = prot->EnumRunning(&pem);
		if(FAILED(hr)) {
			prot->Release();
			pbc->Release();
			DIE(L"Could not get moniker enumerator");
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

			if(wcscmp(pName, canonicalPath) == 0) {
				hr = pmon->BindToObject(pbc, NULL, IID_IDispatch, (void **)&pDisp);
				pmon->Release();
				break;
			}

			pmon->Release();
		}
	
		pmon->Release();
		prot->Release();
		pbc->Release();
		
		if(pDisp != NULL) {
			// Attach our class to the running Word instance
			doc->comDoc.AttachDispatch(pDisp);
			doc->comApp = doc->comDoc.get_Application();
		}
	}

	if(doc->comDoc == NULL) {
		// Either we don't have a document name, or we failed to find it in the ROT, so try getting the active document from Word.
		IUnknown *pUnk = NULL;
		IDispatch *pDisp = NULL;
		// Get the class ID for Word
		CLSID clsid;
		CLSIDFromProgID(L"Word.Application", &clsid); 
		// Get running Word instance
		GetActiveObject(clsid, NULL, (IUnknown**)&pUnk);
		if(pUnk == NULL) {
			DIE(L"Could not find a running Word instance.");
		}
		pUnk->QueryInterface(IID_IDispatch, (void **)&pDisp);
		// Attach our class to the running Word instance
		doc->comApp.AttachDispatch(pDisp);
		pUnk->Release();

		doc->comDoc = doc->comApp.get_ActiveDocument();
	}

	// We have the document, so initialize the rest of our document_t
	doc->comProperties = doc->comDoc.get_CustomDocumentProperties();
	doc->statusScreenUpdating = true;

	CString version = doc->comApp.get_Version();
	CString frontVersion = version.Left(version.Find(_T('.')));
	doc->wordVersion = _ttoi(frontVersion);
	
	// Switch to wdRevisionsViewFinal, since wdRevisionsViewOriginal won't contain new fields
	doc->comWindow = doc->comDoc.get_ActiveWindow();
	CView0 comView = doc->comWindow.get_View();
	try {
		if(comView.get_RevisionsView() != 0 /*wdRevisionsViewFinal*/) {
			comView.put_RevisionsView(0);
		}
	} catch(CException *e) {
		e->Delete();
	}

	// Disable ShowRevisions if necessary
	try {
		doc->statusShowRevisions = doc->restoreShowRevisions = doc->comDoc.get_ShowRevisions();
		if (doc->wordVersion >= 15) {
			CRevisionsFilter revisionsFilter = comView.get_RevisionsFilter();
			doc->restoreRevisionsMarkup = revisionsFilter.get_Markup();
		}
	} catch(CException *e) {
		e->Delete();
		doc->statusShowRevisions = doc->restoreRevisionsMarkup = false;
		doc->restoreRevisionsMarkup = 0;
	}

	// Disable Track Changes if necessary
	try {
		doc->restoreTrackChanges = doc->comDoc.get_TrackRevisions();
		if (doc->restoreTrackChanges) {
			doc->comDoc.put_TrackRevisions(false);
		}
	}
	catch (CException *e) {
		e->Delete();
		doc->restoreTrackChanges = false;
	}

	// Disable Show Field Codes if enabled
	try {
		if (comView.get_ShowFieldCodes()) {
			comView.put_ShowFieldCodes(false);
		}
	}
	catch (CException *e) {
		e->Delete();
	}

	// Create an Undo Record, so we can undo Zotero actions in a single go.
	// Only available in Word 2010 and later
	if (doc->wordVersion >= 14) {
		CUndoRecord undoRecord = doc->comApp.get_UndoRecord();
		undoRecord.StartCustomRecord(UNDO_RECORD_NAME);
	}
	
	doc->allocatedFieldsStart = NULL;
	doc->allocatedFieldsEnd = NULL;
	doc->allocatedFieldListsStart = NULL;
	doc->allocatedFieldListsEnd = NULL;
	doc->insertTextIntoNote = 0;

	*returnValue = doc;
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Frees a document struct. 
void __stdcall freeDocument(document_t* doc) {
	// Remove COleMessageFilter
	comFilterRefs--;
	if(comFilterRefs <= 0 && comFilter != NULL) {
		comFilter->Revoke();
		delete comFilter;
		comFilter = NULL;
	}

	// Free allocated fields
	freeFieldList(doc->allocatedFieldsStart, true);
	
	// Free allocated field lists
	listNode_t* nextNode = doc->allocatedFieldListsStart;
	while(nextNode) {
		listNode_t* currentNode = nextNode;
		freeFieldList((listNode_t*) currentNode->value, false);
		nextNode = currentNode->next;
		free(currentNode);
	}
	
	// Free document
	delete doc;
}

// Free a field list. The second parameter determines whether the fields inside
// the list are freed, or just the list itself.
void freeFieldList(listNode_t* fieldList, bool freeFields) {
	listNode_t* nextNode = fieldList;
	while(nextNode) {
		listNode_t* currentNode = nextNode;
		if(freeFields) {
			freeField((field_t*) currentNode->value);
		}
		nextNode = currentNode->next;
		free(currentNode);
	}
}

// Displays an alert within Word
statusCode __stdcall displayAlert(document_t *doc, const wchar_t dialogText[],
						unsigned short icon, unsigned short buttons,
						unsigned short* returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	UINT nType = MB_SYSTEMMODAL;

	if(icon == DIALOG_ICON_STOP) {
		nType |= MB_ICONSTOP;
	} else if(icon == DIALOG_ICON_NOTICE) {
		nType |= MB_ICONINFORMATION;
	} else if(icon == DIALOG_ICON_CAUTION) {
		nType |= MB_ICONEXCLAMATION;
	}

	if(buttons == DIALOG_BUTTONS_OK) {
		nType |= MB_OK;
	} else if(buttons == DIALOG_BUTTONS_OK_CANCEL) {
		nType |= MB_OKCANCEL;
	} else if(buttons == DIALOG_BUTTONS_YES_NO) {
		nType |= MB_YESNO;
	} else if(buttons == DIALOG_BUTTONS_YES_NO_CANCEL) {
		nType |= MB_YESNOCANCEL;
	}
	
	int buttonClicked = AfxMessageBox(dialogText, nType);
	
	if(buttons == DIALOG_BUTTONS_OK) {
		*returnValue = 0;
	} else if(buttons == DIALOG_BUTTONS_OK_CANCEL) {
		if(buttonClicked == IDOK) {
			*returnValue = 1;
		} else {
			*returnValue = 0;
		}
	} else if(buttons == DIALOG_BUTTONS_YES_NO) {
		if(buttonClicked == IDYES) {
			*returnValue = 1;
		} else {
			*returnValue = 0;
		}
	} else if(buttons == DIALOG_BUTTONS_YES_NO_CANCEL) {
		if(buttonClicked == IDYES) {
			*returnValue = 2;
		} else if(buttonClicked == IDNO) {
			*returnValue = 1;
		} else {
			*returnValue = 0;
		}
	}

	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Disables Track Changes settings so that we can read field codes
statusCode prepareReadFieldCode(document_t* doc) {
	HANDLE_EXCEPTIONS_BEGIN
	if(doc->statusShowRevisions == false) return STATUS_OK;
	
	if(doc->wordVersion >= 14) {
		// Ugh. We can't turn off track changes in Word 2010 if the cursor is in a note.
		CSelection selection = doc->comWindow.get_Selection();
		if(selection.get_StoryType() != 1) {
			// Duplicate old selection range
			CRange selectionRange = selection.get_Range();
			CRange duplicateRange = selectionRange.get_Duplicate();
			
			// Select main text range
			CStoryRanges storyRanges = doc->comDoc.get_StoryRanges();
			CRange mainTextRange = storyRanges.Item(1);
			mainTextRange.Select();
			
			// Turn off "Show Insertions and Deletions"
			doc->comDoc.put_ShowRevisions(false);
			
			// Move selection back to note
			duplicateRange.Select();
		} else {
			doc->comDoc.put_ShowRevisions(false);
		}
	} else {
		doc->comDoc.put_ShowRevisions(false);
	}
	doc->statusShowRevisions = false;

	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Determines whether it is possible to insert a field at the current cursor
// position
statusCode __stdcall canInsertField(document_t *doc, const wchar_t fieldType[],
						            bool* returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	CSelection selection = doc->comWindow.get_Selection();
	long position = selection.get_StoryType();
	
	*returnValue = false;
	if((wcscmp(fieldType, L"Bookmark") != 0 && (position == 2 || position == 3))	// Not a bookmark, and in footnote or endnote
		|| position == 1) {															// or in main text
		*returnValue = true;
	}
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Determines whether the cursor is in a field. Returns the a field struct if
// it is, or NULL if it is not.
statusCode __stdcall cursorInField(document_t *doc, const wchar_t fieldType[],
						           field_t** returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	*returnValue = NULL;
	CSelection selection = doc->comWindow.get_Selection();
	
	if(wcscmp(fieldType, L"Field") == 0) {
		ENSURE_OK(prepareReadFieldCode(doc))

		CFields rangeFields = selection.get_Fields();
		long selectionStart = selection.get_Start();
		long selectionEnd = selection.get_End();
		long selectionStoryType = selection.get_StoryType();
		
		long rangeFieldCount = rangeFields.get_Count();
		if(!rangeFieldCount) {
			// If there are no fields in the current selection, we will need to create a range that would span any field
			CRange selectionRange = selection.get_Range();
			CRange range = selectionRange.get_Duplicate();
			CRange rangeEnd = selectionRange.get_Duplicate();
			range = range.GoToPrevious(7);							// Go to previous field
			rangeEnd = rangeEnd.GoToNext(7);						// Go to next field

			// Might only be one field in the entire doc
			long rangeStartIndex = range.get_Start();
			long rangeEndIndex = rangeEnd.get_Start();
			if (rangeStartIndex == selectionStart ||
				rangeEndIndex <= selectionEnd) {
				CStoryRanges storyRanges = doc->comDoc.get_StoryRanges();
				CRange storyRange = storyRanges.Item(selectionStoryType);
				range = storyRange.get_Duplicate();
				if (rangeStartIndex != selectionStart) {
					range.put_Start(rangeStartIndex);
				}
				if (rangeEndIndex <= selectionEnd) {
					rangeEndIndex = range.get_End();
				}
			}

			// Make range span from previous field to next field
			range.put_End(rangeEndIndex);

			// Check all fields to see if they are in the selection
			rangeFields = range.get_Fields();
			rangeFieldCount = rangeFields.get_Count();
		}

		for(long i=0; i<rangeFieldCount; i++) {
			CField testField = rangeFields.Item(i+1);
			CRange testFieldCode, testFieldResult;
			
			// For whatever reason, the result can be undefined on a field
			// (I've seen it with my own eyes!)
			try {
				testFieldCode = testField.get_Code();
				testFieldResult = testField.get_Result();
			} catch(COleDispatchException *e) {
				e->Delete();
				continue;
			}

			long testFieldStart = testFieldCode.get_Start()-1;
			long testFieldEnd = testFieldResult.get_End()+1;
			
			// If there is no overlap, continue
			if (testFieldCode.get_StoryType() != selectionStoryType ||
				testFieldStart > selectionEnd || testFieldEnd < selectionStart) continue;
			
			// Otherwise, check for an appropriate code
			statusCode status = initField(doc, testField, -1, false, returnValue);
			if(status || *returnValue) return status;
		}
	
		return STATUS_OK;
	}
	
	if(wcscmp(fieldType, L"Bookmark") == 0) {
		CBookmarks bookmarks = selection.get_Bookmarks();
		long count = bookmarks.get_Count();
		for(long i=0; i<count; i++) {
			CBookmark0 testBookmark = bookmarks.Item(i+1);
			statusCode status = initBookmark(doc, testBookmark, -1, false, returnValue);
			if(status || *returnValue) return status;
		}

		return STATUS_OK;
	}

	DIE(L"Field type not implemented.")
	HANDLE_EXCEPTIONS_END
}

// Gets document data. The caller should free the returned string.
statusCode __stdcall getDocumentData(document_t *doc, wchar_t **returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	CRange range = doc->comDoc.get_Content();
	range.Collapse(1 /*wdCollapseStart*/);
	range.MoveEnd(4 /*wdParagraph*/, 1);
	CString text = range.get_Text();
	for (unsigned short i = 0; EXPORTED_DOCUMENT_MARKER[i] != NULL; i++) {
		if (text.Find(EXPORTED_DOCUMENT_MARKER[i]) == 0) {
			*returnValue = _wcsdup(EXPORTED_DOCUMENT_MARKER[0]);
			return STATUS_OK;
		}
	}
	
	CString returnString;
	ENSURE_OK(getProperty(doc, PREFS_PROPERTY, &returnString));
	*returnValue = _wcsdup(returnString);
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Sets document data
statusCode __stdcall setDocumentData(document_t *doc, const wchar_t documentData[]) {
	HANDLE_EXCEPTIONS_BEGIN
	return setProperty(doc, PREFS_PROPERTY, documentData);
	HANDLE_EXCEPTIONS_END
}

// Makes a field at the selection
statusCode __stdcall insertField(document_t *doc, const wchar_t fieldType[],
					             unsigned short noteType, field_t **returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	setScreenUpdatingStatus(doc, false);

	CSelection selection = doc->comWindow.get_Selection();
	
	CRange tempRange = selection.get_Range();
	CRange insertRange = tempRange.get_Duplicate();
	
	if(insertRange.get_StoryType() == 1) {
		// If inserting a note citation in the main story, we need to make a new note
		if(noteType == NOTE_FOOTNOTE) {
			CFootnotes notes = doc->comDoc.get_Footnotes();
			CFootnote note = notes.Add(insertRange, covOptional, covOptional);
			// Move cursor back to main text
			CRange referenceRange = note.get_Reference();
			CRange dupRange = referenceRange.get_Duplicate();
			dupRange.Collapse(0 /*wdCollapseEnd*/);
			dupRange.Select();
			// Now inserting field into note
			insertRange = note.get_Range();
		} else if(noteType == NOTE_ENDNOTE) {
			CEndnotes notes = doc->comDoc.get_Endnotes();
			CEndnote note = notes.Add(insertRange, covOptional, covOptional);
			// Move cursor back to main text
			CRange referenceRange = note.get_Reference();
			CRange dupRange = referenceRange.get_Duplicate();
			dupRange.Collapse(0 /*wdCollapseEnd*/);
			dupRange.Select();
			// Now inserting field into note
			insertRange = note.get_Range();
		}
	}

	// Make field
	ENSURE_OK(insertFieldRaw(doc, fieldType, insertRange, returnValue))

	// If a bookmark, move selection after field
	if(wcscmp(fieldType, L"Bookmark") == 0) {
		CRange comRange = (*returnValue)->comContentRange.get_Duplicate();
		comRange.Collapse(0 /*wdCollapseEnd*/);
		comRange.Select();
	}

	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Retrieve all fields from the document
statusCode __stdcall getFields(document_t *doc, const wchar_t fieldType[],
					           listNode_t** returnNode) {
	HANDLE_EXCEPTIONS_BEGIN

	bool isField = wcscmp(fieldType, L"Field") == 0;
	bool isBookmark = wcscmp(fieldType, L"Bookmark") == 0;
	if(!isField && !isBookmark) {
		DIE(L"Field type not implemented");
	}
	
	prepareReadFieldCode(doc);
	
	listNode_t* fieldListStart = NULL;
	listNode_t* fieldListEnd = NULL;
	field_t* currentFields[3];
	CStoryRanges comStoryRanges = doc->comDoc.get_StoryRanges();

	if(isField) {
		IEnumVARIANT* fieldEnum[3];
		CRange comStoryRange;
		CFields comFields;
		LPUNKNOWN pUnk;

		// Get enumerators
		for(short i=0; i<3; i++) {
			try {
				comStoryRange = comStoryRanges.Item(i+1);
			} catch(COleDispatchException* e) {
				e->Delete();
				currentFields[i] = NULL;
				continue;
			}
			comFields = comStoryRange.get_Fields();

			// Generate enum
			pUnk = comFields.get__NewEnum();
			pUnk->QueryInterface(IID_IEnumVARIANT, (void **)&fieldEnum[i]);
			pUnk->Release();

			// Get first field
			ENSURE_OK(getNextFieldFromEnumerator(doc, fieldEnum[i], i, &currentFields[i]));
		}
		
		// Build field list
		bool noMoreFields = false;
		while(true) {
			short noteType;
			ENSURE_OK(addNextFieldToList(doc, currentFields, &fieldListStart, &fieldListEnd, &noteType, &noMoreFields));
			if(noMoreFields) break;
			// Get next field of noteType
			ENSURE_OK(getNextFieldFromEnumerator(doc, fieldEnum[noteType], noteType, &currentFields[noteType]));
		}
	} else {
		CRange comStoryRange;
		unsigned long bookmarkIndex[3] = {0, 0, 0};
		unsigned long bookmarkCount[3];
		CBookmarks comBookmarks[3];

		for(short i=0; i<3; i++) {
			// Get bookmarks collection for story
			try {
				comStoryRange = comStoryRanges.Item(i+1);
			} catch(COleDispatchException* e) {
				e->Delete();
				currentFields[i] = NULL;
				bookmarkCount[i] = 0;
				continue;
			}
			comBookmarks[i] = comStoryRange.get_Bookmarks();
			bookmarkCount[i] = comBookmarks[i].get_Count();
		
			// Get first bookmark
			ENSURE_OK(getNextBookmark(doc, comBookmarks[i], &bookmarkIndex[i], bookmarkCount[i], i, &currentFields[i]));
		}
		
		// Build field list
		bool noMoreFields = false;
		while(true) {
			short noteType;
			ENSURE_OK(addNextFieldToList(doc, currentFields, &fieldListStart, &fieldListEnd, &noteType, &noMoreFields));
			if(noMoreFields) break;
			// Get next bookmark of noteType
			ENSURE_OK(getNextBookmark(doc, comBookmarks[noteType], &bookmarkIndex[noteType], bookmarkCount[noteType],
				                      noteType, &currentFields[noteType]));
		}
	}
	
	if(fieldListStart) addValueToList(fieldListStart,
									  &(doc->allocatedFieldListsStart),
									  &(doc->allocatedFieldListsEnd));
	*returnNode = fieldListStart;

	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Add the next field of currentFields to the field list specified by fieldListStart and fieldListEnd. If no field remain,
// noMoreFields will be set to true.
statusCode addNextFieldToList(document_t* doc, field_t* currentFields[], listNode_t** fieldListStart, listNode_t** fieldListEnd,
                              short* noteType, bool* noMoreFields) {
	HANDLE_EXCEPTIONS_BEGIN
	bool isNextField = false;
	
	// Figure out which node should come next
	short noteTypeA;
	for(noteTypeA=0; noteTypeA<3; noteTypeA++) {
		// Check if there is a field.
		if(!currentFields[noteTypeA]) continue;
				
		// We have a field. For now, assume this will be the next field.
		isNextField = true;
				
		for(unsigned short noteTypeB = 0; noteTypeB<3; noteTypeB++) {
			if(noteTypeB == noteTypeA || !currentFields[noteTypeB]) {
				continue;
			}
					
			// If there is another field before this one, then obviously
			// this is not the next field.
			short returnValue;
			ENSURE_OK(compareFields(currentFields[noteTypeA], currentFields[noteTypeB], &returnValue));
			if(returnValue > 0) {
				isNextField = false;
				break;
			}
		}
				
		if(isNextField) {
			// Add as next field
			break;
		}
	}

	if(!isNextField) {
		// No next field
		*noMoreFields = true;
	} else {
		*noteType = noteTypeA;
		addValueToList(currentFields[noteTypeA], fieldListStart, fieldListEnd);
	}

	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Get the next field of a given type from an enumerator
statusCode getNextFieldFromEnumerator(document_t* doc, IEnumVARIANT* enumerator, short noteType, field_t** returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	VARIANT var;
	do {
		// Check that the enumerator has more fields to give
		if(enumerator->Next(1, &var, NULL) != S_OK) {
			*returnValue = NULL;
			return STATUS_OK;
		}
		// Get new field
		CField comField = var.pdispVal;
		ENSURE_OK(initField(doc, comField, noteType, false, returnValue))
		// Only return if this is a Zotero field
		if(*returnValue != NULL) return STATUS_OK;
	} while(*returnValue == NULL);

	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Get the next bookmark of a given type from the bookmarks collection
statusCode getNextBookmark(document_t* doc, CBookmarks comBookmarks, unsigned long* bookmarkIndex,
	                       const unsigned long bookmarkCount, short noteType, field_t** returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	*returnValue = NULL;
	do {
		// Check that this is not the next bookmark
		if(*bookmarkIndex >= bookmarkCount) {
			*returnValue = NULL;
			return STATUS_OK;
		}
		// Get the next bookmark
		CBookmark0 comBookmark = comBookmarks.Item((*bookmarkIndex)++ + 1);
		ENSURE_OK(initBookmark(doc, comBookmark, noteType, false, returnValue));
		// Only return if this is a Zotero field
		if(*returnValue != NULL) return STATUS_OK;
	} while(*returnValue == NULL);

	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Convert fields from one type to anoter
statusCode __stdcall convert(document_t *doc, field_t* fields[], unsigned long nFields,
				   const wchar_t toFieldType[], unsigned short noteType[]) {
	HANDLE_EXCEPTIONS_BEGIN
	setScreenUpdatingStatus(doc, false);
	for(unsigned long i=0; i<nFields; i++) {
		field_t* field = fields[i];
		bool isBookmark = field->comBookmark != NULL;
		bool convertToBookmark = !isBookmark && wcscmp(toFieldType, L"Bookmark") == 0;
		bool convertToField = isBookmark && wcscmp(toFieldType, L"Field") == 0;
		if(convertToBookmark || convertToField) {
			// Save and remove field code
			CRange comContentRange = field->comContentRange;
			wchar_t *oldCode = _wcsdup(field->code);
			
			// Remove code from old field
			ENSURE_OK(removeCode(field));

			// Create new field
			field_t* newField;
			ENSURE_OK(insertFieldRaw(doc, toFieldType, comContentRange, &newField));

			// Restore code
			ENSURE_OK(setCode(newField, oldCode));
			free(oldCode);

			// Use new field
			field = newField;
		}

		if(field->noteType != noteType[i]) {
			ENSURE_OK(convertToNoteType(field, noteType[i]));
		}
	}
	
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Set the "Bibliography" style according to the requirements of a specific style
statusCode __stdcall setBibliographyStyle(document_t *doc, long firstLineIndent, 
								          long bodyIndent, unsigned long lineSpacing,
								          unsigned long entrySpacing, long tabStops[],
								          unsigned long tabStopCount) {
	HANDLE_EXCEPTIONS_BEGIN
	// Get bibliography style
	CStyles styles = doc->comDoc.get_Styles();
	CStyle style;
	try {
		if(doc->wordVersion >= 12) {
			try {
				style = styles.Item(BIBLIOGRAPHY_STYLE_ENUM);
			} catch(COleDispatchException* e) {
				e->Delete();
				style = styles.Item(BIBLIOGRAPHY_STYLE_NAME);
			}
		} else {
			style = styles.Item(BIBLIOGRAPHY_STYLE_NAME);
		}
	} catch(COleDispatchException* e) {
		e->Delete();
		style = styles.Add(BIBLIOGRAPHY_STYLE_NAME,
			COleVariant((short) 1) /* wdStyleTypeParagraph */);
	}

	// Set formatting attributes
	CParagraphFormat paraFormat = style.get_ParagraphFormat();
	paraFormat.put_FirstLineIndent(((float) firstLineIndent)/20);
	paraFormat.put_LeftIndent(((float) bodyIndent)/20);
	paraFormat.put_LineSpacing(((float) lineSpacing)/20);
	paraFormat.put_SpaceAfter(((float) entrySpacing)/20);

	// Set tab stops
	CTabStops comTabStops = paraFormat.get_TabStops();
	comTabStops.ClearAll();
	for(unsigned long i=0; i<tabStopCount; i++) {
		comTabStops.Add(((float) tabStops[i])/20,
			COleVariant((short) 0) /* wdAlignTabLeft */,
			COleVariant((short) 0) /* wdTabLeaderSpaces */);
	}

    return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

statusCode __stdcall exportDocument(document_t *doc, const wchar_t fieldType[], const wchar_t importInstructions[]) {
	HANDLE_EXCEPTIONS_BEGIN

	// Export fields
	listNode_t *fields;
	ENSURE_OK(getFields(doc, fieldType, &fields));
	listNode_t *fieldsStart = fields;
	while (fields != NULL) {
		field_t *field = (field_t *)fields->value;
		ENSURE_OK(setText(field, field->code, false));
		ENSURE_OK(removeCode(field));
		CHyperlinks links = field->comContentRange.get_Hyperlinks();
		links.Add(field->comContentRange, COleVariant(IMPORT_LINK_URL), covOptional, covOptional, covOptional, covOptional);
		fields = fields->next;
	}
	freeFieldList(fieldsStart, false);

	// Document data
	CRange docDataRange = doc->comDoc.get_Content();
	wchar_t *docData;
	ENSURE_OK(getDocumentData(doc, &docData));

	docDataRange.InsertParagraphAfter();
	docDataRange.Collapse(0 /*wdCollapseEnd*/);
	docDataRange.InsertAfter(IMPORT_DOC_PREFS_PREFIX);
	docDataRange.InsertAfter(docData);
	CHyperlinks links = docDataRange.get_Hyperlinks();
	links.Add(docDataRange, COleVariant(IMPORT_LINK_URL), covOptional, covOptional, covOptional, covOptional);
	free(docData);

	// Import instructions
	docDataRange = doc->comDoc.get_Content();
	docDataRange.InsertParagraphBefore();
	docDataRange.InsertParagraphBefore();
	docDataRange.InsertBefore(importInstructions);
	// Export marker
	docDataRange.InsertParagraphBefore();
	docDataRange.InsertParagraphBefore();
	docDataRange.InsertBefore(EXPORTED_DOCUMENT_MARKER[0]);

	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

statusCode __stdcall importDocument(document_t *doc, const wchar_t fieldType[], bool *returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	*returnValue = false;

	CStoryRanges comStoryRanges = doc->comDoc.get_StoryRanges();
	CRange comStoryRange;
	// Main body, footnotes and endnotes (1 indexed!)
	for (long i = 1; i <= 3; i++) {
		try {
			comStoryRange = comStoryRanges.Item(i);
		}
		catch (COleDispatchException* e) {
			e->Delete();
			continue;
		}
		CHyperlinks comLinks = comStoryRange.get_Hyperlinks();
		long count = comLinks.get_Count();
		// 1 indexed!!!
		for (long j = count; j > 0; j--) {
			CHyperlink comLink = comLinks.Item(COleVariant(j));
			CRange comRange = comLink.get_Range();
			CString linkText = comRange.get_Text();
			if (linkText.Find(IMPORT_ITEM_PREFIX) == 0 || linkText.Find(IMPORT_BIBL_PREFIX) == 0) {
				field_t *field;
				ENSURE_OK(insertFieldRaw(doc, fieldType, comRange, &field));
				ENSURE_OK(setCode(field, linkText));
				if (i == 2) {
					setStyle(doc, &comRange, FOOTNOTE_STYLE_ENUM, FOOTNOTE_STYLE_NAME);
				}
				else if (i == 3) {
					setStyle(doc, &comRange, ENDNOTE_STYLE_ENUM, ENDNOTE_STYLE_NAME);
				}
			}
			else if (linkText.Find(IMPORT_DOC_PREFS_PREFIX) == 0) {
				*returnValue = true;
				linkText.Delete(0, lstrlen(IMPORT_DOC_PREFS_PREFIX));
				ENSURE_OK(setDocumentData(doc, linkText));
				comRange.put_Text(L"");
			}
		}
	}
	
	// Remove 3 paragraphs: export marker, empty paragraph, import instructions and another empty paragraph
	CRange range = doc->comDoc.get_Content();
	range.Collapse(1 /*wdCollapseStart*/);
	range.MoveEnd(4 /*wdParagraph*/, 4);
	range.put_Text(L"");

	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

statusCode __stdcall insertText(document_t *doc, const wchar_t htmlString[]) {
	HANDLE_EXCEPTIONS_BEGIN
	setScreenUpdatingStatus(doc, false);

	CSelection selection = doc->comWindow.get_Selection();
	CRange insertRange = selection.get_Range();
	if (doc->insertTextIntoNote && insertRange.get_StoryType() == 1) {
		// Due to the stupid way we're handling this by inserting a field first and then un-inserting
		// if text has to be inserted into the doc we have to re-insert a note here.
		if (doc->insertTextIntoNote == NOTE_FOOTNOTE) {
			CFootnotes notes = doc->comDoc.get_Footnotes();
			CFootnote note = notes.Add(insertRange, covOptional, covOptional);
			insertRange = note.get_Range();
		}
		else {
			CEndnotes notes = doc->comDoc.get_Endnotes();
			CEndnote note = notes.Add(insertRange, covOptional, covOptional);
			insertRange = note.get_Range();
		}
	}

	// Get current font info
	CFont0 comFont = selection.get_Font();
	CString fontName = comFont.get_Name();

	// Get a temp file
	char* utf8String;
	int nBytes;

	// Convert to UTF-8
	nBytes = WideCharToMultiByte(CP_UTF8, 0, htmlString, -1, NULL, 0, NULL, NULL);
	utf8String = new char[nBytes];
	WideCharToMultiByte(CP_UTF8, 0, htmlString, -1, utf8String, nBytes, NULL, NULL);

	// Open and write file
	DWORD nWritten;
	HANDLE tempFileHandle = getTemporaryFile();
	WriteFile(tempFileHandle, utf8String, nBytes - 1, &nWritten, NULL);
	SetEndOfFile(tempFileHandle);
	delete[] utf8String;

	// Read from file into range
	insertRange.put_Text(L"  ");
	CRange toDelete = insertRange.get_Duplicate();
	toDelete.Collapse(0);
	CRange comDupRange = insertRange.get_Duplicate();
	comDupRange.MoveEnd(1, -1);
	comDupRange.InsertFile(getTemporaryFilePath(), &covOptional, &covFalse, &covFalse, &covFalse);
	toDelete.MoveStart(1, -1);
	toDelete.put_Text(L"");
	
	// Put font back on
	comFont = insertRange.get_Font();
	comFont.put_Name(fontName);

	selection.put_Start(insertRange.get_End());
	selection.put_End(insertRange.get_End());

	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

statusCode __stdcall convertPlaceholdersToFields(document_t *doc, const wchar_t* placeholders[], const unsigned long nPlaceholders,
		const unsigned short noteType, const wchar_t fieldType[], listNode_t** returnNode) {
	HANDLE_EXCEPTIONS_BEGIN
	bool isField = wcscmp(fieldType, L"Field") == 0;
	bool isBookmark = wcscmp(fieldType, L"Bookmark") == 0;
	if (!isField && !isBookmark) {
		DIE(L"Field type not implemented");
	}

	listNode_t* fieldListStart = NULL;
	listNode_t* fieldListEnd = NULL;

	CStoryRanges comStoryRanges = doc->comDoc.get_StoryRanges();
	CRange comStoryRange;
	// Main body, footnotes and endnotes (1 indexed!)
	for (long i = 1; i <= 3; i++) {
		try {
			comStoryRange = comStoryRanges.Item(i);
		}
		catch (COleDispatchException* e) {
			e->Delete();
			continue;
		}
		CHyperlinks comLinks = comStoryRange.get_Hyperlinks();
		long count = comLinks.get_Count();
		// 1 indexed!!!
		for (long j = count; j > 0; j--) {
			CHyperlink comLink = comLinks.Item(COleVariant(j));
			CString linkUrl = comLink.get_Address();
			CString placeholderID = linkUrl.Right(LINK_PLACEHOLDER_LENGTH);
			const wchar_t *code = NULL;
			for (long k = 0; k < nPlaceholders; k++) {
				if (placeholderID.Find(placeholders[k]) != 0) {
					continue;
				}
				code = L"TEMP";
				break;
			}
			if (code == NULL) {
				continue;
			}
			CRange insertRange = comLink.get_Range();
			field_t *newField;

			if (insertRange.get_StoryType() == 1) {
				insertRange.put_Text(L"");
				// If inserting a note citation in the main story, we need to make a new note
				if (noteType == NOTE_FOOTNOTE) {
					CFootnotes notes = doc->comDoc.get_Footnotes();
					CFootnote note = notes.Add(insertRange, covOptional, covOptional);
					// Move cursor back to main text
					CRange referenceRange = note.get_Reference();
					CRange dupRange = referenceRange.get_Duplicate();
					dupRange.Collapse(0 /*wdCollapseEnd*/);
					dupRange.Select();
					// Now inserting field into note
					insertRange = note.get_Range();
				}
				else if (noteType == NOTE_ENDNOTE) {
					CEndnotes notes = doc->comDoc.get_Endnotes();
					CEndnote note = notes.Add(insertRange, covOptional, covOptional);
					// Move cursor back to main text
					CRange referenceRange = note.get_Reference();
					CRange dupRange = referenceRange.get_Duplicate();
					dupRange.Collapse(0 /*wdCollapseEnd*/);
					dupRange.Select();
					// Now inserting field into note
					insertRange = note.get_Range();
				}
			}

			ENSURE_OK(insertFieldRaw(doc, fieldType, insertRange, &newField));
			if (newField->code) free(newField->code);
			ENSURE_OK(setCode(newField, code));
			addValueToList(newField, &fieldListStart, &fieldListEnd);
		}
	}
	*returnNode = fieldListStart;

	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

statusCode __stdcall cleanup(document_t *doc) {
	HANDLE_EXCEPTIONS_BEGIN
	if(doc->restoreShowRevisions) {
		if (doc->wordVersion >= 15) {
			CView0 comView = doc->comWindow.get_View();
			CRevisionsFilter revisionsFilter = comView.get_RevisionsFilter();
			revisionsFilter.put_Markup(doc->restoreRevisionsMarkup);
		}
		else {
			doc->comDoc.put_ShowRevisions(true);
		}
		doc->statusShowRevisions = true;
	}
	setScreenUpdatingStatus(doc, true);
	deleteTemporaryFile();
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

statusCode __stdcall complete(document_t *doc) {
	HANDLE_EXCEPTIONS_BEGIN
	if (doc->wordVersion >= 14) {
		CUndoRecord undoRecord = doc->comApp.get_UndoRecord();
		undoRecord.EndCustomRecord();
	}
	if (doc->restoreTrackChanges) {
		doc->comDoc.put_TrackRevisions(true);
		doc->restoreTrackChanges = false;
	}
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Gets a property from the document
statusCode getProperty(document_t *doc, CString propertyName,
					   CString* returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	int i = 0;
	CString comPropertyName;
	CCustomProperty comProperty;
	
	// Get all existing properties corresponding to this propertyName
	while(true) {
		i++;
		comPropertyName.Format(_T("%s_%d"), propertyName, i);

		try {
			comProperty = doc->comProperties.Item(comPropertyName);
			*returnValue += comProperty.get_Value();
		} catch(COleDispatchException* e) {
			e->Delete();
			break;
		}
	}
	
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Sets the value of a property stored in the document
statusCode setProperty(document_t *doc, CString propertyName, CString propertyValue) {
	HANDLE_EXCEPTIONS_BEGIN
	long i = 0;
	long propertyValueLength = propertyValue.GetLength();
	CString subpropertyName, subpropertyValue;
	CCustomProperty comProperty;
	
	// Change existing properties
	while(true) {
		subpropertyName.Format(_T("%s_%i"), propertyName, (i+1));

		if(propertyValueLength > (i*MAX_PROPERTY_LENGTH)) {
			// if more to add, add it
			subpropertyValue = propertyValue.Mid(i*MAX_PROPERTY_LENGTH, MAX_PROPERTY_LENGTH);
			try {
				comProperty = doc->comProperties.Item(subpropertyName);
				comProperty.put_Value(subpropertyValue);
			} catch(COleDispatchException* e) {
				e->Delete();
				doc->comProperties.Add(subpropertyName, false, 4, subpropertyValue);
			}
		} else {
			// If no more to add, delete extraneous properties until there are no more
			try {
				comProperty = doc->comProperties.Item(subpropertyName);
				comProperty.Delete();
			} catch(COleDispatchException* e) {
				e->Delete();
				break;
			}
		}

		i++;
	}

	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Make a new field at insertRange, without setting the field code
statusCode insertFieldRaw(document_t *doc, const wchar_t fieldType[],
						  CRange comWhere, field_t** returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	if(wcscmp(fieldType, L"Field") == 0) {
		CFields comFields = doc->comDoc.get_Fields();
		// Creating the field as a wdFieldQuote type, since creating as a wdFieldAddin field screws up the ranges
		CField field = comFields.Add(comWhere, COleVariant((short) 35), COleVariant(FIELD_PLACEHOLDER), COleVariant((short)true));
		return initField(doc, field, -1, true, returnValue);
	}
	
	if(wcscmp(fieldType, L"Bookmark") == 0) {
		CBookmarks comBookmarks = doc->comDoc.get_Bookmarks();
		comWhere.put_Text(FIELD_PLACEHOLDER);
		CString bookmarkName;
		bookmarkName.Format(L"%s_%s", BOOKMARK_REFERENCE_PROPERTY, generateRandomString(12));
		CBookmark0 comBookmark = comBookmarks.Add(bookmarkName, comWhere);
		return initBookmark(doc, comBookmark, -1, true, returnValue);
	}
	
	DIE(L"Field type not implemented");
	HANDLE_EXCEPTIONS_END
}

// Turns screenUpdating on or off
statusCode setScreenUpdatingStatus(document_t* doc, bool status) {
	HANDLE_EXCEPTIONS_BEGIN
	if(status != doc->statusScreenUpdating) {
		doc->comApp.put_ScreenUpdating(status);
		doc->statusScreenUpdating = status;
	}
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Adds a field to a field list
void addValueToList(void* value, listNode_t** listStart,
					listNode_t** listEnd) {
	listNode_t* node = (listNode_t*) malloc(sizeof(listNode_t));
	node->value = value;
	node->next = NULL;
	
	if(*listEnd) {
		(*listEnd)->next = node;
		*listEnd = node;
	} else {
		*listStart = *listEnd = node;
	}
}