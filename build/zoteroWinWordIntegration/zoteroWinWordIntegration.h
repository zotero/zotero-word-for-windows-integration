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

	***** END LICENSE BLOCK *****
 */
// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently, but
// are changed infrequently

#pragma once
#pragma comment(lib, "mpr.lib")

#include "stdafx.h"
#include "CApplication.h"
#include "CBookmark0.h"
#include "CBookmarks.h"
#include "CCustomProperties.h"
#include "CCustomProperty.h"
#include "CDocument0.h"
#include "CEndnote.h"
#include "CEndnotes.h"
#include "CFootnote.h"
#include "CFootnotes.h"
#include "CHyperlink.h"
#include "CHyperlinks.h"
#include "CField.h"
#include "CFields.h"
#include "CFont0.h"
#include "CParagraphFormat.h"
#include "CRange.h"
#include "CSelection.h"
#include "CStoryRanges.h"
#include "CStyle.h"
#include "CStyles.h"
#include "CTabStop.h"
#include "CTabStops.h"
#include "CUndoRecord.h"
#include "CView0.h"
#include "CWindow0.h"

enum STATUS {
	STATUS_OK = 0,
	STATUS_EXCEPTION = 1,
	STATUS_EXCEPTION_ALREADY_DISPLAYED = 2
};

enum DIALOG_ICON {
	DIALOG_ICON_STOP = 0,
	DIALOG_ICON_NOTICE = 1,
	DIALOG_ICON_CAUTION = 2
};

enum DIALOG_BUTTONS {
	DIALOG_BUTTONS_OK = 0,
	DIALOG_BUTTONS_OK_CANCEL = 1,
	DIALOG_BUTTONS_YES_NO = 2,
	DIALOG_BUTTONS_YES_NO_CANCEL = 3
};

enum NOTE_TYPE {
	NOTE_FOOTNOTE = 1,
	NOTE_ENDNOTE = 2
};

#define MAX_PROPERTY_LENGTH 255
#define FIELD_PLACEHOLDER L"{Citation}"
#define BOOKMARK_REFERENCE_PROPERTY L"ZOTERO_BREF"
#define UNDO_RECORD_NAME L"Zotero Command"
#define RTF_TEMP_BOOKMARK L"ZOTERO_TEMP_BOOKMARK"
#define PREFS_PROPERTY L"ZOTERO_PREF"
#define BOOKMARK_PREFIX L"ZOTERO_"
#define BIBLIOGRAPHY_STYLE_NAME L"Bibliography"
#define BIBLIOGRAPHY_STYLE_ENUM -266 /**WdBuiltinStyle.wdStyleBibliography**/
#define FOOTNOTE_STYLE_NAME L"Footnote Text"
#define FOOTNOTE_STYLE_ENUM -30 /**WdBuiltinStyle.wdStyleFootnoteText**/
#define ENDNOTE_STYLE_NAME L"Endnote Text"
#define ENDNOTE_STYLE_ENUM -44 /**WdBuiltinStyle.wdStyleEndnoteText**/

// Returns expr if expr is non-zero.
#define ENSURE_OK(expr) \
{ \
statusCode statusToEnsure = expr; \
if(statusToEnsure) return statusToEnsure; \
}

// Sets an error code doc and then returns STATUS_EXCEPTION.
#define DIE(err) \
{\
	throwError(err, __FUNCTION__, __FILE__, __LINE__);\
	return STATUS_EXCEPTION;\
}

#define IGNORING_SB_ERRORS_BEGIN setErrorMonitor(false);
#define IGNORING_SB_ERRORS_END setErrorMonitor(true);

#define HANDLE_EXCEPTIONS_BEGIN \
try {

#define HANDLE_EXCEPTIONS_END \
} catch(CException *exception) { \
	throwError(exception, __FUNCTION__, __FILE__); \
	exception->Delete(); \
	return STATUS_EXCEPTION; \
} catch(...) { \
	throwError(L"An unhandled exception occured.", __FUNCTION__, __FILE__, 0);\
	return STATUS_EXCEPTION; \
}

typedef struct ListNode {
	void* value;
	struct ListNode* next;
} listNode_t;

typedef struct Document {
	CApplication comApp;
	CDocument0 comDoc;
	CCustomProperties comProperties;
	CWindow0 comWindow;

	long wordVersion;
	bool restoreShowRevisions;
	bool statusShowRevisions;
	bool statusScreenUpdating;
	bool restoreTrackChanges;
	
	listNode_t* allocatedFieldsStart;
	listNode_t* allocatedFieldsEnd;
	listNode_t* allocatedFieldListsStart;
	listNode_t* allocatedFieldListsEnd;
} document_t;

typedef struct Field {
	// The field code
	wchar_t* code;
	
	// The field text
	wchar_t* text;
	
	// The note type (0, NOTE_FOOTNOTE, or NOTE_ENDNOTE)
	short noteType;
	
	// The bookmark name
	wchar_t* bookmarkName;
	
	// The location of this field relative to the start of the main body text.
	// For a footnote, this would be the position of the superscripted note
	// reference.
	long textLocation;
	
	// The location of this field relative to the start of the footnote/endnote
	// story.
	long noteLocation;
	
	// Only one of these will be set
	CField comField;
	CBookmark0 comBookmark;
	
	// The corresponding document
	document_t* doc;
	
	// The range corresponding to the field code, for a field
	CRange comCodeRange;
	
	// The range corresponding to the content of a field
	CRange comContentRange;

	// The footnote in which this field resides.
	CFootnote comFootnote;

	// The endnote in which this field resides.
	CEndnote comEndnote;
} field_t;

typedef unsigned short statusCode;

// utilities.cpp
extern "C" {
	__declspec(dllexport) const wchar_t* __stdcall getError(void);
	__declspec(dllexport) void __stdcall clearError(void);
	__declspec(dllexport) void __stdcall freeData(void* data);
}

bool errorHasOccurred(void);
void throwError(const wchar_t error[], const char function[], const char file[],
				unsigned int line);
void throwError(const CException* error, const char function[], const char file[]);

HANDLE getTemporaryFile(void);
void deleteTemporaryFile(void);
wchar_t* getTemporaryFilePath(void);

CString generateRandomString(unsigned int length);

void setStyle(document_t *doc, CRange *range, long styleEnum, CString styleName);

// document.cpp
extern "C" {
	__declspec(dllexport) statusCode __stdcall getDocument(const wchar_t documentName[], document_t** returnValue);
	__declspec(dllexport) void __stdcall freeDocument(document_t *doc);
	__declspec(dllexport) statusCode __stdcall displayAlert(document_t *doc, const wchar_t dialogText[],
							                                unsigned short icon, unsigned short buttons,
							                                unsigned short* returnValue);
	__declspec(dllexport) statusCode __stdcall canInsertField(document_t *doc, const wchar_t fieldType[],
							                                  bool* returnValue);
	__declspec(dllexport) statusCode __stdcall cursorInField(document_t *doc, const wchar_t fieldType[],
							                                 field_t** returnValue);
	__declspec(dllexport) statusCode __stdcall getDocumentData(document_t *doc, wchar_t **returnValue);
	__declspec(dllexport) statusCode __stdcall setDocumentData(document_t *doc, const wchar_t documentData[]);
	__declspec(dllexport) statusCode __stdcall insertField(document_t *doc, const wchar_t fieldType[],
						                                   unsigned short noteType, field_t **returnValue);
	__declspec(dllexport) statusCode __stdcall getFields(document_t *doc, const wchar_t fieldType[],
						                                 listNode_t** returnNode);
	__declspec(dllexport) statusCode __stdcall convert(document_t *doc, field_t* fields[], unsigned long nFields,
					                                   const wchar_t toFieldType[], unsigned short noteType[]);
	__declspec(dllexport) statusCode __stdcall setBibliographyStyle(document_t *doc, long firstLineIndent, 
									                                long bodyIndent, unsigned long lineSpacing,
	                                                                unsigned long entrySpacing, long tabStops[],
									                                unsigned long tabStopCount);
	__declspec(dllexport) statusCode __stdcall exportDocument(document_t *doc, const wchar_t fieldType[], const wchar_t importInstructions[]);
	__declspec(dllexport) statusCode __stdcall importDocument(document_t *doc, const wchar_t fieldType[], bool* returnValue);
	__declspec(dllexport) statusCode __stdcall cleanup(document_t *doc);
	__declspec(dllexport) statusCode __stdcall complete(document_t *doc);
}

statusCode getProperty(document_t *doc, CString propertyName,
					   CString* returnValue);
statusCode setProperty(document_t *doc, CString propertyName,
					   CString propertyValue);
statusCode prepareReadFieldCode(document_t *doc);
statusCode setScreenUpdatingStatus(document_t* doc, bool status);
statusCode insertFieldRaw(document_t *doc, const wchar_t fieldType[],
						  CRange comWhere, field_t** returnValue);
void addValueToList(void* value, listNode_t** listStart, listNode_t** listEnd);

// field.cpp
extern "C" {
	__declspec(dllexport) void __stdcall freeField(field_t* field);
	__declspec(dllexport) statusCode __stdcall deleteField(field_t* field);
	__declspec(dllexport) statusCode __stdcall removeCode(field_t* field);
	__declspec(dllexport) statusCode __stdcall selectField(field_t* field);
	__declspec(dllexport) statusCode __stdcall setText(field_t* field, const wchar_t string[], bool isRich);
	__declspec(dllexport) statusCode __stdcall getText(field_t* field, wchar_t** returnValue);
	__declspec(dllexport) statusCode __stdcall setCode(field_t *field, const wchar_t code[]);
	__declspec(dllexport) statusCode __stdcall getNoteIndex(field_t* field, unsigned long *returnValue);
}

statusCode initField(document_t *doc, CField comField, short noteType,
					 bool ignoreCode, field_t **returnValue);
statusCode initBookmark(document_t *doc, CBookmark0 comBookmark, short noteType,
						bool ignoreCode, field_t **returnValue);
statusCode compareFields(field_t* a, field_t* b, short *returnValue);
statusCode convertToNoteType(field_t* field, short noteType);