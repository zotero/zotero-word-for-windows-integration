/*
	***** BEGIN LICENSE BLOCK *****
	
	Copyright (c) 2009-2012  Zotero
	                         Center for History and New Media
						     George Mason University, Fairfax, Virginia, USA

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

#include "zoteroWinWordIntegration.h"

CString* lastErrorString;

// Checks if an error has occurred
bool errorHasOccurred(void) {
	return lastErrorString != NULL;
}

// Manually throws an exception up to JS
void throwError(const wchar_t error[], const char function[], const char file[],
				unsigned int line) {
	if(!lastErrorString) lastErrorString = new CString();

	if(line == 0) {
		lastErrorString->Format(L"%s [%S:%S]", error, function, file);
	} else {
		lastErrorString->Format(L"%s [%S:%S:%d]", error, function, file, line);
	}
}

// Converts a CException to a JS exception
void throwError(const CException* error, const char function[], const char file[]) {
	wchar_t errorString[512];
	error->GetErrorMessage(errorString, 512);
	throwError(errorString, function, file, 0);
}

// Clears the last error encountered
void __stdcall clearError(void) {
	delete lastErrorString;
	lastErrorString = NULL;
}

// Gets the last error encountered
const wchar_t* __stdcall getError(void) {
	return *lastErrorString;
}

HANDLE tempFile = NULL;
wchar_t tempFilePath[MAX_PATH+1];

// Gets a FILE for the temporary file, truncating it to zero length
HANDLE getTemporaryFile(void) {
	if(tempFile == NULL) {
		wchar_t tempPath[MAX_PATH+1];
		GetTempPath(MAX_PATH, tempPath);
		GetTempFileName(tempPath, L"ZOTERO", 0, tempFilePath);
		tempFile = CreateFile(tempFilePath, GENERIC_WRITE, FILE_SHARE_READ, NULL,
			CREATE_ALWAYS, FILE_ATTRIBUTE_TEMPORARY, NULL);
	} else {
		SetFilePointer(tempFile, 0, NULL, FILE_BEGIN);
	}
	return tempFile;
}

// Deletes the temp file
void deleteTemporaryFile(void) {
	if(tempFile == NULL) return;
	CloseHandle(tempFile);
	tempFile = NULL;
	DeleteFile(tempFilePath);
}

// Gets an NSString representing the temp file
wchar_t* getTemporaryFilePath(void) {
	return tempFilePath;
}

// Frees a pointer
void __stdcall freeData(void* ptr) {
	free(ptr);
}

// Generates a random string
CString generateRandomString(unsigned int length) {
	// seed random number generator
	static bool rndGeneratorSeeded = false;
	if(rndGeneratorSeeded == false) {
		srand((unsigned)time(0));
		rndGeneratorSeeded = true;
	}
	
	// generator random string of desired length
	CString randString = L"";
	static const CString characters = L"abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
	for(unsigned int i=0; i<length; i++) {
		randString += characters.GetAt(rand() % 62);
	}
	return randString;
}

void setStyle(document_t *doc, CRange *range, long styleEnum, CString styleName) {
	// Set style on range
	try {
		if (doc->wordVersion >= 12) {
			// In Word 2007+, we should use WdBuiltinStyle enum
			// instead of the string name to reference the bibliography style.
			try {
				range->put_Style(styleEnum);
			}
			catch (...) {
				range->put_Style(styleName);
			}
		}
		else {
			range->put_Style(styleName);
		}
	}
	catch (...) {
		// This is probably not necessary, but it's better than throwing, and I can't test all
		// Word versions at this time
	}
}