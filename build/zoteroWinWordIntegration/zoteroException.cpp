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

#include "zoteroException.h"

#ifndef nsCOMPtr_h___
#include "nsCOMPtr.h"
#endif
#include "zoteroWinWordIntegrationErrorHandler.h"
#include "nsServiceManagerUtils.h"
//#include "nsComponentManagerUtils.h"
//#include "nsIConsoleService.h"
//#include "nsIScriptError.h"
//
//zoteroException::zoteroException()
//{
//	/* member initializers and constructor code */
//	errorID = 0;
//	freeMessage = false;
//}

zoteroException::zoteroException(char *aMessage, char *aFunction, char *aFilename)
{
	errorID = 0;
	message = aMessage;
	function = aFunction;
	filename = aFilename;
	freeMessage = false;
}

zoteroException::zoteroException(CException *exception, char *aFunction, char *aFilename)
{
	errorID = 0;
	wchar_t aMessage[512];
	if(exception->GetErrorMessage(aMessage, 512, &errorID)) {
		size_t returnValue;
		message = new char[512]();
		wcstombs_s(&returnValue, message, 512, aMessage, _TRUNCATE);
		freeMessage = true;
	} else {
		message = NULL;
		freeMessage = false;
	}
	exception->Delete();
	function = aFunction;
	filename = aFilename;
}

zoteroException::~zoteroException()
{
	if(freeMessage) {
		delete[] message;
	}
}

// Parses a CException, returning the error message to XPCOM
void zoteroException::report() {
	char *errorMessage = (char*) NS_Alloc(2048);
	nsCOMPtr<zoteroWinWordIntegrationErrorHandler> errorHandler = do_GetService("@zotero.org/Zotero/integration/winWordErrorHandler;1");
	_snprintf_s(errorMessage, 2048, _TRUNCATE, "%s  code: \"%d\"  function: \"%s\"  location: \"%s\"", message, errorID, function, filename);
	errorHandler->ThrowError(errorMessage);
}