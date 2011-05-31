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

#pragma once
#include "stdafx.h"
#include "resource.h"
#ifndef __gen_nsIException_h__
#include "nsIException.h"
#endif

class zoteroException : public nsIException
{
public:
	NS_DECL_ISUPPORTS
	NS_DECL_NSIEXCEPTION
	
	zoteroException();
	zoteroException(char *message, char *aFunction, char *aFilename);
	zoteroException(CException *exception, char *aFunction, char *aFilename);
	void report();
	char *message;
	char *function;
	char *filename;
	uint errorID;
	bool freeMessage;
private:
	~zoteroException();
};

#define ZOTERO_EXCEPTION_CATCHER_START \
	try {

#define ZOTERO_EXCEPTION_CATCHER_END \
	} catch(CException *exception) { \
		zoteroException *zException = new zoteroException(exception, __FUNCTION__, __FILE__); \
		zException->report(); \
		return NS_ERROR_FAILURE; \
	} catch(...) { \
		return NS_ERROR_FAILURE; \
	}

#define ZOTERO_THROW_EXCEPTION(string) \
	zoteroException *zException = new zoteroException(string, __FUNCTION__, __FILE__); \
	zException->report();