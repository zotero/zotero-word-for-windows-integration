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

#ifndef _ZOTEROINTEGRATION_H_
#define _ZOTEROINTEGRATION_H_
#include "zoteroIntegration.h"
#endif

#pragma once
#include "stdafx.h"
#include "resource.h"

#define ZOTEROWINWORDDOCUMENT_CONTRACTID "@zotero.org/Zotero/integration/document?agent=WinWord;1"
#define ZOTEROWINWORDDOCUMENT_CLASSNAME "WinWord Document"
#define ZOTEROWINWORDDOCUMENT_CID  { 0x295b47af, 0xa084, 0x41b1, 0xb2, 0x6c, 0xb0, 0xce, 0xfa, 0x3, 0x92, 0x46 }

const CString FIELD_PREFIX = _T(" ADDIN ZOTERO_");
const CString BACKUP_FIELD_PREFIX = _T(" ADDIN CITE_");
const CString BOOKMARK_PREFIX = _T("ZOTERO_");
const CString BACKUP_BOOKMARK_PREFIX = _T("CITE_");
const CString PREFS_PROPERTY = _T("ZOTERO_PREF");
const CString BACKUP_PREFS_PROPERTY = _T("CITE_PREF");
const CString BOOKMARK_REFERENCE_PROPERTY = _T("ZOTERO_BREF_");
const CString BACKUP_BOOKMARK_REFERENCE_PROPERTY = _T("CITE_BREF_");
const CString FIELD_PLACEHOLDER = _T("{Citation}");
#define MAX_PROPERTY_LENGTH 255

/* Header file */
class zoteroWinWordDocument : public zoteroIntegrationDocument
{
public:
	NS_DECL_ISUPPORTS
	NS_DECL_ZOTEROINTEGRATIONDOCUMENT

	CDocument0 comDoc;

	zoteroWinWordDocument();
	zoteroWinWordDocument(const PRUnichar *);
	CString getProperty(CString propertyName);
	void setProperty(CString propertyName, CString propertyValue);
	CString getRandomString(int length);
	nsresult makeNewField(const char *fieldType, CRange insertRange, zoteroIntegrationField **_retval);
	void setScreenUpdatingStatus(bool status);
	void initFromActiveObject();

private:
	CApplication comApp;
	CCustomProperties comProperties;
	bool currentScreenUpdatingStatus;
	void initFilter();

	~zoteroWinWordDocument();

protected:
	/* additional members */
};