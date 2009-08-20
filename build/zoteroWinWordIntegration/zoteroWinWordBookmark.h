/*
    ***** BEGIN LICENSE BLOCK *****
	
	Copyright (c) 2009  Center for History and New Media
						George Mason University, Fairfax, Virginia, USA
						http://chnm.gmu.edu
	
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

#include "stdafx.h"
#include "resource.h"

#define ZOTEROWINWORDBOOKMARK_CONTRACTID "@zotero.org/Zotero/integration/field?agent=WinWord&type=Bookmark;1"
#define ZOTEROWINWORDBOOKMARK_CLASSNAME "WinWord Bookmark"
#define ZOTEROWINWORDBOOKMARK_CID  { 0xc073819d, 0x5290, 0x45c2, 0x90, 0x4d, 0xde, 0x93, 0x55, 0xd1, 0x33, 0xc6 }
#define MAX_BOOKMARK_LENGTH 50

#define ZOTEROWINWORDBOOKMARKENUMERATOR_CONTRACTID "@zotero.org/Zotero/integration/enumerator?agent=WinWord&type=Bookmark;1"
#define ZOTEROWINWORDBOOKMARKENUMERATOR_CLASSNAME "WinWord nsISimpleEnumerator"
#define ZOTEROWINWORDBOOKMARKENUMERATOR_CID  { 0xf6d32f2f, 0xdef, 0x43c2, 0xb3, 0x7, 0xab, 0x5d, 0xe9, 0x82, 0x31, 0xbf }

/* Header file */
class zoteroWinWordBookmark : public zoteroWinWordField
{
public:
	zoteroWinWordBookmark();
	zoteroWinWordBookmark(zoteroWinWordDocument *aDoc, CBookmark0 field);

	NS_IMETHODIMP RemoveCode();
	NS_IMETHODIMP SetCode(const PRUnichar *code);
	NS_IMETHODIMP GetCode(PRUnichar **_retval);
	NS_IMETHODIMP SetText(const PRUnichar *text, PRBool isRich);
	CRange getFieldRange();

private:
	CBookmark0 comBookmark;
	CString bookmarkName;

protected:
	void init();
	void loadFromRange(CRange comRange);
	void deleteField();
};

class zoteroWinWordBookmarkEnumerator : public zoteroWinWordEnumerator
{
public:
	zoteroWinWordBookmarkEnumerator();
	zoteroWinWordBookmarkEnumerator(zoteroWinWordDocument *aDoc);

private:
	CBookmarks comBookmarks[3];
	long bookmarkCount[3];
	long bookmarkIndex[3];
	
	zoteroWinWordField* getField(LPDISPATCH comField);

protected:
	void fetchNextItem(short i);
};