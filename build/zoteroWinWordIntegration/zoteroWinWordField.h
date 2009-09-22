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

#include "stdafx.h"
#include "resource.h"

#define ZOTEROWINWORDFIELD_CONTRACTID "@zotero.org/Zotero/integration/field?agent=WinWord&type=Field;1"
#define ZOTEROWINWORDFIELD_CLASSNAME "WinWord Field"
#define ZOTEROWINWORDFIELD_CID  { 0x5aff3b25, 0x7dbb, 0x456c, 0x92, 0xd8, 0xac, 0xdc, 0xa1, 0x35, 0x65, 0xd7 }

#define ZOTEROWINWORDFIELDENUMERATOR_CONTRACTID "@zotero.org/Zotero/integration/enumerator?agent=WinWord&type=Field;1"
#define ZOTEROWINWORDFIELDENUMERATOR_CLASSNAME "WinWord nsISimpleEnumerator"
#define ZOTEROWINWORDFIELDENUMERATOR_CID  { 0x78d55e39, 0xc015, 0x49a0, 0x9a, 0x87, 0x9c, 0x8f, 0x62, 0x36, 0xa2, 0x56 }

/* Header file */
class zoteroWinWordField : public zoteroIntegrationField
{
public:
	NS_DECL_ISUPPORTS
	NS_DECL_ZOTEROINTEGRATIONFIELD

	long offset1;
	long offset2;
	CRange comTextRange;

	zoteroWinWordField();
	zoteroWinWordField(zoteroWinWordDocument *aDoc, CField field);
	zoteroWinWordField(zoteroWinWordDocument *aDoc, CField field, CRange codeRange, CString code);
	// gah! this destructor has to be virtual, or else our CString doesn't get deallocated.
	virtual ~zoteroWinWordField();

	void convertToNoteType(PRUint16 noteType);
	void loadOffsets();
	virtual CRange getFieldRange();
	bool operator<(zoteroWinWordField &other);

private:
	CField comField;
	CRange comCodeRange;
	CString rawCode;

	bool isWholeNote();

protected:
	CFootnote comFootnote;
	CEndnote comEndnote;
	zoteroWinWordDocument *doc;

	virtual void init(bool needCode);
	virtual void loadFromRange(CRange comRange);
	virtual void deleteField();
};

class zoteroWinWordEnumerator : public nsISimpleEnumerator
{
public:
	NS_DECL_ISUPPORTS
	NS_DECL_NSISIMPLEENUMERATOR
	~zoteroWinWordEnumerator();

protected:
	zoteroWinWordDocument *doc;
	zoteroWinWordField *fieldItem[3];
	virtual void fetchNextItem(short i) = 0;
};

class zoteroWinWordFieldEnumerator : public zoteroWinWordEnumerator
{
public:
	zoteroWinWordFieldEnumerator();
	zoteroWinWordFieldEnumerator(zoteroWinWordDocument *aDoc);

private:
	IEnumVARIANT *fieldEnum[3];
	~zoteroWinWordFieldEnumerator();

protected:
	void fetchNextItem(short i);
};