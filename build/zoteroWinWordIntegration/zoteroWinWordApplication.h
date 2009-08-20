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

#define ZOTEROWINWORDAPPLICATION_CONTRACTID "@zotero.org/Zotero/integration/application?agent=WinWord;1"
#define ZOTEROWINWORDAPPLICATION_CLASSNAME "WinWord Application"
#define ZOTEROWINWORDAPPLICATION_CID  { 0xd10e90bf, 0x773e, 0x4be4, 0x95, 0x92, 0x9c, 0x13, 0x88, 0x29, 0xb8, 0xca }

/* Header file */
class zoteroWinWordApplication : public zoteroIntegrationApplication
{
public:
	NS_DECL_ISUPPORTS
	NS_DECL_ZOTEROINTEGRATIONAPPLICATION
	
	zoteroWinWordApplication();

private:
	~zoteroWinWordApplication();

protected:
	/* additional members */
};