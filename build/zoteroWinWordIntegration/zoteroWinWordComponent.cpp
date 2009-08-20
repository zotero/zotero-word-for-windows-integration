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

#include "nsXPCOM.h"
#include "nsIGenericFactory.h"
	
/**
 * Components to be registered
 */
#include "zoteroWinWordApplication.h"
#include "zoteroWinWordDocument.h"
#include "zoteroWinWordField.h"
#include "zoteroWinWordBookmark.h"
	
NS_GENERIC_FACTORY_CONSTRUCTOR(zoteroWinWordApplication)
NS_GENERIC_FACTORY_CONSTRUCTOR(zoteroWinWordDocument)
NS_GENERIC_FACTORY_CONSTRUCTOR(zoteroWinWordFieldEnumerator)
NS_GENERIC_FACTORY_CONSTRUCTOR(zoteroWinWordBookmarkEnumerator)
NS_GENERIC_FACTORY_CONSTRUCTOR(zoteroWinWordField)
NS_GENERIC_FACTORY_CONSTRUCTOR(zoteroWinWordBookmark)
	
//----------------------------------------------------------
	
static const nsModuleComponentInfo components[] =
{
	{
		ZOTEROWINWORDAPPLICATION_CLASSNAME,
		ZOTEROWINWORDAPPLICATION_CID,
		ZOTEROWINWORDAPPLICATION_CONTRACTID,
		zoteroWinWordApplicationConstructor
	},
	{
		ZOTEROWINWORDDOCUMENT_CLASSNAME,
		ZOTEROWINWORDDOCUMENT_CID,
		ZOTEROWINWORDDOCUMENT_CONTRACTID,
		zoteroWinWordDocumentConstructor
	},
	{
		ZOTEROWINWORDFIELDENUMERATOR_CLASSNAME,
		ZOTEROWINWORDFIELDENUMERATOR_CID,
		ZOTEROWINWORDFIELDENUMERATOR_CONTRACTID,
		zoteroWinWordFieldEnumeratorConstructor
	},
	{
		ZOTEROWINWORDBOOKMARKENUMERATOR_CLASSNAME,
		ZOTEROWINWORDBOOKMARKENUMERATOR_CID,
		ZOTEROWINWORDBOOKMARKENUMERATOR_CONTRACTID,
		zoteroWinWordBookmarkEnumeratorConstructor
	},
	{
		ZOTEROWINWORDFIELD_CLASSNAME,
		ZOTEROWINWORDFIELD_CID,
		ZOTEROWINWORDFIELD_CONTRACTID,
		zoteroWinWordFieldConstructor
	},
	{
		ZOTEROWINWORDBOOKMARK_CLASSNAME,
		ZOTEROWINWORDBOOKMARK_CID,
		ZOTEROWINWORDBOOKMARK_CONTRACTID,
		zoteroWinWordBookmarkConstructor
	}
};
	
NS_IMPL_NSGETMODULE(MyExtension, components)
