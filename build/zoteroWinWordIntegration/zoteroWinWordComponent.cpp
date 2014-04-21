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

/**
 * Components to be registered
 */
#include "zoteroWinWordApplication.h"
#include "zoteroWinWordDocument.h"
#include "zoteroWinWordField.h"
#include "zoteroWinWordBookmark.h"

#include "mozilla/ModuleUtils.h"
#include "nsIClassInfoImpl.h"

NS_GENERIC_FACTORY_CONSTRUCTOR(zoteroWinWordApplication)
NS_GENERIC_FACTORY_CONSTRUCTOR(zoteroWinWordDocument)
NS_GENERIC_FACTORY_CONSTRUCTOR(zoteroWinWordFieldEnumerator)
NS_GENERIC_FACTORY_CONSTRUCTOR(zoteroWinWordBookmarkEnumerator)
NS_GENERIC_FACTORY_CONSTRUCTOR(zoteroWinWordField)
NS_GENERIC_FACTORY_CONSTRUCTOR(zoteroWinWordBookmark)

// The following line defines CID variables
NS_DEFINE_NAMED_CID(ZOTEROWINWORDAPPLICATION_CID);
NS_DEFINE_NAMED_CID(ZOTEROWINWORDDOCUMENT_CID);
NS_DEFINE_NAMED_CID(ZOTEROWINWORDFIELDENUMERATOR_CID);
NS_DEFINE_NAMED_CID(ZOTEROWINWORDBOOKMARKENUMERATOR_CID);
NS_DEFINE_NAMED_CID(ZOTEROWINWORDFIELD_CID);
NS_DEFINE_NAMED_CID(ZOTEROWINWORDBOOKMARK_CID);

// Build a table of ClassIDs (CIDs) which are implemented by this module. CIDs
// should be completely unique UUIDs.
// each entry has the form { CID, service, factoryproc, constructorproc }
// where factoryproc is usually NULL.
static const mozilla::Module::CIDEntry kCIDs[] = {
    { &kZOTEROWINWORDAPPLICATION_CID, false, NULL, zoteroWinWordApplicationConstructor },
    { &kZOTEROWINWORDDOCUMENT_CID, false, NULL, zoteroWinWordDocumentConstructor },
    { &kZOTEROWINWORDFIELDENUMERATOR_CID, false, NULL, zoteroWinWordFieldEnumeratorConstructor },
    { &kZOTEROWINWORDBOOKMARKENUMERATOR_CID, false, NULL, zoteroWinWordBookmarkEnumeratorConstructor },
    { &kZOTEROWINWORDFIELD_CID, false, NULL, zoteroWinWordFieldConstructor },
    { &kZOTEROWINWORDBOOKMARK_CID, false, NULL, zoteroWinWordBookmarkConstructor },
    { NULL }
};

// Build a table which maps contract IDs to CIDs.
// A contract is a string which identifies a particular set of functionality. In some
// cases an extension component may override the contract ID of a builtin gecko component
// to modify or extend functionality.
static const mozilla::Module::ContractIDEntry kContracts[] = {
    { ZOTEROWINWORDAPPLICATION_CONTRACTID, &kZOTEROWINWORDAPPLICATION_CID },
    { ZOTEROWINWORDDOCUMENT_CONTRACTID, &kZOTEROWINWORDDOCUMENT_CID },
    { ZOTEROWINWORDFIELDENUMERATOR_CONTRACTID, &kZOTEROWINWORDFIELDENUMERATOR_CID },
    { ZOTEROWINWORDBOOKMARKENUMERATOR_CONTRACTID, &kZOTEROWINWORDBOOKMARKENUMERATOR_CID },
    { ZOTEROWINWORDFIELD_CONTRACTID, &kZOTEROWINWORDFIELD_CID },
    { ZOTEROWINWORDBOOKMARK_CONTRACTID, &kZOTEROWINWORDBOOKMARK_CID },
    { NULL }
};

static const mozilla::Module kzoteroWinWordModule = {
    mozilla::Module::kVersion,
    kCIDs,
    kContracts,
    NULL
};

// The following line implements the one-and-only "NSModule" symbol exported from this
// shared library.
NSMODULE_DEFN(zoteroWinWordModule) = &kzoteroWinWordModule;

// The following line implements the one-and-only "NSGetModule" symbol
// for compatibility with mozilla 1.9.2. You should only use this
// if you need a binary which is backwards-compatible and if you use
// interfaces carefully across multiple versions.
NS_IMPL_MOZILLA192_NSGETMODULE(&kzoteroWinWordModule)