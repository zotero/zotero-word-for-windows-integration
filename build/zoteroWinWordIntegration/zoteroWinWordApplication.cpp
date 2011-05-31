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

#include "zoteroWinWordApplication.h"
#include "zoteroWinWordDocument.h"
#include "nsStringAPI.h"
#include "zoteroException.h"

/* Implementation file */
NS_IMPL_ISUPPORTS1(zoteroWinWordApplication, zoteroIntegrationApplication)

zoteroWinWordApplication::zoteroWinWordApplication()
{
  /* member initializers and constructor code */
}

zoteroWinWordApplication::~zoteroWinWordApplication()
{
  /* destructor code */
}

/* readonly attribute zoteroIntegrationDocument activeDocument; */
NS_IMETHODIMP zoteroWinWordApplication::GetActiveDocument(zoteroIntegrationDocument **_retval)
{
	ZOTERO_EXCEPTION_CATCHER_START
	*_retval = new zoteroWinWordDocument();
	(*_retval)->AddRef();
	return NS_OK;
	ZOTERO_EXCEPTION_CATCHER_END
}

NS_IMETHODIMP zoteroWinWordApplication::GetDocument(const PRUnichar* documentIdentifier, zoteroIntegrationDocument **_retval)
{
	ZOTERO_EXCEPTION_CATCHER_START
	*_retval = new zoteroWinWordDocument(documentIdentifier);
	(*_retval)->AddRef();
	return NS_OK;
	ZOTERO_EXCEPTION_CATCHER_END
}

/* readonly attribute ACString primaryFieldType; */
NS_IMETHODIMP zoteroWinWordApplication::GetPrimaryFieldType(nsACString & aPrimaryFieldType)
{
	aPrimaryFieldType.AppendLiteral("Field");
    return NS_OK;
}

/* readonly attribute ACString secondaryFieldType; */
NS_IMETHODIMP zoteroWinWordApplication::GetSecondaryFieldType(nsACString & aSecondaryFieldType)
{
	aSecondaryFieldType.AppendLiteral("Bookmark");
    return NS_OK;
}

/* End of implementation class template. */
