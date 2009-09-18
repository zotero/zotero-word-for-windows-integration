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

#include "zoteroWinWordDocument.h"
#include "zoteroWinWordField.h"
#include "zoteroWinWordBookmark.h"
#include "zoteroException.h"

zoteroWinWordBookmark::zoteroWinWordBookmark() {}
zoteroWinWordBookmark::zoteroWinWordBookmark(zoteroWinWordDocument *aDoc, CBookmark0 bookmark)
{
	comBookmark = bookmark;
	doc = aDoc;
	init(true);
}

zoteroWinWordBookmark::zoteroWinWordBookmark(zoteroWinWordDocument *aDoc, CBookmark0 bookmark, CString aBookmarkName)
{
	try {
		comBookmark = bookmark;
		doc = aDoc;
		bookmarkName = aBookmarkName;
		init(false);
	} catch(...) {}
}


/* void removeCode (); */
NS_IMETHODIMP zoteroWinWordBookmark::RemoveCode()
{
	ZOTERO_EXCEPTION_CATCHER_START
	comBookmark.Delete();
	doc->setProperty(bookmarkName, L"");
	return NS_OK;
	ZOTERO_EXCEPTION_CATCHER_END
}

/* wstring getCode (); */
NS_IMETHODIMP zoteroWinWordBookmark::GetCode(PRUnichar **_retval)
{
	ZOTERO_EXCEPTION_CATCHER_START
	CString prefix;
	CString comString = doc->getProperty(bookmarkName);
	if(wcsncmp(comString, BOOKMARK_PREFIX, BOOKMARK_PREFIX.GetLength()) == 0) {
		prefix = BOOKMARK_PREFIX;
	} else if(wcsncmp(comString, BACKUP_BOOKMARK_PREFIX, BACKUP_BOOKMARK_PREFIX.GetLength()) == 0) {
		prefix = BACKUP_BOOKMARK_PREFIX;
	} else {
		return NS_ERROR_FAILURE;
	}

	long length = comString.GetLength()-prefix.GetLength();
	*_retval = (PRUnichar *) NS_Alloc((length+1) * sizeof(PRUnichar));
	lstrcpyn(*_retval, ((LPCTSTR)comString)+(prefix.GetLength()), length+1);
	return NS_OK;
	ZOTERO_EXCEPTION_CATCHER_END
}

/* void setCode (in wstring code); */
NS_IMETHODIMP zoteroWinWordBookmark::SetCode(const PRUnichar *code)
{
	ZOTERO_EXCEPTION_CATCHER_START
	doc->setProperty(bookmarkName, BOOKMARK_PREFIX+code);
	return NS_OK;
	ZOTERO_EXCEPTION_CATCHER_END
}

/* void setText (in wstring text, in boolean isRich); */
NS_IMETHODIMP zoteroWinWordBookmark::SetText(const PRUnichar *text, PRBool isRich)
{
	ZOTERO_EXCEPTION_CATCHER_START
	if(isRich) {
		// InsertFile method will clobber the bookmark, so add it to the RTF
		wchar_t *bookmarkText = new wchar_t[32+bookmarkName.GetLength()*2+lstrlen(text)];
		lstrcpy(bookmarkText, L"{\\rtf {\\bkmkstart ");
		lstrcat(bookmarkText, bookmarkName);
		lstrcat(bookmarkText, L"}{");
		lstrcat(bookmarkText, text+6);
		lstrcat(bookmarkText, L"{\\bkmkend ");
		lstrcat(bookmarkText, bookmarkName);
		lstrcat(bookmarkText, L"}}");
		nsresult retval = zoteroWinWordField::SetText(bookmarkText, true);
		delete[] bookmarkText;
		return retval;
	} else {
		nsresult retval = zoteroWinWordField::SetText(text, false);
		if(retval == NS_OK) {
			// setting the text of the bookmark erases it, so we need to recreate
			CDocument0 comDoc = comTextRange.get_Document();
			CBookmarks comBookmarks = comDoc.get_Bookmarks();
			comBookmark = comBookmarks.Add(bookmarkName, comTextRange);
		}
		return retval;
	}
	ZOTERO_EXCEPTION_CATCHER_END
}

void zoteroWinWordBookmark::init(bool needName)
{
	offset1 = -1;
	if(needName) {
		bookmarkName = comBookmark.get_Name();
	}
	comTextRange = comBookmark.get_Range();
}

void zoteroWinWordBookmark::loadFromRange(CRange comRange)
{
	CBookmarks comBookmarks = comRange.get_Bookmarks();
	// this gets called from convertToNoteType(), which may end up deleting
	// the bookmark, so we may have to re-create it. this is not a problem
	// with fields.
	if(comBookmarks.get_Count()) {
		comBookmark = comBookmarks.Item(1);
	} else {
		CBookmark0 comBookmark = comBookmarks.Add(bookmarkName, comRange);
	}
	init(false);
}

CRange zoteroWinWordBookmark::getFieldRange()
{
	return comTextRange.get_Duplicate();
}

void zoteroWinWordBookmark::deleteField()
{
	SetText(L"", false);
	RemoveCode();
}

/* End of implementation class template. */

zoteroWinWordBookmarkEnumerator::zoteroWinWordBookmarkEnumerator() {}
zoteroWinWordBookmarkEnumerator::zoteroWinWordBookmarkEnumerator(zoteroWinWordDocument *aDoc) {
	doc = aDoc;
	CStoryRanges comStoryRanges = doc->comDoc.get_StoryRanges();
	CRange comStoryRange;
	for(short i=0; i<3; i++) {
		bookmarkIndex[i] = 0;

		try {
			comStoryRange = comStoryRanges.Item(i+1);
		} catch(COleDispatchException* e) {
			e->Delete();
			fieldItem[i] = NULL;
			bookmarkCount[i] = 0;
			continue;
		}

		comBookmarks[i] = comStoryRange.get_Bookmarks();
		bookmarkCount[i] = comBookmarks[i].get_Count();
		
		// get first zoteroWinWordField
		fetchNextItem(i);
	}
}

void zoteroWinWordBookmarkEnumerator::fetchNextItem(short i) {
	if(bookmarkIndex[i] < bookmarkCount[i]) {
		CBookmark0 comBookmark = comBookmarks[i].Item(bookmarkIndex[i]++ + 1);
		CString bookmarkName = comBookmark.get_Name();
		if(wcsncmp(bookmarkName, BOOKMARK_PREFIX, BOOKMARK_PREFIX.GetLength()) == 0
				|| wcsncmp(bookmarkName, BACKUP_BOOKMARK_PREFIX, BACKUP_BOOKMARK_PREFIX.GetLength()) == 0) {
			doc->AddRef();
			fieldItem[i] = new zoteroWinWordBookmark(doc, comBookmark);
		} else {
			fetchNextItem(i);
		}
	} else {
		fieldItem[i] = NULL;
	}
}