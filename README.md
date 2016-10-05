# Zotero Word for Windows Integration

This is a Firefox add-on that consists of a library written in C++ that communicates with Microsoft Word out of process using OLE Automation, a js-ctypes wrapper for said library, and a template that is installed into Microsoft Word to communicate with Zotero.

## C++ Library Build Requirements
- Visual Studio (currently 2013)
- Microsoft Office (previously build with 2010, but newer versions should work)

## To Build
- Open `build/zoteroWinWordIntegration/zoteroWinWordIntegration.sln`
- Change `imports.h `to point to the appropriate files (may be in different places with newer Office)
- Set to Release configuration in the dropdown in the toolbar
- Build->Build Solution
