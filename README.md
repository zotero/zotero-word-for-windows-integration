# Zotero Word for Windows Integration

This is a Firefox add-on that consists of a library written in C++ that communicates with Microsoft Word out of process using OLE Automation, a js-ctypes wrapper for said library, and a template that is installed into Microsoft Word to communicate with Zotero.

## C++ Library Build Requirements
- Visual Studio (currently 2013)
- Microsoft Office (previously build with 2010, but newer versions should work)

## To Build the C++ Library
- Open `build/zoteroWinWordIntegration/zoteroWinWordIntegration.sln`
- Change `imports.h `to point to the appropriate files (may be in different places with newer Office)
- Set to Release configuration in the dropdown in the toolbar
- Set to Win32 target in dropdown to the right of Release dropdown
- Build->Build Solution
- Set to x64 target in dropdown
- Build->Build Solution

## Template Build Requirements
- Templates should be built with the oldest version of Word to be supported. Otherwise older versions of Word may fail to function properly. This is currently:
  - Word 2007 (for the ribbonized dotm template)
  - Word 2003 (for the old dot template)

## To Modify/Build the Templates
- Open the template from inside Microsoft Word
- Go to View->Macros->View Macros (Ribbonized Word) or Tools->Macros->View Macros (Word 2003) and click "Edit" for one of the Zotero macros
- Edit/replace code as desired
- Go to Debug->Compile Project to ensure there are no code errors
- Run `build/template/unpack_templates.sh`
