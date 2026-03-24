# Zotero Word for Windows Integration

This is a Zotero extension that consists of a library written in C++ that communicates with Microsoft Word out of process using OLE Automation, a js-ctypes wrapper for said library, and a template that is installed into Microsoft Word to communicate with Zotero.

## Development Setup

After cloning, install git hooks:

```
./scripts/install_hooks.sh
```

This sets up a pre-commit hook that automatically runs `build/template/unpack_templates.sh` when `install/Zotero.dotm` is modified, keeping the unpacked template files in sync.

## C++ Library Build Requirements
- Visual Studio 2022
- Microsoft Office (previously built with 2010, but newer versions should work)

## To Build the C++ Library
- Open `build/zoteroWinWordIntegration/zoteroWinWordIntegration.sln`
- Change `imports.h `to point to the appropriate files (may be in different places with newer Office)
- Set to Release configuration in the dropdown in the toolbar
- Set to Win32 target in dropdown to the right of Release dropdown
- Build->Build Solution
- Set to x64 target in dropdown
- Build->Build Solution
- Set to ARM64 target in dropdown
- Build->Build Solution

## Template Build Requirements
- Templates should be built with the oldest version of Word to be supported. Otherwise older versions of Word may fail to function properly.

## To Modify/Build the Templates
- Open the template from inside Microsoft Word
- Go to View->Macros->View Macros and click "Edit" for one of the Zotero macros
- Edit/replace code as desired
- Go to Debug->Compile Project to ensure there are no code errors
- Run `build/template/unpack_templates.sh`

## Development Starter's Guide

Start by opening the dotm template in Word. Word templates have support for custom macros 
and adding UI elements to call the macros, which is how the extension is implemented on Word. 
RibbonUI can be edited by extracting the dotm file or using the [Custom UI Editor](https://bettersolutions.com/vba/ribbon/custom-ui-editor.htm). 
In VBA macro code you will find that [SendMessage](https://msdn.microsoft.com/en-us/library/windows/desktop/ms644950(v=vs.85).aspx)
protocol is used to issue commands to Zotero process from Word. These commands are received in [commandLineHandler.js](https://github.com/zotero/zotero/blob/81591eb890b76bb98ff348e1f7d08356188ec8f0/chrome/content/zotero/xpcom/commandLineHandler.js#L102-L110)
where they are passed to integration.js.

Zotero talks to Word via [js-ctypes bindings](https://github.com/zotero/zotero-word-for-windows-integration/blob/e960ae3a3cc676f6279d61c8b222bb9a8f149da6/components/zoteroWinWordIntegration.mjs)
to a C++ OLE Automation based [library](https://github.com/zotero/zotero-word-for-windows-integration/blob/e960ae3a3cc676f6279d61c8b222bb9a8f149da6/build/zoteroWinWordIntegration/zoteroWinWordIntegration.h).
To generate new interfaces for Word interop communications you should use the Add New Class wizard in
Visual Studio and select 'MFC Class from Typelib'. The interop API docs can be found in the [MSDN](https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word._document?view=word-pia).
