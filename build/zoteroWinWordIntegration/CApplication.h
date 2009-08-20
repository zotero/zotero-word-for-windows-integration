// Machine generated IDispatch wrapper class(es) created with Add Class from Typelib Wizard
#import "C:\\Program Files (x86)\\Common Files\\Microsoft Shared\\OFFICE12\\MSO.DLL" auto_rename
#import "C:\\Program Files (x86)\\Common Files\\Microsoft Shared\\VBA\\VBA6\\VBE6EXT.OLB" auto_rename
#import "C:\\Program Files (x86)\\Microsoft Office\\Office12\\MSWORD.OLB" auto_rename

// CApplication wrapper class

class CApplication : public COleDispatchDriver
{
public:
	CApplication(){} // Calls COleDispatchDriver default constructor
	CApplication(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CApplication(const CApplication& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// Attributes
public:

	// Operations
public:


	// _Application methods
public:
	LPDISPATCH get_Application()
	{
		LPDISPATCH result;
		InvokeHelper(0x3e8, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Creator()
	{
		long result;
		InvokeHelper(0x3e9, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Parent()
	{
		LPDISPATCH result;
		InvokeHelper(0x3ea, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	CString get_Name()
	{
		CString result;
		InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Documents()
	{
		LPDISPATCH result;
		InvokeHelper(0x6, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Windows()
	{
		LPDISPATCH result;
		InvokeHelper(0x2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_ActiveDocument()
	{
		LPDISPATCH result;
		InvokeHelper(0x3, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_ActiveWindow()
	{
		LPDISPATCH result;
		InvokeHelper(0x4, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Selection()
	{
		LPDISPATCH result;
		InvokeHelper(0x5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_WordBasic()
	{
		LPDISPATCH result;
		InvokeHelper(0x1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_RecentFiles()
	{
		LPDISPATCH result;
		InvokeHelper(0x7, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_NormalTemplate()
	{
		LPDISPATCH result;
		InvokeHelper(0x8, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_System()
	{
		LPDISPATCH result;
		InvokeHelper(0x9, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_AutoCorrect()
	{
		LPDISPATCH result;
		InvokeHelper(0xa, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_FontNames()
	{
		LPDISPATCH result;
		InvokeHelper(0xb, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_LandscapeFontNames()
	{
		LPDISPATCH result;
		InvokeHelper(0xc, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_PortraitFontNames()
	{
		LPDISPATCH result;
		InvokeHelper(0xd, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Languages()
	{
		LPDISPATCH result;
		InvokeHelper(0xe, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Assistant()
	{
		LPDISPATCH result;
		InvokeHelper(0xf, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Browser()
	{
		LPDISPATCH result;
		InvokeHelper(0x10, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_FileConverters()
	{
		LPDISPATCH result;
		InvokeHelper(0x11, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_MailingLabel()
	{
		LPDISPATCH result;
		InvokeHelper(0x12, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Dialogs()
	{
		LPDISPATCH result;
		InvokeHelper(0x13, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_CaptionLabels()
	{
		LPDISPATCH result;
		InvokeHelper(0x14, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_AutoCaptions()
	{
		LPDISPATCH result;
		InvokeHelper(0x15, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_AddIns()
	{
		LPDISPATCH result;
		InvokeHelper(0x16, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	BOOL get_Visible()
	{
		BOOL result;
		InvokeHelper(0x17, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_Visible(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x17, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	CString get_Version()
	{
		CString result;
		InvokeHelper(0x18, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	BOOL get_ScreenUpdating()
	{
		BOOL result;
		InvokeHelper(0x1a, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_ScreenUpdating(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x1a, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_PrintPreview()
	{
		BOOL result;
		InvokeHelper(0x1b, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_PrintPreview(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x1b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_Tasks()
	{
		LPDISPATCH result;
		InvokeHelper(0x1c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	BOOL get_DisplayStatusBar()
	{
		BOOL result;
		InvokeHelper(0x1d, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_DisplayStatusBar(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x1d, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_SpecialMode()
	{
		BOOL result;
		InvokeHelper(0x1e, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	long get_UsableWidth()
	{
		long result;
		InvokeHelper(0x21, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	long get_UsableHeight()
	{
		long result;
		InvokeHelper(0x22, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	BOOL get_MathCoprocessorAvailable()
	{
		BOOL result;
		InvokeHelper(0x24, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	BOOL get_MouseAvailable()
	{
		BOOL result;
		InvokeHelper(0x25, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	VARIANT get_International(long Index)
	{
		VARIANT result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x2e, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, parms, Index);
		return result;
	}
	CString get_Build()
	{
		CString result;
		InvokeHelper(0x2f, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	BOOL get_CapsLock()
	{
		BOOL result;
		InvokeHelper(0x30, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	BOOL get_NumLock()
	{
		BOOL result;
		InvokeHelper(0x31, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	CString get_UserName()
	{
		CString result;
		InvokeHelper(0x34, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_UserName(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x34, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	CString get_UserInitials()
	{
		CString result;
		InvokeHelper(0x35, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_UserInitials(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x35, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	CString get_UserAddress()
	{
		CString result;
		InvokeHelper(0x36, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_UserAddress(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x36, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_MacroContainer()
	{
		LPDISPATCH result;
		InvokeHelper(0x37, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	BOOL get_DisplayRecentFiles()
	{
		BOOL result;
		InvokeHelper(0x38, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_DisplayRecentFiles(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x38, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_CommandBars()
	{
		LPDISPATCH result;
		InvokeHelper(0x39, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_SynonymInfo(LPCTSTR Word, VARIANT * LanguageID)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR VTS_PVARIANT ;
		InvokeHelper(0x3b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms, Word, LanguageID);
		return result;
	}
	LPDISPATCH get_VBE()
	{
		LPDISPATCH result;
		InvokeHelper(0x3d, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	CString get_DefaultSaveFormat()
	{
		CString result;
		InvokeHelper(0x40, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_DefaultSaveFormat(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x40, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_ListGalleries()
	{
		LPDISPATCH result;
		InvokeHelper(0x41, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	CString get_ActivePrinter()
	{
		CString result;
		InvokeHelper(0x42, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_ActivePrinter(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x42, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_Templates()
	{
		LPDISPATCH result;
		InvokeHelper(0x43, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_CustomizationContext()
	{
		LPDISPATCH result;
		InvokeHelper(0x44, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void put_CustomizationContext(LPDISPATCH newValue)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x44, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_KeyBindings()
	{
		LPDISPATCH result;
		InvokeHelper(0x45, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_KeysBoundTo(long KeyCategory, LPCTSTR Command, VARIANT * CommandParameter)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 VTS_BSTR VTS_PVARIANT ;
		InvokeHelper(0x46, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms, KeyCategory, Command, CommandParameter);
		return result;
	}
	LPDISPATCH get_FindKey(long KeyCode, VARIANT * KeyCode2)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 VTS_PVARIANT ;
		InvokeHelper(0x47, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms, KeyCode, KeyCode2);
		return result;
	}
	CString get_Caption()
	{
		CString result;
		InvokeHelper(0x50, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_Caption(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x50, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	CString get_Path()
	{
		CString result;
		InvokeHelper(0x51, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	BOOL get_DisplayScrollBars()
	{
		BOOL result;
		InvokeHelper(0x52, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_DisplayScrollBars(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x52, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	CString get_StartupPath()
	{
		CString result;
		InvokeHelper(0x53, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_StartupPath(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x53, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_BackgroundSavingStatus()
	{
		long result;
		InvokeHelper(0x55, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	long get_BackgroundPrintingStatus()
	{
		long result;
		InvokeHelper(0x56, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	long get_Left()
	{
		long result;
		InvokeHelper(0x57, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_Left(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x57, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_Top()
	{
		long result;
		InvokeHelper(0x58, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_Top(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x58, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_Width()
	{
		long result;
		InvokeHelper(0x59, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_Width(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x59, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_Height()
	{
		long result;
		InvokeHelper(0x5a, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_Height(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x5a, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_WindowState()
	{
		long result;
		InvokeHelper(0x5b, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_WindowState(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x5b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_DisplayAutoCompleteTips()
	{
		BOOL result;
		InvokeHelper(0x5c, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_DisplayAutoCompleteTips(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x5c, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_Options()
	{
		LPDISPATCH result;
		InvokeHelper(0x5d, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_DisplayAlerts()
	{
		long result;
		InvokeHelper(0x5e, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_DisplayAlerts(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x5e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_CustomDictionaries()
	{
		LPDISPATCH result;
		InvokeHelper(0x5f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	CString get_PathSeparator()
	{
		CString result;
		InvokeHelper(0x60, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_StatusBar(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x61, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_MAPIAvailable()
	{
		BOOL result;
		InvokeHelper(0x62, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	BOOL get_DisplayScreenTips()
	{
		BOOL result;
		InvokeHelper(0x63, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_DisplayScreenTips(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x63, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_EnableCancelKey()
	{
		long result;
		InvokeHelper(0x64, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_EnableCancelKey(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x64, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_UserControl()
	{
		BOOL result;
		InvokeHelper(0x65, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_FileSearch()
	{
		LPDISPATCH result;
		InvokeHelper(0x67, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_MailSystem()
	{
		long result;
		InvokeHelper(0x68, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	CString get_DefaultTableSeparator()
	{
		CString result;
		InvokeHelper(0x69, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_DefaultTableSeparator(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x69, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_ShowVisualBasicEditor()
	{
		BOOL result;
		InvokeHelper(0x6a, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_ShowVisualBasicEditor(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x6a, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	CString get_BrowseExtraFileTypes()
	{
		CString result;
		InvokeHelper(0x6c, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_BrowseExtraFileTypes(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x6c, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_IsObjectValid(LPDISPATCH Object)
	{
		BOOL result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x6d, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, parms, Object);
		return result;
	}
	LPDISPATCH get_HangulHanjaDictionaries()
	{
		LPDISPATCH result;
		InvokeHelper(0x6e, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_MailMessage()
	{
		LPDISPATCH result;
		InvokeHelper(0x15c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	BOOL get_FocusInMailHeader()
	{
		BOOL result;
		InvokeHelper(0x182, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void Quit(VARIANT * SaveChanges, VARIANT * OriginalFormat, VARIANT * RouteDocument)
	{
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x451, DISPATCH_METHOD, VT_EMPTY, NULL, parms, SaveChanges, OriginalFormat, RouteDocument);
	}
	void ScreenRefresh()
	{
		InvokeHelper(0x12d, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void PrintOutOld(VARIANT * Background, VARIANT * Append, VARIANT * Range, VARIANT * OutputFileName, VARIANT * From, VARIANT * To, VARIANT * Item, VARIANT * Copies, VARIANT * Pages, VARIANT * PageType, VARIANT * PrintToFile, VARIANT * Collate, VARIANT * FileName, VARIANT * ActivePrinterMacGX, VARIANT * ManualDuplexPrint)
	{
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x12e, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Background, Append, Range, OutputFileName, From, To, Item, Copies, Pages, PageType, PrintToFile, Collate, FileName, ActivePrinterMacGX, ManualDuplexPrint);
	}
	void LookupNameProperties(LPCTSTR Name)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x12f, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Name);
	}
	void SubstituteFont(LPCTSTR UnavailableFont, LPCTSTR SubstituteFont)
	{
		static BYTE parms[] = VTS_BSTR VTS_BSTR ;
		InvokeHelper(0x130, DISPATCH_METHOD, VT_EMPTY, NULL, parms, UnavailableFont, SubstituteFont);
	}
	BOOL Repeat(VARIANT * Times)
	{
		BOOL result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x131, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, Times);
		return result;
	}
	void DDEExecute(long Channel, LPCTSTR Command)
	{
		static BYTE parms[] = VTS_I4 VTS_BSTR ;
		InvokeHelper(0x136, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Channel, Command);
	}
	long DDEInitiate(LPCTSTR App, LPCTSTR Topic)
	{
		long result;
		static BYTE parms[] = VTS_BSTR VTS_BSTR ;
		InvokeHelper(0x137, DISPATCH_METHOD, VT_I4, (void*)&result, parms, App, Topic);
		return result;
	}
	void DDEPoke(long Channel, LPCTSTR Item, LPCTSTR Data)
	{
		static BYTE parms[] = VTS_I4 VTS_BSTR VTS_BSTR ;
		InvokeHelper(0x138, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Channel, Item, Data);
	}
	CString DDERequest(long Channel, LPCTSTR Item)
	{
		CString result;
		static BYTE parms[] = VTS_I4 VTS_BSTR ;
		InvokeHelper(0x139, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, Channel, Item);
		return result;
	}
	void DDETerminate(long Channel)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x13a, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Channel);
	}
	void DDETerminateAll()
	{
		InvokeHelper(0x13b, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	long BuildKeyCode(long Arg1, VARIANT * Arg2, VARIANT * Arg3, VARIANT * Arg4)
	{
		long result;
		static BYTE parms[] = VTS_I4 VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x13c, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Arg1, Arg2, Arg3, Arg4);
		return result;
	}
	CString KeyString(long KeyCode, VARIANT * KeyCode2)
	{
		CString result;
		static BYTE parms[] = VTS_I4 VTS_PVARIANT ;
		InvokeHelper(0x13d, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, KeyCode, KeyCode2);
		return result;
	}
	void OrganizerCopy(LPCTSTR Source, LPCTSTR Destination, LPCTSTR Name, long Object)
	{
		static BYTE parms[] = VTS_BSTR VTS_BSTR VTS_BSTR VTS_I4 ;
		InvokeHelper(0x13e, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Source, Destination, Name, Object);
	}
	void OrganizerDelete(LPCTSTR Source, LPCTSTR Name, long Object)
	{
		static BYTE parms[] = VTS_BSTR VTS_BSTR VTS_I4 ;
		InvokeHelper(0x13f, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Source, Name, Object);
	}
	void OrganizerRename(LPCTSTR Source, LPCTSTR Name, LPCTSTR NewName, long Object)
	{
		static BYTE parms[] = VTS_BSTR VTS_BSTR VTS_BSTR VTS_I4 ;
		InvokeHelper(0x140, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Source, Name, NewName, Object);
	}
	void AddAddress(SAFEARRAY * * TagID, SAFEARRAY * * Value)
	{
		static BYTE parms[] = VTS_UNKNOWN VTS_UNKNOWN ;
		InvokeHelper(0x141, DISPATCH_METHOD, VT_EMPTY, NULL, parms, TagID, Value);
	}
	CString GetAddress(VARIANT * Name, VARIANT * AddressProperties, VARIANT * UseAutoText, VARIANT * DisplaySelectDialog, VARIANT * SelectDialog, VARIANT * CheckNamesDialog, VARIANT * RecentAddressesChoice, VARIANT * UpdateRecentAddresses)
	{
		CString result;
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x142, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, Name, AddressProperties, UseAutoText, DisplaySelectDialog, SelectDialog, CheckNamesDialog, RecentAddressesChoice, UpdateRecentAddresses);
		return result;
	}
	BOOL CheckGrammar(LPCTSTR String)
	{
		BOOL result;
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x143, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, String);
		return result;
	}
	BOOL CheckSpelling(LPCTSTR Word, VARIANT * CustomDictionary, VARIANT * IgnoreUppercase, VARIANT * MainDictionary, VARIANT * CustomDictionary2, VARIANT * CustomDictionary3, VARIANT * CustomDictionary4, VARIANT * CustomDictionary5, VARIANT * CustomDictionary6, VARIANT * CustomDictionary7, VARIANT * CustomDictionary8, VARIANT * CustomDictionary9, VARIANT * CustomDictionary10)
	{
		BOOL result;
		static BYTE parms[] = VTS_BSTR VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x144, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, Word, CustomDictionary, IgnoreUppercase, MainDictionary, CustomDictionary2, CustomDictionary3, CustomDictionary4, CustomDictionary5, CustomDictionary6, CustomDictionary7, CustomDictionary8, CustomDictionary9, CustomDictionary10);
		return result;
	}
	void ResetIgnoreAll()
	{
		InvokeHelper(0x146, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	LPDISPATCH GetSpellingSuggestions(LPCTSTR Word, VARIANT * CustomDictionary, VARIANT * IgnoreUppercase, VARIANT * MainDictionary, VARIANT * SuggestionMode, VARIANT * CustomDictionary2, VARIANT * CustomDictionary3, VARIANT * CustomDictionary4, VARIANT * CustomDictionary5, VARIANT * CustomDictionary6, VARIANT * CustomDictionary7, VARIANT * CustomDictionary8, VARIANT * CustomDictionary9, VARIANT * CustomDictionary10)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x147, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Word, CustomDictionary, IgnoreUppercase, MainDictionary, SuggestionMode, CustomDictionary2, CustomDictionary3, CustomDictionary4, CustomDictionary5, CustomDictionary6, CustomDictionary7, CustomDictionary8, CustomDictionary9, CustomDictionary10);
		return result;
	}
	void GoBack()
	{
		InvokeHelper(0x148, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Help(VARIANT * HelpType)
	{
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x149, DISPATCH_METHOD, VT_EMPTY, NULL, parms, HelpType);
	}
	void AutomaticChange()
	{
		InvokeHelper(0x14a, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void ShowMe()
	{
		InvokeHelper(0x14b, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void HelpTool()
	{
		InvokeHelper(0x14c, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	LPDISPATCH NewWindow()
	{
		LPDISPATCH result;
		InvokeHelper(0x159, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void ListCommands(BOOL ListAllCommands)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x15a, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ListAllCommands);
	}
	void ShowClipboard()
	{
		InvokeHelper(0x15d, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void OnTime(VARIANT * When, LPCTSTR Name, VARIANT * Tolerance)
	{
		static BYTE parms[] = VTS_PVARIANT VTS_BSTR VTS_PVARIANT ;
		InvokeHelper(0x15e, DISPATCH_METHOD, VT_EMPTY, NULL, parms, When, Name, Tolerance);
	}
	void NextLetter()
	{
		InvokeHelper(0x15f, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	short MountVolume(LPCTSTR Zone, LPCTSTR Server, LPCTSTR Volume, VARIANT * User, VARIANT * UserPassword, VARIANT * VolumePassword)
	{
		short result;
		static BYTE parms[] = VTS_BSTR VTS_BSTR VTS_BSTR VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x161, DISPATCH_METHOD, VT_I2, (void*)&result, parms, Zone, Server, Volume, User, UserPassword, VolumePassword);
		return result;
	}
	CString CleanString(LPCTSTR String)
	{
		CString result;
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x162, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, String);
		return result;
	}
	void SendFax()
	{
		InvokeHelper(0x164, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void ChangeFileOpenDirectory(LPCTSTR Path)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x165, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Path);
	}
	void RunOld(LPCTSTR MacroName)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x166, DISPATCH_METHOD, VT_EMPTY, NULL, parms, MacroName);
	}
	void GoForward()
	{
		InvokeHelper(0x167, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Move(long Left, long Top)
	{
		static BYTE parms[] = VTS_I4 VTS_I4 ;
		InvokeHelper(0x168, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Left, Top);
	}
	void Resize(long Width, long Height)
	{
		static BYTE parms[] = VTS_I4 VTS_I4 ;
		InvokeHelper(0x169, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Width, Height);
	}
	float InchesToPoints(float Inches)
	{
		float result;
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x172, DISPATCH_METHOD, VT_R4, (void*)&result, parms, Inches);
		return result;
	}
	float CentimetersToPoints(float Centimeters)
	{
		float result;
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x173, DISPATCH_METHOD, VT_R4, (void*)&result, parms, Centimeters);
		return result;
	}
	float MillimetersToPoints(float Millimeters)
	{
		float result;
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x174, DISPATCH_METHOD, VT_R4, (void*)&result, parms, Millimeters);
		return result;
	}
	float PicasToPoints(float Picas)
	{
		float result;
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x175, DISPATCH_METHOD, VT_R4, (void*)&result, parms, Picas);
		return result;
	}
	float LinesToPoints(float Lines)
	{
		float result;
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x176, DISPATCH_METHOD, VT_R4, (void*)&result, parms, Lines);
		return result;
	}
	float PointsToInches(float Points)
	{
		float result;
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x17c, DISPATCH_METHOD, VT_R4, (void*)&result, parms, Points);
		return result;
	}
	float PointsToCentimeters(float Points)
	{
		float result;
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x17d, DISPATCH_METHOD, VT_R4, (void*)&result, parms, Points);
		return result;
	}
	float PointsToMillimeters(float Points)
	{
		float result;
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x17e, DISPATCH_METHOD, VT_R4, (void*)&result, parms, Points);
		return result;
	}
	float PointsToPicas(float Points)
	{
		float result;
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x17f, DISPATCH_METHOD, VT_R4, (void*)&result, parms, Points);
		return result;
	}
	float PointsToLines(float Points)
	{
		float result;
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x180, DISPATCH_METHOD, VT_R4, (void*)&result, parms, Points);
		return result;
	}
	void Activate()
	{
		InvokeHelper(0x181, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	float PointsToPixels(float Points, VARIANT * fVertical)
	{
		float result;
		static BYTE parms[] = VTS_R4 VTS_PVARIANT ;
		InvokeHelper(0x183, DISPATCH_METHOD, VT_R4, (void*)&result, parms, Points, fVertical);
		return result;
	}
	float PixelsToPoints(float Pixels, VARIANT * fVertical)
	{
		float result;
		static BYTE parms[] = VTS_R4 VTS_PVARIANT ;
		InvokeHelper(0x184, DISPATCH_METHOD, VT_R4, (void*)&result, parms, Pixels, fVertical);
		return result;
	}
	void KeyboardLatin()
	{
		InvokeHelper(0x190, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void KeyboardBidi()
	{
		InvokeHelper(0x191, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void ToggleKeyboard()
	{
		InvokeHelper(0x192, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	long Keyboard(long LangId)
	{
		long result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x1be, DISPATCH_METHOD, VT_I4, (void*)&result, parms, LangId);
		return result;
	}
	CString ProductCode()
	{
		CString result;
		InvokeHelper(0x194, DISPATCH_METHOD, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH DefaultWebOptions()
	{
		LPDISPATCH result;
		InvokeHelper(0x195, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void DiscussionSupport(VARIANT * Range, VARIANT * cid, VARIANT * piCSE)
	{
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x197, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Range, cid, piCSE);
	}
	void SetDefaultTheme(LPCTSTR Name, long DocumentType)
	{
		static BYTE parms[] = VTS_BSTR VTS_I4 ;
		InvokeHelper(0x19e, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Name, DocumentType);
	}
	CString GetDefaultTheme(long DocumentType)
	{
		CString result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x1a0, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, DocumentType);
		return result;
	}
	LPDISPATCH get_EmailOptions()
	{
		LPDISPATCH result;
		InvokeHelper(0x185, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Language()
	{
		long result;
		InvokeHelper(0x187, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_COMAddIns()
	{
		LPDISPATCH result;
		InvokeHelper(0x6f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	BOOL get_CheckLanguage()
	{
		BOOL result;
		InvokeHelper(0x70, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_CheckLanguage(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x70, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_LanguageSettings()
	{
		LPDISPATCH result;
		InvokeHelper(0x193, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	BOOL get_Dummy1()
	{
		BOOL result;
		InvokeHelper(0x196, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_AnswerWizard()
	{
		LPDISPATCH result;
		InvokeHelper(0x199, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_FeatureInstall()
	{
		long result;
		InvokeHelper(0x1bf, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_FeatureInstall(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x1bf, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	void PrintOut2000(VARIANT * Background, VARIANT * Append, VARIANT * Range, VARIANT * OutputFileName, VARIANT * From, VARIANT * To, VARIANT * Item, VARIANT * Copies, VARIANT * Pages, VARIANT * PageType, VARIANT * PrintToFile, VARIANT * Collate, VARIANT * FileName, VARIANT * ActivePrinterMacGX, VARIANT * ManualDuplexPrint, VARIANT * PrintZoomColumn, VARIANT * PrintZoomRow, VARIANT * PrintZoomPaperWidth, VARIANT * PrintZoomPaperHeight)
	{
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x1bc, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Background, Append, Range, OutputFileName, From, To, Item, Copies, Pages, PageType, PrintToFile, Collate, FileName, ActivePrinterMacGX, ManualDuplexPrint, PrintZoomColumn, PrintZoomRow, PrintZoomPaperWidth, PrintZoomPaperHeight);
	}
	VARIANT Run(LPCTSTR MacroName, VARIANT * varg1, VARIANT * varg2, VARIANT * varg3, VARIANT * varg4, VARIANT * varg5, VARIANT * varg6, VARIANT * varg7, VARIANT * varg8, VARIANT * varg9, VARIANT * varg10, VARIANT * varg11, VARIANT * varg12, VARIANT * varg13, VARIANT * varg14, VARIANT * varg15, VARIANT * varg16, VARIANT * varg17, VARIANT * varg18, VARIANT * varg19, VARIANT * varg20, VARIANT * varg21, VARIANT * varg22, VARIANT * varg23, VARIANT * varg24, VARIANT * varg25, VARIANT * varg26, VARIANT * varg27, VARIANT * varg28, VARIANT * varg29, VARIANT * varg30)
	{
		VARIANT result;
		static BYTE parms[] = VTS_BSTR VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x1bd, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, MacroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22, varg23, varg24, varg25, varg26, varg27, varg28, varg29, varg30);
		return result;
	}
	void PrintOut(VARIANT * Background, VARIANT * Append, VARIANT * Range, VARIANT * OutputFileName, VARIANT * From, VARIANT * To, VARIANT * Item, VARIANT * Copies, VARIANT * Pages, VARIANT * PageType, VARIANT * PrintToFile, VARIANT * Collate, VARIANT * FileName, VARIANT * ActivePrinterMacGX, VARIANT * ManualDuplexPrint, VARIANT * PrintZoomColumn, VARIANT * PrintZoomRow, VARIANT * PrintZoomPaperWidth, VARIANT * PrintZoomPaperHeight)
	{
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x1c0, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Background, Append, Range, OutputFileName, From, To, Item, Copies, Pages, PageType, PrintToFile, Collate, FileName, ActivePrinterMacGX, ManualDuplexPrint, PrintZoomColumn, PrintZoomRow, PrintZoomPaperWidth, PrintZoomPaperHeight);
	}
	long get_AutomationSecurity()
	{
		long result;
		InvokeHelper(0x1c1, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_AutomationSecurity(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x1c1, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_FileDialog(long FileDialogType)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x1c2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms, FileDialogType);
		return result;
	}
	CString get_EmailTemplate()
	{
		CString result;
		InvokeHelper(0x1c3, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_EmailTemplate(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x1c3, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_ShowWindowsInTaskbar()
	{
		BOOL result;
		InvokeHelper(0x1c4, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_ShowWindowsInTaskbar(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x1c4, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_NewDocument()
	{
		LPDISPATCH result;
		InvokeHelper(0x1c6, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	BOOL get_ShowStartupDialog()
	{
		BOOL result;
		InvokeHelper(0x1c7, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_ShowStartupDialog(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x1c7, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_AutoCorrectEmail()
	{
		LPDISPATCH result;
		InvokeHelper(0x1c8, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_TaskPanes()
	{
		LPDISPATCH result;
		InvokeHelper(0x1c9, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	BOOL get_DefaultLegalBlackline()
	{
		BOOL result;
		InvokeHelper(0x1cb, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_DefaultLegalBlackline(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x1cb, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL Dummy2()
	{
		BOOL result;
		InvokeHelper(0x1ca, DISPATCH_METHOD, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_SmartTagRecognizers()
	{
		LPDISPATCH result;
		InvokeHelper(0x1cc, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_SmartTagTypes()
	{
		LPDISPATCH result;
		InvokeHelper(0x1cd, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_XMLNamespaces()
	{
		LPDISPATCH result;
		InvokeHelper(0x1cf, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void PutFocusInMailHeader()
	{
		InvokeHelper(0x1d0, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	BOOL get_ArbitraryXMLSupportAvailable()
	{
		BOOL result;
		InvokeHelper(0x1d1, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	CString get_BuildFull()
	{
		CString result;
		InvokeHelper(0x1d2, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	CString get_BuildFeatureCrew()
	{
		CString result;
		InvokeHelper(0x1d3, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void LoadMasterList(LPCTSTR FileName)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x1d5, DISPATCH_METHOD, VT_EMPTY, NULL, parms, FileName);
	}
	LPDISPATCH CompareDocuments(LPDISPATCH OriginalDocument, LPDISPATCH RevisedDocument, long Destination, long Granularity, BOOL CompareFormatting, BOOL CompareCaseChanges, BOOL CompareWhitespace, BOOL CompareTables, BOOL CompareHeaders, BOOL CompareFootnotes, BOOL CompareTextboxes, BOOL CompareFields, BOOL CompareComments, BOOL CompareMoves, LPCTSTR RevisedAuthor, BOOL IgnoreAllComparisonWarnings)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH VTS_I4 VTS_I4 VTS_BOOL VTS_BOOL VTS_BOOL VTS_BOOL VTS_BOOL VTS_BOOL VTS_BOOL VTS_BOOL VTS_BOOL VTS_BOOL VTS_BSTR VTS_BOOL ;
		InvokeHelper(0x1d6, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, OriginalDocument, RevisedDocument, Destination, Granularity, CompareFormatting, CompareCaseChanges, CompareWhitespace, CompareTables, CompareHeaders, CompareFootnotes, CompareTextboxes, CompareFields, CompareComments, CompareMoves, RevisedAuthor, IgnoreAllComparisonWarnings);
		return result;
	}
	LPDISPATCH MergeDocuments(LPDISPATCH OriginalDocument, LPDISPATCH RevisedDocument, long Destination, long Granularity, BOOL CompareFormatting, BOOL CompareCaseChanges, BOOL CompareWhitespace, BOOL CompareTables, BOOL CompareHeaders, BOOL CompareFootnotes, BOOL CompareTextboxes, BOOL CompareFields, BOOL CompareComments, BOOL CompareMoves, LPCTSTR OriginalAuthor, LPCTSTR RevisedAuthor, long FormatFrom)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH VTS_I4 VTS_I4 VTS_BOOL VTS_BOOL VTS_BOOL VTS_BOOL VTS_BOOL VTS_BOOL VTS_BOOL VTS_BOOL VTS_BOOL VTS_BOOL VTS_BSTR VTS_BSTR VTS_I4 ;
		InvokeHelper(0x1d7, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, OriginalDocument, RevisedDocument, Destination, Granularity, CompareFormatting, CompareCaseChanges, CompareWhitespace, CompareTables, CompareHeaders, CompareFootnotes, CompareTextboxes, CompareFields, CompareComments, CompareMoves, OriginalAuthor, RevisedAuthor, FormatFrom);
		return result;
	}
	LPDISPATCH get_Bibliography()
	{
		LPDISPATCH result;
		InvokeHelper(0x1d8, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	BOOL get_ShowStylePreviews()
	{
		BOOL result;
		InvokeHelper(0x1d9, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_ShowStylePreviews(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x1d9, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_RestrictLinkedStyles()
	{
		BOOL result;
		InvokeHelper(0x1da, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_RestrictLinkedStyles(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x1da, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_OMathAutoCorrect()
	{
		LPDISPATCH result;
		InvokeHelper(0x1db, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	BOOL get_DisplayDocumentInformationPanel()
	{
		BOOL result;
		InvokeHelper(0x1dc, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_DisplayDocumentInformationPanel(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x1dc, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_Assistance()
	{
		LPDISPATCH result;
		InvokeHelper(0x1dd, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	BOOL get_OpenAttachmentsInFullScreen()
	{
		BOOL result;
		InvokeHelper(0x1de, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_OpenAttachmentsInFullScreen(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x1de, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_ActiveEncryptionSession()
	{
		long result;
		InvokeHelper(0x1df, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	BOOL get_DontResetInsertionPointProperties()
	{
		BOOL result;
		InvokeHelper(0x1e0, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_DontResetInsertionPointProperties(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x1e0, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}

	// _Application properties
public:

};
