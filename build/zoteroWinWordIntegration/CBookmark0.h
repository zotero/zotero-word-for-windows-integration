// Machine generated IDispatch wrapper class(es) created with Add Class from Typelib Wizard
#import "C:\\Program Files (x86)\\Common Files\\Microsoft Shared\\Office14\\MSO.DLL" auto_rename
#import "C:\\Program Files (x86)\\Common Files\\Microsoft Shared\\VBA\\VBA6\\VBE6EXT.OLB" auto_rename
#import "C:\\Program Files (x86)\\Microsoft Office\\Office14\\MSWORD.OLB" auto_rename

// CBookmark0 wrapper class

class CBookmark0 : public COleDispatchDriver
{
public:
	CBookmark0(){} // Calls COleDispatchDriver default constructor
	CBookmark0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CBookmark0(const CBookmark0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// Attributes
public:

	// Operations
public:


	// Bookmark methods
public:
	CString get_Name()
	{
		CString result;
		InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Range()
	{
		LPDISPATCH result;
		InvokeHelper(0x1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	BOOL get_Empty()
	{
		BOOL result;
		InvokeHelper(0x2, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	long get_Start()
	{
		long result;
		InvokeHelper(0x3, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_Start(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x3, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_End()
	{
		long result;
		InvokeHelper(0x4, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_End(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x4, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_Column()
	{
		BOOL result;
		InvokeHelper(0x5, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	long get_StoryType()
	{
		long result;
		InvokeHelper(0x6, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
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
	void Select()
	{
		InvokeHelper(0xffff, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Delete()
	{
		InvokeHelper(0xb, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	LPDISPATCH Copy(LPCTSTR Name)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0xc, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Name);
		return result;
	}

	// Bookmark properties
public:

};
