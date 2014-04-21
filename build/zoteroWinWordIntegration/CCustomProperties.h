// Machine generated IDispatch wrapper class(es) created with Add Class from Typelib Wizard
#import "C:\\Program Files (x86)\\Common Files\\Microsoft Shared\\Office14\\MSO.DLL" auto_rename
#import "C:\\Program Files (x86)\\Common Files\\Microsoft Shared\\VBA\\VBA6\\VBE6EXT.OLB" auto_rename
#import "C:\\Program Files (x86)\\Microsoft Office\\Office14\\MSWORD.OLB" auto_rename

// CCustomProperties wrapper class

class CCustomProperties : public COleDispatchDriver
{
public:
	CCustomProperties(){} // Calls COleDispatchDriver default constructor
	CCustomProperties(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CCustomProperties(const CCustomProperties& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// Attributes
public:

	// Operations
public:


	// CustomProperties methods
public:
	LPUNKNOWN get__NewEnum()
	{
		LPUNKNOWN result;
		InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, NULL);
		return result;
	}
	long get_Count()
	{
		long result;
		InvokeHelper(0x2, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH Item(LPCTSTR Name)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms, Name);
		return result;
	}
	LPDISPATCH Item(long Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms, Index);
		return result;
	}
	// NOTE: The generated header here is actually for another (similar) class; it has been modified to work
	LPDISPATCH Add(LPCTSTR Name, bool LinkToContent, short Type, LPCTSTR Value)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR VTS_BOOL VTS_I4 VTS_BSTR ;
		InvokeHelper(0x5, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Name, LinkToContent, Type, Value);
		return result;
	}

	// CustomProperties properties
public:

};
