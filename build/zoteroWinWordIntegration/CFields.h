// Machine generated IDispatch wrapper class(es) created with Add Class from Typelib Wizard
#import "C:\\Program Files (x86)\\Common Files\\Microsoft Shared\\OFFICE12\\MSO.DLL" auto_rename
#import "C:\\Program Files (x86)\\Common Files\\Microsoft Shared\\VBA\\VBA6\\VBE6EXT.OLB" auto_rename
#import "C:\\Program Files (x86)\\Microsoft Office\\Office12\\MSWORD.OLB" auto_rename

// CFields wrapper class

class CFields : public COleDispatchDriver
{
public:
	CFields(){} // Calls COleDispatchDriver default constructor
	CFields(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CFields(const CFields& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// Attributes
public:

	// Operations
public:


	// Fields methods
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
	long get_Count()
	{
		long result;
		InvokeHelper(0x1, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	long get_Locked()
	{
		long result;
		InvokeHelper(0x2, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_Locked(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x2, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPUNKNOWN get__NewEnum()
	{
		LPUNKNOWN result;
		InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH Item(long Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x0, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Index);
		return result;
	}
	void ToggleShowCodes()
	{
		InvokeHelper(0x64, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	long Update()
	{
		long result;
		InvokeHelper(0x65, DISPATCH_METHOD, VT_I4, (void*)&result, NULL);
		return result;
	}
	void Unlink()
	{
		InvokeHelper(0x66, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void UpdateSource()
	{
		InvokeHelper(0x68, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	LPDISPATCH Add(LPDISPATCH Range, VARIANT * Type, VARIANT * Text, VARIANT * PreserveFormatting)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x69, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Range, Type, Text, PreserveFormatting);
		return result;
	}

	// Fields properties
public:

};
