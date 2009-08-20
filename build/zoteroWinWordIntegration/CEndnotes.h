// Machine generated IDispatch wrapper class(es) created with Add Class from Typelib Wizard
#import "C:\\Program Files (x86)\\Common Files\\Microsoft Shared\\OFFICE12\\MSO.DLL" auto_rename
#import "C:\\Program Files (x86)\\Common Files\\Microsoft Shared\\VBA\\VBA6\\VBE6EXT.OLB" auto_rename
#import "C:\\Program Files (x86)\\Microsoft Office\\Office12\\MSWORD.OLB" auto_rename

// CEndnotes wrapper class

class CEndnotes : public COleDispatchDriver
{
public:
	CEndnotes(){} // Calls COleDispatchDriver default constructor
	CEndnotes(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CEndnotes(const CEndnotes& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// Attributes
public:

	// Operations
public:


	// Endnotes methods
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
	long get_Location()
	{
		long result;
		InvokeHelper(0x64, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_Location(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x64, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_NumberStyle()
	{
		long result;
		InvokeHelper(0x65, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_NumberStyle(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x65, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_StartingNumber()
	{
		long result;
		InvokeHelper(0x66, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_StartingNumber(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x66, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_NumberingRule()
	{
		long result;
		InvokeHelper(0x67, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_NumberingRule(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x67, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_Separator()
	{
		LPDISPATCH result;
		InvokeHelper(0x68, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_ContinuationSeparator()
	{
		LPDISPATCH result;
		InvokeHelper(0x69, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_ContinuationNotice()
	{
		LPDISPATCH result;
		InvokeHelper(0x6a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH Item(long Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x0, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Index);
		return result;
	}
	LPDISPATCH Add(LPDISPATCH Range, VARIANT * Reference, VARIANT * Text)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x4, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Range, Reference, Text);
		return result;
	}
	void Convert()
	{
		InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void SwapWithFootnotes()
	{
		InvokeHelper(0x6, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void ResetSeparator()
	{
		InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void ResetContinuationSeparator()
	{
		InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void ResetContinuationNotice()
	{
		InvokeHelper(0x9, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}

	// Endnotes properties
public:

};
