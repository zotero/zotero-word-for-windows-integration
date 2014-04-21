// Machine generated IDispatch wrapper class(es) created with Add Class from Typelib Wizard
#import "C:\\Program Files (x86)\\Common Files\\Microsoft Shared\\Office14\\MSO.DLL" auto_rename
#import "C:\\Program Files (x86)\\Common Files\\Microsoft Shared\\VBA\\VBA6\\VBE6EXT.OLB" auto_rename
#import "C:\\Program Files (x86)\\Microsoft Office\\Office14\\MSWORD.OLB" auto_rename

// CStoryRanges wrapper class

class CStoryRanges : public COleDispatchDriver
{
public:
	CStoryRanges(){} // Calls COleDispatchDriver default constructor
	CStoryRanges(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CStoryRanges(const CStoryRanges& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// Attributes
public:

	// Operations
public:


	// StoryRanges methods
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
	LPDISPATCH Item(long Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x0, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Index);
		return result;
	}

	// StoryRanges properties
public:

};
