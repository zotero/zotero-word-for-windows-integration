// Machine generated IDispatch wrapper class(es) created with Add Class from Typelib Wizard

// #import "C:\\Program Files (x86)\\Microsoft Office 2013\\Office15\\MSWORD.OLB" no_namespace
// CRevisionsFilter wrapper class

class CRevisionsFilter : public COleDispatchDriver
{
public:
	CRevisionsFilter() {} // Calls COleDispatchDriver default constructor
	CRevisionsFilter(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CRevisionsFilter(const CRevisionsFilter& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// Attributes
public:

	// Operations
public:


	// RevisionsFilter methods
public:
	long get_View()
	{
		long result;
		InvokeHelper(0x1, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	void put_View(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x1, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	long get_Markup()
	{
		long result;
		InvokeHelper(0x2, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	void put_Markup(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x2, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	LPDISPATCH get_Reviewers()
	{
		LPDISPATCH result;
		InvokeHelper(0x3, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	void ToggleShowAllReviewers()
	{
		InvokeHelper(0x65, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}

	// RevisionsFilter properties
public:

};
