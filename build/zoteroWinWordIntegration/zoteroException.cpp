#include "zoteroException.h"
#include "nsCRTGlue.h"

#ifndef nsCOMPtr_h___
#include "nsCOMPtr.h"
#endif
#ifndef __gen_nsIExceptionService_h__
#include "nsIExceptionService.h"
#endif
#include "nsServiceManagerUtils.h"

NS_IMPL_ISUPPORTS1(zoteroException, nsIException)

zoteroException::zoteroException()
{
	/* member initializers and constructor code */
	errorID = 0;
	freeMessage = false;
}

zoteroException::zoteroException(char *aMessage, char *aFunction, char *aFilename)
{
	errorID = 0;
	message = aMessage;
	function = aFunction;
	filename = aFilename;
	freeMessage = false;
}

zoteroException::zoteroException(CException *exception, char *aFunction, char *aFilename)
{
	errorID = 0;
	wchar_t aMessage[512];
	if(exception->GetErrorMessage(aMessage, 512, &errorID)) {
		size_t returnValue;
		message = new char[512]();
		wcstombs_s(&returnValue, message, 512, aMessage, _TRUNCATE);
		freeMessage = true;
	} else {
		message = NULL;
		freeMessage = false;
	}
	exception->Delete();
	function = aFunction;
	filename = aFilename;
}

zoteroException::~zoteroException()
{
	if(freeMessage) {
		delete[] message;
	}
}

/* [binaryname (MessageMoz)] readonly attribute string message; */
NS_IMETHODIMP zoteroException::GetMessageMoz(char * *aMessage)
{
	*aMessage = NS_strdup(message);
	return NS_OK;
}

/* readonly attribute nsresult result; */
NS_IMETHODIMP zoteroException::GetResult(nsresult *aResult)
{
	*aResult = NS_ERROR_FAILURE;
	return NS_OK;
}

/* readonly attribute string name; */
NS_IMETHODIMP zoteroException::GetName(char * *aName)
{
	*aName = NS_strdup(message);
	return NS_OK;
}

/* readonly attribute string filename; */
NS_IMETHODIMP zoteroException::GetFilename(char * *aFilename)
{
	*aFilename = NS_strdup(filename);
	return NS_OK;
}

/* readonly attribute PRUint32 lineNumber; */
NS_IMETHODIMP zoteroException::GetLineNumber(PRUint32 *aLineNumber)
{
	*aLineNumber = nsnull;
	return NS_OK;
}

/* readonly attribute PRUint32 columnNumber; */
NS_IMETHODIMP zoteroException::GetColumnNumber(PRUint32 *aColumnNumber)
{
	*aColumnNumber = nsnull;
	return NS_OK;
}

/* readonly attribute nsIStackFrame location; */
NS_IMETHODIMP zoteroException::GetLocation(nsIStackFrame * *aLocation)
{
	*aLocation = nsnull;
	return NS_OK;
}

/* readonly attribute nsIException inner; */
NS_IMETHODIMP zoteroException::GetInner(nsIException * *aInner)
{
	*aInner = nsnull;
	return NS_OK;
}

/* readonly attribute nsISupports data; */
NS_IMETHODIMP zoteroException::GetData(nsISupports * *aData)
{
	*aData = nsnull;
	return NS_OK;
}

/* string toString (); */
NS_IMETHODIMP zoteroException::ToString(char **_retval NS_OUTPARAM)
{
	*_retval = (char *) NS_Alloc(2048);
	_snprintf_s(*_retval, 2048, _TRUNCATE, "[zoteroWinWordIntegration Exception... \"%s\"  code: \"%d\"  function: \"%s\"  location: \"%s\"]", message, errorID, function, filename);
	return NS_OK;
}

// Parses a CException, returning the error message to XPCOM
void zoteroException::report() {
	nsCOMPtr<nsIExceptionService> exceptionService = do_GetService(NS_EXCEPTIONSERVICE_CONTRACTID);
	nsCOMPtr<nsIExceptionManager> exceptionManager;
	nsresult rv = exceptionService->GetCurrentExceptionManager(getter_AddRefs(exceptionManager));
	if(NS_SUCCEEDED(rv)) {
		AddRef();
		exceptionManager->SetCurrentException(this);
	}
}