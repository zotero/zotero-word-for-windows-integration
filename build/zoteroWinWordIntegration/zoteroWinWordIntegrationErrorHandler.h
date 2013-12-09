/*
 * DO NOT EDIT.  THIS FILE IS GENERATED FROM /cygdrive/c/Users/Simon/Documents/GitHub/zotero-word-for-windows-integration/build/zoteroWinWordIntegrationErrorHandler.idl
 */

#ifndef __gen_zoteroWinWordIntegrationErrorHandler_h__
#define __gen_zoteroWinWordIntegrationErrorHandler_h__


#ifndef __gen_nsISupports_h__
#include "nsISupports.h"
#endif

/* For IDL files that don't want to include root IDL files. */
#ifndef NS_NO_VTABLE
#define NS_NO_VTABLE
#endif

/* starting interface:    zoteroWinWordIntegrationErrorHandler */
#define ZOTEROWINWORDINTEGRATIONERRORHANDLER_IID_STR "fbe7a823-da62-4ca5-b7a4-d0759203b719"

#define ZOTEROWINWORDINTEGRATIONERRORHANDLER_IID \
  {0xfbe7a823, 0xda62, 0x4ca5, \
    { 0xb7, 0xa4, 0xd0, 0x75, 0x92, 0x03, 0xb7, 0x19 }}

class NS_NO_VTABLE zoteroWinWordIntegrationErrorHandler : public nsISupports {
 public: 

  NS_DECLARE_STATIC_IID_ACCESSOR(ZOTEROWINWORDINTEGRATIONERRORHANDLER_IID)

  /* void throwError (in string error); */
  NS_IMETHOD ThrowError(const char * error) = 0;

};

  NS_DEFINE_STATIC_IID_ACCESSOR(zoteroWinWordIntegrationErrorHandler, ZOTEROWINWORDINTEGRATIONERRORHANDLER_IID)

/* Use this macro when declaring classes that implement this interface. */
#define NS_DECL_ZOTEROWINWORDINTEGRATIONERRORHANDLER \
  NS_IMETHOD ThrowError(const char * error); 

/* Use this macro to declare functions that forward the behavior of this interface to another object. */
#define NS_FORWARD_ZOTEROWINWORDINTEGRATIONERRORHANDLER(_to) \
  NS_IMETHOD ThrowError(const char * error) { return _to ThrowError(error); } 

/* Use this macro to declare functions that forward the behavior of this interface to another object in a safe way. */
#define NS_FORWARD_SAFE_ZOTEROWINWORDINTEGRATIONERRORHANDLER(_to) \
  NS_IMETHOD ThrowError(const char * error) { return !_to ? NS_ERROR_NULL_POINTER : _to->ThrowError(error); } 


#endif /* __gen_zoteroWinWordIntegrationErrorHandler_h__ */
