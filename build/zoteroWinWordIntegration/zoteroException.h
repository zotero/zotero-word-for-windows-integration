#pragma once
#include "stdafx.h"
#include "resource.h"
#ifndef __gen_nsIException_h__
#include "nsIException.h"
#endif

class zoteroException : public nsIException
{
public:
	NS_DECL_ISUPPORTS
	NS_DECL_NSIEXCEPTION
	
	zoteroException();
	zoteroException(char *message, char *aFunction, char *aFilename);
	zoteroException(CException *exception, char *aFunction, char *aFilename);
	void report();
	char *message;
	char *function;
	char *filename;
	uint errorID;
	bool freeMessage;
private:
	~zoteroException();
};

#define ZOTERO_EXCEPTION_CATCHER_START \
	try {

#define ZOTERO_EXCEPTION_CATCHER_END \
	} catch(CException *exception) { \
		zoteroException *zException = new zoteroException(exception, __FUNCTION__, __FILE__); \
		zException->report(); \
		return NS_ERROR_FAILURE; \
	} catch(...) { \
		return NS_ERROR_FAILURE; \
	}

#define ZOTERO_THROW_EXCEPTION(string) \
	zoteroException *zException = new zoteroException(string, __FUNCTION__, __FILE__); \
	zException->report();