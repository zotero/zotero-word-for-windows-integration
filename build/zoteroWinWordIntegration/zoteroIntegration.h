/*
 * DO NOT EDIT.  THIS FILE IS GENERATED FROM zoteroIntegration.idl
 */

#ifndef __gen_zoteroIntegration_h__
#define __gen_zoteroIntegration_h__


#ifndef __gen_nsISupports_h__
#include "nsISupports.h"
#endif

#ifndef __gen_nsISimpleEnumerator_h__
#include "nsISimpleEnumerator.h"
#endif

#ifndef __gen_nsIObserver_h__
#include "nsIObserver.h"
#endif

/* For IDL files that don't want to include root IDL files. */
#ifndef NS_NO_VTABLE
#define NS_NO_VTABLE
#endif

/* starting interface:    zoteroIntegrationField */
#define ZOTEROINTEGRATIONFIELD_IID_STR "aedb37a0-48bb-11de-8a39-0800200c9a66"

#define ZOTEROINTEGRATIONFIELD_IID \
  {0xaedb37a0, 0x48bb, 0x11de, \
    { 0x8a, 0x39, 0x08, 0x00, 0x20, 0x0c, 0x9a, 0x66 }}

class NS_NO_VTABLE NS_SCRIPTABLE zoteroIntegrationField : public nsISupports {
 public: 

  NS_DECLARE_STATIC_IID_ACCESSOR(ZOTEROINTEGRATIONFIELD_IID)

  /* void delete (); */
  NS_SCRIPTABLE NS_IMETHOD Delete(void) = 0;

  /* void select (); */
  NS_SCRIPTABLE NS_IMETHOD Select(void) = 0;

  /* void removeCode (); */
  NS_SCRIPTABLE NS_IMETHOD RemoveCode(void) = 0;

  /* wstring getText (); */
  NS_SCRIPTABLE NS_IMETHOD GetText(PRUnichar * *_retval NS_OUTPARAM) = 0;

  /* void setText (in wstring text, in boolean isRich); */
  NS_SCRIPTABLE NS_IMETHOD SetText(const PRUnichar * text, bool isRich) = 0;

  /* wstring getCode (); */
  NS_SCRIPTABLE NS_IMETHOD GetCode(PRUnichar * *_retval NS_OUTPARAM) = 0;

  /* void setCode (in wstring code); */
  NS_SCRIPTABLE NS_IMETHOD SetCode(const PRUnichar * code) = 0;

  /* unsigned long getNoteIndex (); */
  NS_SCRIPTABLE NS_IMETHOD GetNoteIndex(PRUint32 *_retval NS_OUTPARAM) = 0;

  /* boolean equals (in zoteroIntegrationField field); */
  NS_SCRIPTABLE NS_IMETHOD Equals(zoteroIntegrationField *field, bool *_retval NS_OUTPARAM) = 0;

  /* zoteroIntegrationField getNextField (); */
  NS_SCRIPTABLE NS_IMETHOD GetNextField(zoteroIntegrationField * *_retval NS_OUTPARAM) = 0;

  /* zoteroIntegrationField getPreviousField (); */
  NS_SCRIPTABLE NS_IMETHOD GetPreviousField(zoteroIntegrationField * *_retval NS_OUTPARAM) = 0;

};

  NS_DEFINE_STATIC_IID_ACCESSOR(zoteroIntegrationField, ZOTEROINTEGRATIONFIELD_IID)

/* Use this macro when declaring classes that implement this interface. */
#define NS_DECL_ZOTEROINTEGRATIONFIELD \
  NS_SCRIPTABLE NS_IMETHOD Delete(void); \
  NS_SCRIPTABLE NS_IMETHOD Select(void); \
  NS_SCRIPTABLE NS_IMETHOD RemoveCode(void); \
  NS_SCRIPTABLE NS_IMETHOD GetText(PRUnichar * *_retval NS_OUTPARAM); \
  NS_SCRIPTABLE NS_IMETHOD SetText(const PRUnichar * text, bool isRich); \
  NS_SCRIPTABLE NS_IMETHOD GetCode(PRUnichar * *_retval NS_OUTPARAM); \
  NS_SCRIPTABLE NS_IMETHOD SetCode(const PRUnichar * code); \
  NS_SCRIPTABLE NS_IMETHOD GetNoteIndex(PRUint32 *_retval NS_OUTPARAM); \
  NS_SCRIPTABLE NS_IMETHOD Equals(zoteroIntegrationField *field, bool *_retval NS_OUTPARAM); \
  NS_SCRIPTABLE NS_IMETHOD GetNextField(zoteroIntegrationField * *_retval NS_OUTPARAM); \
  NS_SCRIPTABLE NS_IMETHOD GetPreviousField(zoteroIntegrationField * *_retval NS_OUTPARAM); 

/* Use this macro to declare functions that forward the behavior of this interface to another object. */
#define NS_FORWARD_ZOTEROINTEGRATIONFIELD(_to) \
  NS_SCRIPTABLE NS_IMETHOD Delete(void) { return _to Delete(); } \
  NS_SCRIPTABLE NS_IMETHOD Select(void) { return _to Select(); } \
  NS_SCRIPTABLE NS_IMETHOD RemoveCode(void) { return _to RemoveCode(); } \
  NS_SCRIPTABLE NS_IMETHOD GetText(PRUnichar * *_retval NS_OUTPARAM) { return _to GetText(_retval); } \
  NS_SCRIPTABLE NS_IMETHOD SetText(const PRUnichar * text, bool isRich) { return _to SetText(text, isRich); } \
  NS_SCRIPTABLE NS_IMETHOD GetCode(PRUnichar * *_retval NS_OUTPARAM) { return _to GetCode(_retval); } \
  NS_SCRIPTABLE NS_IMETHOD SetCode(const PRUnichar * code) { return _to SetCode(code); } \
  NS_SCRIPTABLE NS_IMETHOD GetNoteIndex(PRUint32 *_retval NS_OUTPARAM) { return _to GetNoteIndex(_retval); } \
  NS_SCRIPTABLE NS_IMETHOD Equals(zoteroIntegrationField *field, bool *_retval NS_OUTPARAM) { return _to Equals(field, _retval); } \
  NS_SCRIPTABLE NS_IMETHOD GetNextField(zoteroIntegrationField * *_retval NS_OUTPARAM) { return _to GetNextField(_retval); } \
  NS_SCRIPTABLE NS_IMETHOD GetPreviousField(zoteroIntegrationField * *_retval NS_OUTPARAM) { return _to GetPreviousField(_retval); } 

/* Use this macro to declare functions that forward the behavior of this interface to another object in a safe way. */
#define NS_FORWARD_SAFE_ZOTEROINTEGRATIONFIELD(_to) \
  NS_SCRIPTABLE NS_IMETHOD Delete(void) { return !_to ? NS_ERROR_NULL_POINTER : _to->Delete(); } \
  NS_SCRIPTABLE NS_IMETHOD Select(void) { return !_to ? NS_ERROR_NULL_POINTER : _to->Select(); } \
  NS_SCRIPTABLE NS_IMETHOD RemoveCode(void) { return !_to ? NS_ERROR_NULL_POINTER : _to->RemoveCode(); } \
  NS_SCRIPTABLE NS_IMETHOD GetText(PRUnichar * *_retval NS_OUTPARAM) { return !_to ? NS_ERROR_NULL_POINTER : _to->GetText(_retval); } \
  NS_SCRIPTABLE NS_IMETHOD SetText(const PRUnichar * text, bool isRich) { return !_to ? NS_ERROR_NULL_POINTER : _to->SetText(text, isRich); } \
  NS_SCRIPTABLE NS_IMETHOD GetCode(PRUnichar * *_retval NS_OUTPARAM) { return !_to ? NS_ERROR_NULL_POINTER : _to->GetCode(_retval); } \
  NS_SCRIPTABLE NS_IMETHOD SetCode(const PRUnichar * code) { return !_to ? NS_ERROR_NULL_POINTER : _to->SetCode(code); } \
  NS_SCRIPTABLE NS_IMETHOD GetNoteIndex(PRUint32 *_retval NS_OUTPARAM) { return !_to ? NS_ERROR_NULL_POINTER : _to->GetNoteIndex(_retval); } \
  NS_SCRIPTABLE NS_IMETHOD Equals(zoteroIntegrationField *field, bool *_retval NS_OUTPARAM) { return !_to ? NS_ERROR_NULL_POINTER : _to->Equals(field, _retval); } \
  NS_SCRIPTABLE NS_IMETHOD GetNextField(zoteroIntegrationField * *_retval NS_OUTPARAM) { return !_to ? NS_ERROR_NULL_POINTER : _to->GetNextField(_retval); } \
  NS_SCRIPTABLE NS_IMETHOD GetPreviousField(zoteroIntegrationField * *_retval NS_OUTPARAM) { return !_to ? NS_ERROR_NULL_POINTER : _to->GetPreviousField(_retval); } 

#if 0
/* Use the code below as a template for the implementation class for this interface. */

/* Header file */
class _MYCLASS_ : public zoteroIntegrationField
{
public:
  NS_DECL_ISUPPORTS
  NS_DECL_ZOTEROINTEGRATIONFIELD

  _MYCLASS_();

private:
  ~_MYCLASS_();

protected:
  /* additional members */
};

/* Implementation file */
NS_IMPL_ISUPPORTS1(_MYCLASS_, zoteroIntegrationField)

_MYCLASS_::_MYCLASS_()
{
  /* member initializers and constructor code */
}

_MYCLASS_::~_MYCLASS_()
{
  /* destructor code */
}

/* void delete (); */
NS_IMETHODIMP _MYCLASS_::Delete()
{
    return NS_ERROR_NOT_IMPLEMENTED;
}

/* void select (); */
NS_IMETHODIMP _MYCLASS_::Select()
{
    return NS_ERROR_NOT_IMPLEMENTED;
}

/* void removeCode (); */
NS_IMETHODIMP _MYCLASS_::RemoveCode()
{
    return NS_ERROR_NOT_IMPLEMENTED;
}

/* wstring getText (); */
NS_IMETHODIMP _MYCLASS_::GetText(PRUnichar * *_retval NS_OUTPARAM)
{
    return NS_ERROR_NOT_IMPLEMENTED;
}

/* void setText (in wstring text, in boolean isRich); */
NS_IMETHODIMP _MYCLASS_::SetText(const PRUnichar * text, bool isRich)
{
    return NS_ERROR_NOT_IMPLEMENTED;
}

/* wstring getCode (); */
NS_IMETHODIMP _MYCLASS_::GetCode(PRUnichar * *_retval NS_OUTPARAM)
{
    return NS_ERROR_NOT_IMPLEMENTED;
}

/* void setCode (in wstring code); */
NS_IMETHODIMP _MYCLASS_::SetCode(const PRUnichar * code)
{
    return NS_ERROR_NOT_IMPLEMENTED;
}

/* unsigned long getNoteIndex (); */
NS_IMETHODIMP _MYCLASS_::GetNoteIndex(PRUint32 *_retval NS_OUTPARAM)
{
    return NS_ERROR_NOT_IMPLEMENTED;
}

/* boolean equals (in zoteroIntegrationField field); */
NS_IMETHODIMP _MYCLASS_::Equals(zoteroIntegrationField *field, bool *_retval NS_OUTPARAM)
{
    return NS_ERROR_NOT_IMPLEMENTED;
}

/* zoteroIntegrationField getNextField (); */
NS_IMETHODIMP _MYCLASS_::GetNextField(zoteroIntegrationField * *_retval NS_OUTPARAM)
{
    return NS_ERROR_NOT_IMPLEMENTED;
}

/* zoteroIntegrationField getPreviousField (); */
NS_IMETHODIMP _MYCLASS_::GetPreviousField(zoteroIntegrationField * *_retval NS_OUTPARAM)
{
    return NS_ERROR_NOT_IMPLEMENTED;
}

/* End of implementation class template. */
#endif


/* starting interface:    zoteroIntegrationDocument */
#define ZOTEROINTEGRATIONDOCUMENT_IID_STR "be1c5c1f-9ed2-4154-98fb-822d1fede569"

#define ZOTEROINTEGRATIONDOCUMENT_IID \
  {0xbe1c5c1f, 0x9ed2, 0x4154, \
    { 0x98, 0xfb, 0x82, 0x2d, 0x1f, 0xed, 0xe5, 0x69 }}

class NS_NO_VTABLE NS_SCRIPTABLE zoteroIntegrationDocument : public nsISupports {
 public: 

  NS_DECLARE_STATIC_IID_ACCESSOR(ZOTEROINTEGRATIONDOCUMENT_IID)

  /* short displayAlert (in wstring dialogText, in unsigned short icon, in unsigned short buttons); */
  NS_SCRIPTABLE NS_IMETHOD DisplayAlert(const PRUnichar * dialogText, PRUint16 icon, PRUint16 buttons, PRInt16 *_retval NS_OUTPARAM) = 0;

  /* void activate (); */
  NS_SCRIPTABLE NS_IMETHOD Activate(void) = 0;

  /* boolean canInsertField (in string fieldType); */
  NS_SCRIPTABLE NS_IMETHOD CanInsertField(const char * fieldType, bool *_retval NS_OUTPARAM) = 0;

  /* zoteroIntegrationField cursorInField (in string fieldType); */
  NS_SCRIPTABLE NS_IMETHOD CursorInField(const char * fieldType, zoteroIntegrationField * *_retval NS_OUTPARAM) = 0;

  /* wstring getDocumentData (); */
  NS_SCRIPTABLE NS_IMETHOD GetDocumentData(PRUnichar * *_retval NS_OUTPARAM) = 0;

  /* void setDocumentData (in wstring data); */
  NS_SCRIPTABLE NS_IMETHOD SetDocumentData(const PRUnichar * data) = 0;

  /* zoteroIntegrationField insertField (in string fieldType, in unsigned short noteType); */
  NS_SCRIPTABLE NS_IMETHOD InsertField(const char * fieldType, PRUint16 noteType, zoteroIntegrationField * *_retval NS_OUTPARAM) = 0;

  /* nsISimpleEnumerator getFields (in string fieldType); */
  NS_SCRIPTABLE NS_IMETHOD GetFields(const char * fieldType, nsISimpleEnumerator * *_retval NS_OUTPARAM) = 0;

  /* void getFieldsAsync (in string fieldType, in nsIObserver observer); */
  NS_SCRIPTABLE NS_IMETHOD GetFieldsAsync(const char * fieldType, nsIObserver *observer) = 0;

  /* void convert (in nsISimpleEnumerator fields, in string toFieldType, [array, size_is (count)] in unsigned short toNoteType, in unsigned long count); */
  NS_SCRIPTABLE NS_IMETHOD Convert(nsISimpleEnumerator *fields, const char * toFieldType, PRUint16 *toNoteType, PRUint32 count) = 0;

  /* void setBibliographyStyle (in long firstLineIndent, in long bodyIndent, in unsigned long lineSpacing, in unsigned long entrySpacing, [array, size_is (tabStopCount)] in long tabStops, in unsigned long tabStopCount); */
  NS_SCRIPTABLE NS_IMETHOD SetBibliographyStyle(PRInt32 firstLineIndent, PRInt32 bodyIndent, PRUint32 lineSpacing, PRUint32 entrySpacing, PRInt32 *tabStops, PRUint32 tabStopCount) = 0;

  /* void cleanup (); */
  NS_SCRIPTABLE NS_IMETHOD Cleanup(void) = 0;

  enum {
    DIALOG_ICON_STOP = 0U,
    DIALOG_ICON_NOTICE = 1U,
    DIALOG_ICON_CAUTION = 2U,
    DIALOG_BUTTONS_OK = 0U,
    DIALOG_BUTTONS_OK_CANCEL = 1U,
    DIALOG_BUTTONS_YES_NO = 2U,
    DIALOG_BUTTONS_YES_NO_CANCEL = 3U,
    NOTE_FOOTNOTE = 1U,
    NOTE_ENDNOTE = 2U
  };

};

  NS_DEFINE_STATIC_IID_ACCESSOR(zoteroIntegrationDocument, ZOTEROINTEGRATIONDOCUMENT_IID)

/* Use this macro when declaring classes that implement this interface. */
#define NS_DECL_ZOTEROINTEGRATIONDOCUMENT \
  NS_SCRIPTABLE NS_IMETHOD DisplayAlert(const PRUnichar * dialogText, PRUint16 icon, PRUint16 buttons, PRInt16 *_retval NS_OUTPARAM); \
  NS_SCRIPTABLE NS_IMETHOD Activate(void); \
  NS_SCRIPTABLE NS_IMETHOD CanInsertField(const char * fieldType, bool *_retval NS_OUTPARAM); \
  NS_SCRIPTABLE NS_IMETHOD CursorInField(const char * fieldType, zoteroIntegrationField * *_retval NS_OUTPARAM); \
  NS_SCRIPTABLE NS_IMETHOD GetDocumentData(PRUnichar * *_retval NS_OUTPARAM); \
  NS_SCRIPTABLE NS_IMETHOD SetDocumentData(const PRUnichar * data); \
  NS_SCRIPTABLE NS_IMETHOD InsertField(const char * fieldType, PRUint16 noteType, zoteroIntegrationField * *_retval NS_OUTPARAM); \
  NS_SCRIPTABLE NS_IMETHOD GetFields(const char * fieldType, nsISimpleEnumerator * *_retval NS_OUTPARAM); \
  NS_SCRIPTABLE NS_IMETHOD GetFieldsAsync(const char * fieldType, nsIObserver *observer); \
  NS_SCRIPTABLE NS_IMETHOD Convert(nsISimpleEnumerator *fields, const char * toFieldType, PRUint16 *toNoteType, PRUint32 count); \
  NS_SCRIPTABLE NS_IMETHOD SetBibliographyStyle(PRInt32 firstLineIndent, PRInt32 bodyIndent, PRUint32 lineSpacing, PRUint32 entrySpacing, PRInt32 *tabStops, PRUint32 tabStopCount); \
  NS_SCRIPTABLE NS_IMETHOD Cleanup(void); \

/* Use this macro to declare functions that forward the behavior of this interface to another object. */
#define NS_FORWARD_ZOTEROINTEGRATIONDOCUMENT(_to) \
  NS_SCRIPTABLE NS_IMETHOD DisplayAlert(const PRUnichar * dialogText, PRUint16 icon, PRUint16 buttons, PRInt16 *_retval NS_OUTPARAM) { return _to DisplayAlert(dialogText, icon, buttons, _retval); } \
  NS_SCRIPTABLE NS_IMETHOD Activate(void) { return _to Activate(); } \
  NS_SCRIPTABLE NS_IMETHOD CanInsertField(const char * fieldType, bool *_retval NS_OUTPARAM) { return _to CanInsertField(fieldType, _retval); } \
  NS_SCRIPTABLE NS_IMETHOD CursorInField(const char * fieldType, zoteroIntegrationField * *_retval NS_OUTPARAM) { return _to CursorInField(fieldType, _retval); } \
  NS_SCRIPTABLE NS_IMETHOD GetDocumentData(PRUnichar * *_retval NS_OUTPARAM) { return _to GetDocumentData(_retval); } \
  NS_SCRIPTABLE NS_IMETHOD SetDocumentData(const PRUnichar * data) { return _to SetDocumentData(data); } \
  NS_SCRIPTABLE NS_IMETHOD InsertField(const char * fieldType, PRUint16 noteType, zoteroIntegrationField * *_retval NS_OUTPARAM) { return _to InsertField(fieldType, noteType, _retval); } \
  NS_SCRIPTABLE NS_IMETHOD GetFields(const char * fieldType, nsISimpleEnumerator * *_retval NS_OUTPARAM) { return _to GetFields(fieldType, _retval); } \
  NS_SCRIPTABLE NS_IMETHOD GetFieldsAsync(const char * fieldType, nsIObserver *observer) { return _to GetFieldsAsync(fieldType, observer); } \
  NS_SCRIPTABLE NS_IMETHOD Convert(nsISimpleEnumerator *fields, const char * toFieldType, PRUint16 *toNoteType, PRUint32 count) { return _to Convert(fields, toFieldType, toNoteType, count); } \
  NS_SCRIPTABLE NS_IMETHOD SetBibliographyStyle(PRInt32 firstLineIndent, PRInt32 bodyIndent, PRUint32 lineSpacing, PRUint32 entrySpacing, PRInt32 *tabStops, PRUint32 tabStopCount) { return _to SetBibliographyStyle(firstLineIndent, bodyIndent, lineSpacing, entrySpacing, tabStops, tabStopCount); } \
  NS_SCRIPTABLE NS_IMETHOD Cleanup(void) { return _to Cleanup(); } \

/* Use this macro to declare functions that forward the behavior of this interface to another object in a safe way. */
#define NS_FORWARD_SAFE_ZOTEROINTEGRATIONDOCUMENT(_to) \
  NS_SCRIPTABLE NS_IMETHOD DisplayAlert(const PRUnichar * dialogText, PRUint16 icon, PRUint16 buttons, PRInt16 *_retval NS_OUTPARAM) { return !_to ? NS_ERROR_NULL_POINTER : _to->DisplayAlert(dialogText, icon, buttons, _retval); } \
  NS_SCRIPTABLE NS_IMETHOD Activate(void) { return !_to ? NS_ERROR_NULL_POINTER : _to->Activate(); } \
  NS_SCRIPTABLE NS_IMETHOD CanInsertField(const char * fieldType, bool *_retval NS_OUTPARAM) { return !_to ? NS_ERROR_NULL_POINTER : _to->CanInsertField(fieldType, _retval); } \
  NS_SCRIPTABLE NS_IMETHOD CursorInField(const char * fieldType, zoteroIntegrationField * *_retval NS_OUTPARAM) { return !_to ? NS_ERROR_NULL_POINTER : _to->CursorInField(fieldType, _retval); } \
  NS_SCRIPTABLE NS_IMETHOD GetDocumentData(PRUnichar * *_retval NS_OUTPARAM) { return !_to ? NS_ERROR_NULL_POINTER : _to->GetDocumentData(_retval); } \
  NS_SCRIPTABLE NS_IMETHOD SetDocumentData(const PRUnichar * data) { return !_to ? NS_ERROR_NULL_POINTER : _to->SetDocumentData(data); } \
  NS_SCRIPTABLE NS_IMETHOD InsertField(const char * fieldType, PRUint16 noteType, zoteroIntegrationField * *_retval NS_OUTPARAM) { return !_to ? NS_ERROR_NULL_POINTER : _to->InsertField(fieldType, noteType, _retval); } \
  NS_SCRIPTABLE NS_IMETHOD GetFields(const char * fieldType, nsISimpleEnumerator * *_retval NS_OUTPARAM) { return !_to ? NS_ERROR_NULL_POINTER : _to->GetFields(fieldType, _retval); } \
  NS_SCRIPTABLE NS_IMETHOD GetFieldsAsync(const char * fieldType, nsIObserver *observer) { return !_to ? NS_ERROR_NULL_POINTER : _to->GetFieldsAsync(fieldType, observer); } \
  NS_SCRIPTABLE NS_IMETHOD Convert(nsISimpleEnumerator *fields, const char * toFieldType, PRUint16 *toNoteType, PRUint32 count) { return !_to ? NS_ERROR_NULL_POINTER : _to->Convert(fields, toFieldType, toNoteType, count); } \
  NS_SCRIPTABLE NS_IMETHOD SetBibliographyStyle(PRInt32 firstLineIndent, PRInt32 bodyIndent, PRUint32 lineSpacing, PRUint32 entrySpacing, PRInt32 *tabStops, PRUint32 tabStopCount) { return !_to ? NS_ERROR_NULL_POINTER : _to->SetBibliographyStyle(firstLineIndent, bodyIndent, lineSpacing, entrySpacing, tabStops, tabStopCount); } \
  NS_SCRIPTABLE NS_IMETHOD Cleanup(void) { return !_to ? NS_ERROR_NULL_POINTER : _to->Cleanup(); } \

#if 0
/* Use the code below as a template for the implementation class for this interface. */

/* Header file */
class _MYCLASS_ : public zoteroIntegrationDocument
{
public:
  NS_DECL_ISUPPORTS
  NS_DECL_ZOTEROINTEGRATIONDOCUMENT

  _MYCLASS_();

private:
  ~_MYCLASS_();

protected:
  /* additional members */
};

/* Implementation file */
NS_IMPL_ISUPPORTS1(_MYCLASS_, zoteroIntegrationDocument)

_MYCLASS_::_MYCLASS_()
{
  /* member initializers and constructor code */
}

_MYCLASS_::~_MYCLASS_()
{
  /* destructor code */
}

/* short displayAlert (in wstring dialogText, in unsigned short icon, in unsigned short buttons); */
NS_IMETHODIMP _MYCLASS_::DisplayAlert(const PRUnichar * dialogText, PRUint16 icon, PRUint16 buttons, PRInt16 *_retval NS_OUTPARAM)
{
    return NS_ERROR_NOT_IMPLEMENTED;
}

/* void activate (); */
NS_IMETHODIMP _MYCLASS_::Activate()
{
    return NS_ERROR_NOT_IMPLEMENTED;
}

/* boolean canInsertField (in string fieldType); */
NS_IMETHODIMP _MYCLASS_::CanInsertField(const char * fieldType, bool *_retval NS_OUTPARAM)
{
    return NS_ERROR_NOT_IMPLEMENTED;
}

/* zoteroIntegrationField cursorInField (in string fieldType); */
NS_IMETHODIMP _MYCLASS_::CursorInField(const char * fieldType, zoteroIntegrationField * *_retval NS_OUTPARAM)
{
    return NS_ERROR_NOT_IMPLEMENTED;
}

/* wstring getDocumentData (); */
NS_IMETHODIMP _MYCLASS_::GetDocumentData(PRUnichar * *_retval NS_OUTPARAM)
{
    return NS_ERROR_NOT_IMPLEMENTED;
}

/* void setDocumentData (in wstring data); */
NS_IMETHODIMP _MYCLASS_::SetDocumentData(const PRUnichar * data)
{
    return NS_ERROR_NOT_IMPLEMENTED;
}

/* zoteroIntegrationField insertField (in string fieldType, in unsigned short noteType); */
NS_IMETHODIMP _MYCLASS_::InsertField(const char * fieldType, PRUint16 noteType, zoteroIntegrationField * *_retval NS_OUTPARAM)
{
    return NS_ERROR_NOT_IMPLEMENTED;
}

/* nsISimpleEnumerator getFields (in string fieldType); */
NS_IMETHODIMP _MYCLASS_::GetFields(const char * fieldType, nsISimpleEnumerator * *_retval NS_OUTPARAM)
{
    return NS_ERROR_NOT_IMPLEMENTED;
}

/* void getFieldsAsync (in string fieldType, in nsIObserver observer); */
NS_IMETHODIMP _MYCLASS_::GetFieldsAsync(const char * fieldType, nsIObserver *observer)
{
    return NS_ERROR_NOT_IMPLEMENTED;
}

/* void convert (in nsISimpleEnumerator fields, in string toFieldType, [array, size_is (count)] in unsigned short toNoteType, in unsigned long count); */
NS_IMETHODIMP _MYCLASS_::Convert(nsISimpleEnumerator *fields, const char * toFieldType, PRUint16 *toNoteType, PRUint32 count)
{
    return NS_ERROR_NOT_IMPLEMENTED;
}

/* void setBibliographyStyle (in long firstLineIndent, in long bodyIndent, in unsigned long lineSpacing, in unsigned long entrySpacing, [array, size_is (tabStopCount)] in long tabStops, in unsigned long tabStopCount); */
NS_IMETHODIMP _MYCLASS_::SetBibliographyStyle(PRInt32 firstLineIndent, PRInt32 bodyIndent, PRUint32 lineSpacing, PRUint32 entrySpacing, PRInt32 *tabStops, PRUint32 tabStopCount)
{
    return NS_ERROR_NOT_IMPLEMENTED;
}

/* void cleanup (); */
NS_IMETHODIMP _MYCLASS_::Cleanup()
{
    return NS_ERROR_NOT_IMPLEMENTED;
}

/* End of implementation class template. */
#endif


/* starting interface:    zoteroIntegrationApplication */
#define ZOTEROINTEGRATIONAPPLICATION_IID_STR "7b258e57-20cf-4a73-8420-5d06a538c25e"

#define ZOTEROINTEGRATIONAPPLICATION_IID \
  {0x7b258e57, 0x20cf, 0x4a73, \
    { 0x84, 0x20, 0x5d, 0x06, 0xa5, 0x38, 0xc2, 0x5e }}

class NS_NO_VTABLE NS_SCRIPTABLE zoteroIntegrationApplication : public nsISupports {
 public: 

  NS_DECLARE_STATIC_IID_ACCESSOR(ZOTEROINTEGRATIONAPPLICATION_IID)

  /* readonly attribute ACString primaryFieldType; */
  NS_SCRIPTABLE NS_IMETHOD GetPrimaryFieldType(nsACString & aPrimaryFieldType) = 0;

  /* readonly attribute ACString secondaryFieldType; */
  NS_SCRIPTABLE NS_IMETHOD GetSecondaryFieldType(nsACString & aSecondaryFieldType) = 0;

  /* zoteroIntegrationDocument getActiveDocument (); */
  NS_SCRIPTABLE NS_IMETHOD GetActiveDocument(zoteroIntegrationDocument * *_retval NS_OUTPARAM) = 0;

  /* zoteroIntegrationDocument getDocument (in wstring documentIdentifier); */
  NS_SCRIPTABLE NS_IMETHOD GetDocument(const PRUnichar * documentIdentifier, zoteroIntegrationDocument * *_retval NS_OUTPARAM) = 0;

};

  NS_DEFINE_STATIC_IID_ACCESSOR(zoteroIntegrationApplication, ZOTEROINTEGRATIONAPPLICATION_IID)

/* Use this macro when declaring classes that implement this interface. */
#define NS_DECL_ZOTEROINTEGRATIONAPPLICATION \
  NS_SCRIPTABLE NS_IMETHOD GetPrimaryFieldType(nsACString & aPrimaryFieldType); \
  NS_SCRIPTABLE NS_IMETHOD GetSecondaryFieldType(nsACString & aSecondaryFieldType); \
  NS_SCRIPTABLE NS_IMETHOD GetActiveDocument(zoteroIntegrationDocument * *_retval NS_OUTPARAM); \
  NS_SCRIPTABLE NS_IMETHOD GetDocument(const PRUnichar * documentIdentifier, zoteroIntegrationDocument * *_retval NS_OUTPARAM); 

/* Use this macro to declare functions that forward the behavior of this interface to another object. */
#define NS_FORWARD_ZOTEROINTEGRATIONAPPLICATION(_to) \
  NS_SCRIPTABLE NS_IMETHOD GetPrimaryFieldType(nsACString & aPrimaryFieldType) { return _to GetPrimaryFieldType(aPrimaryFieldType); } \
  NS_SCRIPTABLE NS_IMETHOD GetSecondaryFieldType(nsACString & aSecondaryFieldType) { return _to GetSecondaryFieldType(aSecondaryFieldType); } \
  NS_SCRIPTABLE NS_IMETHOD GetActiveDocument(zoteroIntegrationDocument * *_retval NS_OUTPARAM) { return _to GetActiveDocument(_retval); } \
  NS_SCRIPTABLE NS_IMETHOD GetDocument(const PRUnichar * documentIdentifier, zoteroIntegrationDocument * *_retval NS_OUTPARAM) { return _to GetDocument(documentIdentifier, _retval); } 

/* Use this macro to declare functions that forward the behavior of this interface to another object in a safe way. */
#define NS_FORWARD_SAFE_ZOTEROINTEGRATIONAPPLICATION(_to) \
  NS_SCRIPTABLE NS_IMETHOD GetPrimaryFieldType(nsACString & aPrimaryFieldType) { return !_to ? NS_ERROR_NULL_POINTER : _to->GetPrimaryFieldType(aPrimaryFieldType); } \
  NS_SCRIPTABLE NS_IMETHOD GetSecondaryFieldType(nsACString & aSecondaryFieldType) { return !_to ? NS_ERROR_NULL_POINTER : _to->GetSecondaryFieldType(aSecondaryFieldType); } \
  NS_SCRIPTABLE NS_IMETHOD GetActiveDocument(zoteroIntegrationDocument * *_retval NS_OUTPARAM) { return !_to ? NS_ERROR_NULL_POINTER : _to->GetActiveDocument(_retval); } \
  NS_SCRIPTABLE NS_IMETHOD GetDocument(const PRUnichar * documentIdentifier, zoteroIntegrationDocument * *_retval NS_OUTPARAM) { return !_to ? NS_ERROR_NULL_POINTER : _to->GetDocument(documentIdentifier, _retval); } 

#if 0
/* Use the code below as a template for the implementation class for this interface. */

/* Header file */
class _MYCLASS_ : public zoteroIntegrationApplication
{
public:
  NS_DECL_ISUPPORTS
  NS_DECL_ZOTEROINTEGRATIONAPPLICATION

  _MYCLASS_();

private:
  ~_MYCLASS_();

protected:
  /* additional members */
};

/* Implementation file */
NS_IMPL_ISUPPORTS1(_MYCLASS_, zoteroIntegrationApplication)

_MYCLASS_::_MYCLASS_()
{
  /* member initializers and constructor code */
}

_MYCLASS_::~_MYCLASS_()
{
  /* destructor code */
}

/* readonly attribute ACString primaryFieldType; */
NS_IMETHODIMP _MYCLASS_::GetPrimaryFieldType(nsACString & aPrimaryFieldType)
{
    return NS_ERROR_NOT_IMPLEMENTED;
}

/* readonly attribute ACString secondaryFieldType; */
NS_IMETHODIMP _MYCLASS_::GetSecondaryFieldType(nsACString & aSecondaryFieldType)
{
    return NS_ERROR_NOT_IMPLEMENTED;
}

/* zoteroIntegrationDocument getActiveDocument (); */
NS_IMETHODIMP _MYCLASS_::GetActiveDocument(zoteroIntegrationDocument * *_retval NS_OUTPARAM)
{
    return NS_ERROR_NOT_IMPLEMENTED;
}

/* zoteroIntegrationDocument getDocument (in wstring documentIdentifier); */
NS_IMETHODIMP _MYCLASS_::GetDocument(const PRUnichar * documentIdentifier, zoteroIntegrationDocument * *_retval NS_OUTPARAM)
{
    return NS_ERROR_NOT_IMPLEMENTED;
}

/* End of implementation class template. */
#endif


#endif /* __gen_zoteroIntegration_h__ */
