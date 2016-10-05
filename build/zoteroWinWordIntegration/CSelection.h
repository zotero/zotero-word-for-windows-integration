// Machine generated IDispatch wrapper class(es) created with Add Class from Typelib Wizard
#include "imports.h"

// CSelection wrapper class

class CSelection : public COleDispatchDriver
{
public:
	CSelection(){} // Calls COleDispatchDriver default constructor
	CSelection(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CSelection(const CSelection& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// Attributes
public:

	// Operations
public:


	// Selection methods
public:
	CString get_Text()
	{
		CString result;
		InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_Text(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x0, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_FormattedText()
	{
		LPDISPATCH result;
		InvokeHelper(0x2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void put_FormattedText(LPDISPATCH newValue)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x2, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
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
	LPDISPATCH get_Font()
	{
		LPDISPATCH result;
		InvokeHelper(0x5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void put_Font(LPDISPATCH newValue)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x5, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_Type()
	{
		long result;
		InvokeHelper(0x6, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	long get_StoryType()
	{
		long result;
		InvokeHelper(0x7, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	VARIANT get_Style()
	{
		VARIANT result;
		InvokeHelper(0x8, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
		return result;
	}
	void put_Style(VARIANT * newValue)
	{
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x8, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_Tables()
	{
		LPDISPATCH result;
		InvokeHelper(0x32, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Words()
	{
		LPDISPATCH result;
		InvokeHelper(0x33, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Sentences()
	{
		LPDISPATCH result;
		InvokeHelper(0x34, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Characters()
	{
		LPDISPATCH result;
		InvokeHelper(0x35, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Footnotes()
	{
		LPDISPATCH result;
		InvokeHelper(0x36, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Endnotes()
	{
		LPDISPATCH result;
		InvokeHelper(0x37, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Comments()
	{
		LPDISPATCH result;
		InvokeHelper(0x38, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Cells()
	{
		LPDISPATCH result;
		InvokeHelper(0x39, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Sections()
	{
		LPDISPATCH result;
		InvokeHelper(0x3a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Paragraphs()
	{
		LPDISPATCH result;
		InvokeHelper(0x3b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Borders()
	{
		LPDISPATCH result;
		InvokeHelper(0x44c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void put_Borders(LPDISPATCH newValue)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x44c, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_Shading()
	{
		LPDISPATCH result;
		InvokeHelper(0x3d, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Fields()
	{
		LPDISPATCH result;
		InvokeHelper(0x40, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_FormFields()
	{
		LPDISPATCH result;
		InvokeHelper(0x41, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Frames()
	{
		LPDISPATCH result;
		InvokeHelper(0x42, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_ParagraphFormat()
	{
		LPDISPATCH result;
		InvokeHelper(0x44e, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void put_ParagraphFormat(LPDISPATCH newValue)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x44e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_PageSetup()
	{
		LPDISPATCH result;
		InvokeHelper(0x44d, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void put_PageSetup(LPDISPATCH newValue)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x44d, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_Bookmarks()
	{
		LPDISPATCH result;
		InvokeHelper(0x4b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_StoryLength()
	{
		long result;
		InvokeHelper(0x98, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	long get_LanguageID()
	{
		long result;
		InvokeHelper(0x99, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_LanguageID(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x99, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_LanguageIDFarEast()
	{
		long result;
		InvokeHelper(0x9a, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_LanguageIDFarEast(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x9a, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_LanguageIDOther()
	{
		long result;
		InvokeHelper(0x9b, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_LanguageIDOther(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x9b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_Hyperlinks()
	{
		LPDISPATCH result;
		InvokeHelper(0x9c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Columns()
	{
		LPDISPATCH result;
		InvokeHelper(0x12e, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Rows()
	{
		LPDISPATCH result;
		InvokeHelper(0x12f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_HeaderFooter()
	{
		LPDISPATCH result;
		InvokeHelper(0x132, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	BOOL get_IsEndOfRowMark()
	{
		BOOL result;
		InvokeHelper(0x133, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	long get_BookmarkID()
	{
		long result;
		InvokeHelper(0x134, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	long get_PreviousBookmarkID()
	{
		long result;
		InvokeHelper(0x135, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Find()
	{
		LPDISPATCH result;
		InvokeHelper(0x106, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Range()
	{
		LPDISPATCH result;
		InvokeHelper(0x190, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	VARIANT get_Information(long Type)
	{
		VARIANT result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x191, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, parms, Type);
		return result;
	}
	long get_Flags()
	{
		long result;
		InvokeHelper(0x192, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_Flags(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x192, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_Active()
	{
		BOOL result;
		InvokeHelper(0x193, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	BOOL get_StartIsActive()
	{
		BOOL result;
		InvokeHelper(0x194, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_StartIsActive(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x194, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_IPAtEndOfLine()
	{
		BOOL result;
		InvokeHelper(0x195, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	BOOL get_ExtendMode()
	{
		BOOL result;
		InvokeHelper(0x196, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_ExtendMode(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x196, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_ColumnSelectMode()
	{
		BOOL result;
		InvokeHelper(0x197, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_ColumnSelectMode(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x197, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_Orientation()
	{
		long result;
		InvokeHelper(0x19a, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_Orientation(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x19a, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_InlineShapes()
	{
		LPDISPATCH result;
		InvokeHelper(0x19b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
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
	LPDISPATCH get_Document()
	{
		LPDISPATCH result;
		InvokeHelper(0x3eb, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_ShapeRange()
	{
		LPDISPATCH result;
		InvokeHelper(0x3ec, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void Select()
	{
		InvokeHelper(0xffff, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void SetRange(long Start, long End)
	{
		static BYTE parms[] = VTS_I4 VTS_I4 ;
		InvokeHelper(0x64, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Start, End);
	}
	void Collapse(VARIANT * Direction)
	{
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x65, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Direction);
	}
	void InsertBefore(LPCTSTR Text)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x66, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Text);
	}
	void InsertAfter(LPCTSTR Text)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x68, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Text);
	}
	LPDISPATCH Next(VARIANT * Unit, VARIANT * Count)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x69, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Unit, Count);
		return result;
	}
	LPDISPATCH Previous(VARIANT * Unit, VARIANT * Count)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x6a, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Unit, Count);
		return result;
	}
	long StartOf(VARIANT * Unit, VARIANT * Extend)
	{
		long result;
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x6b, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Unit, Extend);
		return result;
	}
	long EndOf(VARIANT * Unit, VARIANT * Extend)
	{
		long result;
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x6c, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Unit, Extend);
		return result;
	}
	long Move(VARIANT * Unit, VARIANT * Count)
	{
		long result;
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x6d, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Unit, Count);
		return result;
	}
	long MoveStart(VARIANT * Unit, VARIANT * Count)
	{
		long result;
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x6e, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Unit, Count);
		return result;
	}
	long MoveEnd(VARIANT * Unit, VARIANT * Count)
	{
		long result;
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x6f, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Unit, Count);
		return result;
	}
	long MoveWhile(VARIANT * Cset, VARIANT * Count)
	{
		long result;
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x70, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Cset, Count);
		return result;
	}
	long MoveStartWhile(VARIANT * Cset, VARIANT * Count)
	{
		long result;
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x71, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Cset, Count);
		return result;
	}
	long MoveEndWhile(VARIANT * Cset, VARIANT * Count)
	{
		long result;
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x72, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Cset, Count);
		return result;
	}
	long MoveUntil(VARIANT * Cset, VARIANT * Count)
	{
		long result;
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x73, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Cset, Count);
		return result;
	}
	long MoveStartUntil(VARIANT * Cset, VARIANT * Count)
	{
		long result;
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x74, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Cset, Count);
		return result;
	}
	long MoveEndUntil(VARIANT * Cset, VARIANT * Count)
	{
		long result;
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x75, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Cset, Count);
		return result;
	}
	void Cut()
	{
		InvokeHelper(0x77, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Copy()
	{
		InvokeHelper(0x78, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Paste()
	{
		InvokeHelper(0x79, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void InsertBreak(VARIANT * Type)
	{
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x7a, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Type);
	}
	void InsertFile(LPCTSTR FileName, VARIANT * Range, VARIANT * ConfirmConversions, VARIANT * Link, VARIANT * Attachment)
	{
		static BYTE parms[] = VTS_BSTR VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x7b, DISPATCH_METHOD, VT_EMPTY, NULL, parms, FileName, Range, ConfirmConversions, Link, Attachment);
	}
	BOOL InStory(LPDISPATCH Range)
	{
		BOOL result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x7d, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, Range);
		return result;
	}
	BOOL InRange(LPDISPATCH Range)
	{
		BOOL result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x7e, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, Range);
		return result;
	}
	long Delete(VARIANT * Unit, VARIANT * Count)
	{
		long result;
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x7f, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Unit, Count);
		return result;
	}
	long Expand(VARIANT * Unit)
	{
		long result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x81, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Unit);
		return result;
	}
	void InsertParagraph()
	{
		InvokeHelper(0xa0, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void InsertParagraphAfter()
	{
		InvokeHelper(0xa1, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	LPDISPATCH ConvertToTableOld(VARIANT * Separator, VARIANT * NumRows, VARIANT * NumColumns, VARIANT * InitialColumnWidth, VARIANT * Format, VARIANT * ApplyBorders, VARIANT * ApplyShading, VARIANT * ApplyFont, VARIANT * ApplyColor, VARIANT * ApplyHeadingRows, VARIANT * ApplyLastRow, VARIANT * ApplyFirstColumn, VARIANT * ApplyLastColumn, VARIANT * AutoFit)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0xa2, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Separator, NumRows, NumColumns, InitialColumnWidth, Format, ApplyBorders, ApplyShading, ApplyFont, ApplyColor, ApplyHeadingRows, ApplyLastRow, ApplyFirstColumn, ApplyLastColumn, AutoFit);
		return result;
	}
	void InsertDateTimeOld(VARIANT * DateTimeFormat, VARIANT * InsertAsField, VARIANT * InsertAsFullWidth)
	{
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0xa3, DISPATCH_METHOD, VT_EMPTY, NULL, parms, DateTimeFormat, InsertAsField, InsertAsFullWidth);
	}
	void InsertSymbol(long CharacterNumber, VARIANT * Font, VARIANT * Unicode, VARIANT * Bias)
	{
		static BYTE parms[] = VTS_I4 VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0xa4, DISPATCH_METHOD, VT_EMPTY, NULL, parms, CharacterNumber, Font, Unicode, Bias);
	}
	void InsertCrossReference_2002(VARIANT * ReferenceType, long ReferenceKind, VARIANT * ReferenceItem, VARIANT * InsertAsHyperlink, VARIANT * IncludePosition)
	{
		static BYTE parms[] = VTS_PVARIANT VTS_I4 VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0xa5, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ReferenceType, ReferenceKind, ReferenceItem, InsertAsHyperlink, IncludePosition);
	}
	void InsertCaptionXP(VARIANT * Label, VARIANT * Title, VARIANT * TitleAutoText, VARIANT * Position)
	{
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0xa6, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Label, Title, TitleAutoText, Position);
	}
	void CopyAsPicture()
	{
		InvokeHelper(0xa7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void SortOld(VARIANT * ExcludeHeader, VARIANT * FieldNumber, VARIANT * SortFieldType, VARIANT * SortOrder, VARIANT * FieldNumber2, VARIANT * SortFieldType2, VARIANT * SortOrder2, VARIANT * FieldNumber3, VARIANT * SortFieldType3, VARIANT * SortOrder3, VARIANT * SortColumn, VARIANT * Separator, VARIANT * CaseSensitive, VARIANT * LanguageID)
	{
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0xa8, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ExcludeHeader, FieldNumber, SortFieldType, SortOrder, FieldNumber2, SortFieldType2, SortOrder2, FieldNumber3, SortFieldType3, SortOrder3, SortColumn, Separator, CaseSensitive, LanguageID);
	}
	void SortAscending()
	{
		InvokeHelper(0xa9, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void SortDescending()
	{
		InvokeHelper(0xaa, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	BOOL IsEqual(LPDISPATCH Range)
	{
		BOOL result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0xab, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, Range);
		return result;
	}
	float Calculate()
	{
		float result;
		InvokeHelper(0xac, DISPATCH_METHOD, VT_R4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH GoTo(VARIANT * What, VARIANT * Which, VARIANT * Count, VARIANT * Name)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0xad, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, What, Which, Count, Name);
		return result;
	}
	LPDISPATCH GoToNext(long What)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0xae, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, What);
		return result;
	}
	LPDISPATCH GoToPrevious(long What)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0xaf, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, What);
		return result;
	}
	void PasteSpecial(VARIANT * IconIndex, VARIANT * Link, VARIANT * Placement, VARIANT * DisplayAsIcon, VARIANT * DataType, VARIANT * IconFileName, VARIANT * IconLabel)
	{
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0xb0, DISPATCH_METHOD, VT_EMPTY, NULL, parms, IconIndex, Link, Placement, DisplayAsIcon, DataType, IconFileName, IconLabel);
	}
	LPDISPATCH PreviousField()
	{
		LPDISPATCH result;
		InvokeHelper(0xb1, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH NextField()
	{
		LPDISPATCH result;
		InvokeHelper(0xb2, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void InsertParagraphBefore()
	{
		InvokeHelper(0xd4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void InsertCells(VARIANT * ShiftCells)
	{
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0xd6, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ShiftCells);
	}
	void Extend(VARIANT * Character)
	{
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x12c, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Character);
	}
	void Shrink()
	{
		InvokeHelper(0x12d, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	long MoveLeft(VARIANT * Unit, VARIANT * Count, VARIANT * Extend)
	{
		long result;
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x1f4, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Unit, Count, Extend);
		return result;
	}
	long MoveRight(VARIANT * Unit, VARIANT * Count, VARIANT * Extend)
	{
		long result;
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x1f5, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Unit, Count, Extend);
		return result;
	}
	long MoveUp(VARIANT * Unit, VARIANT * Count, VARIANT * Extend)
	{
		long result;
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x1f6, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Unit, Count, Extend);
		return result;
	}
	long MoveDown(VARIANT * Unit, VARIANT * Count, VARIANT * Extend)
	{
		long result;
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x1f7, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Unit, Count, Extend);
		return result;
	}
	long HomeKey(VARIANT * Unit, VARIANT * Extend)
	{
		long result;
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x1f8, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Unit, Extend);
		return result;
	}
	long EndKey(VARIANT * Unit, VARIANT * Extend)
	{
		long result;
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x1f9, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Unit, Extend);
		return result;
	}
	void EscapeKey()
	{
		InvokeHelper(0x1fa, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void TypeText(LPCTSTR Text)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x1fb, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Text);
	}
	void CopyFormat()
	{
		InvokeHelper(0x1fd, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void PasteFormat()
	{
		InvokeHelper(0x1fe, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void TypeParagraph()
	{
		InvokeHelper(0x200, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void TypeBackspace()
	{
		InvokeHelper(0x201, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void NextSubdocument()
	{
		InvokeHelper(0x202, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void PreviousSubdocument()
	{
		InvokeHelper(0x203, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void SelectColumn()
	{
		InvokeHelper(0x204, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void SelectCurrentFont()
	{
		InvokeHelper(0x205, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void SelectCurrentAlignment()
	{
		InvokeHelper(0x206, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void SelectCurrentSpacing()
	{
		InvokeHelper(0x207, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void SelectCurrentIndent()
	{
		InvokeHelper(0x208, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void SelectCurrentTabs()
	{
		InvokeHelper(0x209, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void SelectCurrentColor()
	{
		InvokeHelper(0x20a, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void CreateTextbox()
	{
		InvokeHelper(0x20b, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void WholeStory()
	{
		InvokeHelper(0x20c, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void SelectRow()
	{
		InvokeHelper(0x20d, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void SplitTable()
	{
		InvokeHelper(0x20e, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void InsertRows(VARIANT * NumRows)
	{
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x210, DISPATCH_METHOD, VT_EMPTY, NULL, parms, NumRows);
	}
	void InsertColumns()
	{
		InvokeHelper(0x211, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void InsertFormula(VARIANT * Formula, VARIANT * NumberFormat)
	{
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x212, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Formula, NumberFormat);
	}
	LPDISPATCH NextRevision(VARIANT * Wrap)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x213, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Wrap);
		return result;
	}
	LPDISPATCH PreviousRevision(VARIANT * Wrap)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x214, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Wrap);
		return result;
	}
	void PasteAsNestedTable()
	{
		InvokeHelper(0x215, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	LPDISPATCH CreateAutoTextEntry(LPCTSTR Name, LPCTSTR StyleName)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR VTS_BSTR ;
		InvokeHelper(0x216, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Name, StyleName);
		return result;
	}
	void DetectLanguage()
	{
		InvokeHelper(0x217, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void SelectCell()
	{
		InvokeHelper(0x218, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void InsertRowsBelow(VARIANT * NumRows)
	{
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x219, DISPATCH_METHOD, VT_EMPTY, NULL, parms, NumRows);
	}
	void InsertColumnsRight()
	{
		InvokeHelper(0x21a, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void InsertRowsAbove(VARIANT * NumRows)
	{
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x21b, DISPATCH_METHOD, VT_EMPTY, NULL, parms, NumRows);
	}
	void RtlRun()
	{
		InvokeHelper(0x258, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void LtrRun()
	{
		InvokeHelper(0x259, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void BoldRun()
	{
		InvokeHelper(0x25a, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void ItalicRun()
	{
		InvokeHelper(0x25b, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void RtlPara()
	{
		InvokeHelper(0x25d, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void LtrPara()
	{
		InvokeHelper(0x25e, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void InsertDateTime(VARIANT * DateTimeFormat, VARIANT * InsertAsField, VARIANT * InsertAsFullWidth, VARIANT * DateLanguage, VARIANT * CalendarType)
	{
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x1bc, DISPATCH_METHOD, VT_EMPTY, NULL, parms, DateTimeFormat, InsertAsField, InsertAsFullWidth, DateLanguage, CalendarType);
	}
	void Sort2000(VARIANT * ExcludeHeader, VARIANT * FieldNumber, VARIANT * SortFieldType, VARIANT * SortOrder, VARIANT * FieldNumber2, VARIANT * SortFieldType2, VARIANT * SortOrder2, VARIANT * FieldNumber3, VARIANT * SortFieldType3, VARIANT * SortOrder3, VARIANT * SortColumn, VARIANT * Separator, VARIANT * CaseSensitive, VARIANT * BidiSort, VARIANT * IgnoreThe, VARIANT * IgnoreKashida, VARIANT * IgnoreDiacritics, VARIANT * IgnoreHe, VARIANT * LanguageID)
	{
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x1bd, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ExcludeHeader, FieldNumber, SortFieldType, SortOrder, FieldNumber2, SortFieldType2, SortOrder2, FieldNumber3, SortFieldType3, SortOrder3, SortColumn, Separator, CaseSensitive, BidiSort, IgnoreThe, IgnoreKashida, IgnoreDiacritics, IgnoreHe, LanguageID);
	}
	LPDISPATCH ConvertToTable(VARIANT * Separator, VARIANT * NumRows, VARIANT * NumColumns, VARIANT * InitialColumnWidth, VARIANT * Format, VARIANT * ApplyBorders, VARIANT * ApplyShading, VARIANT * ApplyFont, VARIANT * ApplyColor, VARIANT * ApplyHeadingRows, VARIANT * ApplyLastRow, VARIANT * ApplyFirstColumn, VARIANT * ApplyLastColumn, VARIANT * AutoFit, VARIANT * AutoFitBehavior, VARIANT * DefaultTableBehavior)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x1c9, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Separator, NumRows, NumColumns, InitialColumnWidth, Format, ApplyBorders, ApplyShading, ApplyFont, ApplyColor, ApplyHeadingRows, ApplyLastRow, ApplyFirstColumn, ApplyLastColumn, AutoFit, AutoFitBehavior, DefaultTableBehavior);
		return result;
	}
	long get_NoProofing()
	{
		long result;
		InvokeHelper(0x3ed, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_NoProofing(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x3ed, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_TopLevelTables()
	{
		LPDISPATCH result;
		InvokeHelper(0x3ee, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	BOOL get_LanguageDetected()
	{
		BOOL result;
		InvokeHelper(0x3ef, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_LanguageDetected(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x3ef, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	float get_FitTextWidth()
	{
		float result;
		InvokeHelper(0x3f0, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_FitTextWidth(float newValue)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x3f0, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	void ClearFormatting()
	{
		InvokeHelper(0x3f1, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void PasteAppendTable()
	{
		InvokeHelper(0x3f2, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	LPDISPATCH get_HTMLDivisions()
	{
		LPDISPATCH result;
		InvokeHelper(0x3f3, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_SmartTags()
	{
		LPDISPATCH result;
		InvokeHelper(0x3f7, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_ChildShapeRange()
	{
		LPDISPATCH result;
		InvokeHelper(0x3fd, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	BOOL get_HasChildShapeRange()
	{
		BOOL result;
		InvokeHelper(0x3fe, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_FootnoteOptions()
	{
		LPDISPATCH result;
		InvokeHelper(0x400, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_EndnoteOptions()
	{
		LPDISPATCH result;
		InvokeHelper(0x401, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void ToggleCharacterCode()
	{
		InvokeHelper(0x3f4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void PasteAndFormat(long Type)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x3f5, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Type);
	}
	void PasteExcelTable(BOOL LinkedToExcel, BOOL WordFormatting, BOOL RTF)
	{
		static BYTE parms[] = VTS_BOOL VTS_BOOL VTS_BOOL ;
		InvokeHelper(0x3f6, DISPATCH_METHOD, VT_EMPTY, NULL, parms, LinkedToExcel, WordFormatting, RTF);
	}
	void ShrinkDiscontiguousSelection()
	{
		InvokeHelper(0x3fb, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void InsertStyleSeparator()
	{
		InvokeHelper(0x3fc, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Sort(VARIANT * ExcludeHeader, VARIANT * FieldNumber, VARIANT * SortFieldType, VARIANT * SortOrder, VARIANT * FieldNumber2, VARIANT * SortFieldType2, VARIANT * SortOrder2, VARIANT * FieldNumber3, VARIANT * SortFieldType3, VARIANT * SortOrder3, VARIANT * SortColumn, VARIANT * Separator, VARIANT * CaseSensitive, VARIANT * BidiSort, VARIANT * IgnoreThe, VARIANT * IgnoreKashida, VARIANT * IgnoreDiacritics, VARIANT * IgnoreHe, VARIANT * LanguageID, VARIANT * SubFieldNumber, VARIANT * SubFieldNumber2, VARIANT * SubFieldNumber3)
	{
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x3ff, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ExcludeHeader, FieldNumber, SortFieldType, SortOrder, FieldNumber2, SortFieldType2, SortOrder2, FieldNumber3, SortFieldType3, SortOrder3, SortColumn, Separator, CaseSensitive, BidiSort, IgnoreThe, IgnoreKashida, IgnoreDiacritics, IgnoreHe, LanguageID, SubFieldNumber, SubFieldNumber2, SubFieldNumber3);
	}
	LPDISPATCH get_XMLNodes()
	{
		LPDISPATCH result;
		InvokeHelper(0x136, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_XMLParentNode()
	{
		LPDISPATCH result;
		InvokeHelper(0x137, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Editors()
	{
		LPDISPATCH result;
		InvokeHelper(0x139, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	CString get_XML(BOOL DataOnly)
	{
		CString result;
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x13a, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, parms, DataOnly);
		return result;
	}
	VARIANT get_EnhMetaFileBits()
	{
		VARIANT result;
		InvokeHelper(0x13b, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH GoToEditableRange(VARIANT * EditorID)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x403, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, EditorID);
		return result;
	}
	void InsertXML(LPCTSTR XML, VARIANT * Transform)
	{
		static BYTE parms[] = VTS_BSTR VTS_PVARIANT ;
		InvokeHelper(0x404, DISPATCH_METHOD, VT_EMPTY, NULL, parms, XML, Transform);
	}
	void InsertCaption(VARIANT * Label, VARIANT * Title, VARIANT * TitleAutoText, VARIANT * Position, VARIANT * ExcludeLabel)
	{
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x1a1, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Label, Title, TitleAutoText, Position, ExcludeLabel);
	}
	void InsertCrossReference(VARIANT * ReferenceType, long ReferenceKind, VARIANT * ReferenceItem, VARIANT * InsertAsHyperlink, VARIANT * IncludePosition, VARIANT * SeparateNumbers, VARIANT * SeparatorString)
	{
		static BYTE parms[] = VTS_PVARIANT VTS_I4 VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x1a2, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ReferenceType, ReferenceKind, ReferenceItem, InsertAsHyperlink, IncludePosition, SeparateNumbers, SeparatorString);
	}
	LPDISPATCH get_OMaths()
	{
		LPDISPATCH result;
		InvokeHelper(0x13c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	CString get_WordOpenXML()
	{
		CString result;
		InvokeHelper(0x13d, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void ClearParagraphStyle()
	{
		InvokeHelper(0x406, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void ClearCharacterAllFormatting()
	{
		InvokeHelper(0x407, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void ClearCharacterStyle()
	{
		InvokeHelper(0x408, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void ClearCharacterDirectFormatting()
	{
		InvokeHelper(0x409, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	LPDISPATCH get_ContentControls()
	{
		LPDISPATCH result;
		InvokeHelper(0x40a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_ParentContentControl()
	{
		LPDISPATCH result;
		InvokeHelper(0x40b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void ExportAsFixedFormat(LPCTSTR OutputFileName, long ExportFormat, BOOL OpenAfterExport, long OptimizeFor, BOOL ExportCurrentPage, long Item, BOOL IncludeDocProps, BOOL KeepIRM, long CreateBookmarks, BOOL DocStructureTags, BOOL BitmapMissingFonts, BOOL UseISO19005_1, VARIANT * FixedFormatExtClassPtr)
	{
		static BYTE parms[] = VTS_BSTR VTS_I4 VTS_BOOL VTS_I4 VTS_BOOL VTS_I4 VTS_BOOL VTS_BOOL VTS_I4 VTS_BOOL VTS_BOOL VTS_BOOL VTS_PVARIANT ;
		InvokeHelper(0x40c, DISPATCH_METHOD, VT_EMPTY, NULL, parms, OutputFileName, ExportFormat, OpenAfterExport, OptimizeFor, ExportCurrentPage, Item, IncludeDocProps, KeepIRM, CreateBookmarks, DocStructureTags, BitmapMissingFonts, UseISO19005_1, FixedFormatExtClassPtr);
	}
	void ReadingModeGrowFont()
	{
		InvokeHelper(0x40d, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void ReadingModeShrinkFont()
	{
		InvokeHelper(0x40e, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void ClearParagraphAllFormatting()
	{
		InvokeHelper(0x40f, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void ClearParagraphDirectFormatting()
	{
		InvokeHelper(0x410, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void InsertNewPage()
	{
		InvokeHelper(0x411, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}

	// Selection properties
public:

};
