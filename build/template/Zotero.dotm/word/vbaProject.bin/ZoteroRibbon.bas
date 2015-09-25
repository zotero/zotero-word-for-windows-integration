Attribute VB_Name = "ZoteroRibbon"
' ***** BEGIN LICENSE BLOCK *****
'
' Copyright (c) 2015  Zotero
'                     Center for History and New Media
'                     George Mason University, Fairfax, Virginia, USA
'                     http://zotero.org
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
' ***** END LICENSE BLOCK *****

Option Explicit

Global selectionIsCitation As Boolean
Global selectionIsBibliography As Boolean
Global documentHasFields As Boolean

Global ribbonUI As IRibbonUI
Global theAppEventHandler As New AppEventHandler
Global initialized As Boolean

Sub getZoteroFieldAtSelection(ByRef isCitation As Boolean, ByRef isBibliography As Boolean)
    ' Look for fields
    Dim rgFields As Fields, selectionRg As Range, selectionStart As Long, selectionEnd As Long
    Set rgFields = Selection.Fields
    Set selectionRg = Selection.Range
    selectionStart = selectionRg.Start
    selectionEnd = selectionRg.End

    If rgFields.Count = 0 Then
        ' If there are no fields in the current selection, we will need to create a
        ' range that would span any field
        Dim rg As Range, rgEnd As Range, _
            rgStartIndex As Long, rgEndIndex As Long
        
        ' There is a bug in Word that can cause this to clear the selection highlighting
        Application.ScreenUpdating = False
        On Error GoTo Finally
        Set rg = selectionRg.Duplicate.GoToPrevious(wdGoToField)
        Set rgEnd = selectionRg.Duplicate.GoToNext(wdGoToField)
Finally:
        Application.ScreenUpdating = True
        
        ' Might be only one field in the entire doc
        rgStartIndex = rg.Start
        rgEndIndex = rgEnd.End
        If rgStartIndex = selectionStart Or rgEndIndex <= selectionEnd Then
            Set rg = ActiveDocument.StoryRanges(selectionRg.storyType).Duplicate
            If rgStartIndex <> selectionStart Then rg.Start = rgStartIndex
            If rgEndIndex <= selectionEnd Then rgEndIndex = rg.End
        End If
        
        ' Make range span from previous field to next field
        rg.End = rgEndIndex
        
        Set rgFields = rg.Fields
    End If
    
    Call hasZoteroField(rgFields, selectionStart, selectionEnd, isCitation, isBibliography)
    If isCitation Or isBibliography Then Exit Sub
    Call hasZoteroBookmark(Selection.Bookmarks, isCitation, isBibliography)
End Sub

Function documentHasZoteroFields() As Boolean
    Dim isCitation As Boolean, isBibliography As Boolean
    Dim i As Integer
    For i = 1 To 3 ' Main text, footnotes, and endnotes
        On Error Resume Next
        Call hasZoteroField(ActiveDocument.StoryRanges(i).Fields, -1, -1, _
                            isCitation, isBibliography)
        On Error GoTo 0
        If isCitation Or isBibliography Then
            documentHasZoteroFields = True
            Exit Function
        End If
    Next i
    For i = 1 To 3 ' Main text, footnotes, and endnotes
        On Error Resume Next
        Call hasZoteroBookmark(ActiveDocument.StoryRanges(i).Bookmarks, _
                               isCitation, isBibliography)
        On Error GoTo 0
        If isCitation Or isBibliography Then
            documentHasZoteroFields = True
            Exit Function
        End If
    Next i
    documentHasZoteroFields = False
End Function

' Find the type of the first Zotero field within rgFields that is within selectionStart
' and selectionEnd
Sub hasZoteroField(rgFields As Fields, selectionStart As Long, selectionEnd As Long, _
                   ByRef isCitation As Boolean, ByRef isBibliography As Boolean)
    Dim testField As field
    For Each testField In rgFields
        Dim testFieldCode As String, testFieldStart As Long, testFieldEnd As Long

        On Error GoTo ErrHandler
        testFieldCode = testField.code.Text
        testFieldStart = testField.code.Start
        testFieldEnd = testField.Result.End
        On Error GoTo 0
        
        If selectionStart <> -1 And (testFieldStart > selectionEnd _
                                     Or testFieldEnd < selectionStart) Then
            GoTo LoopEnd
        End If
        
        isCitation = InStr(testFieldCode, " CSL_CITATION") > 0 _
                     Or InStr(testFieldCode, " ADDIN ZOTERO_CITATION") > 0
        isBibliography = InStr(testFieldCode, " CSL_BIBLIOGRAPHY") > 0 _
                         Or InStr(testFieldCode, " ADDIN ZOTERO_BIBLIOGRAPHY") > 0
        If isCitation Or isBibliography Then Exit Sub
LoopEnd:
    Next testField
    Exit Sub
ErrHandler:
    Resume LoopEnd
End
End Sub

' Find the type of the first Zotero bookmark within rgBookmarks
Sub hasZoteroBookmark(rgBookmarks As Bookmarks, ByRef isCitation As Boolean, _
                      ByRef isBibliography As Boolean)
    Dim testBookmark As Bookmark
    For Each testBookmark In rgBookmarks
        Dim bookmarkName As String
        bookmarkName = testBookmark
        If InStr(bookmarkName, "ZOTERO_") > 0 Or InStr(bookmarkName, "CSL_") > 0 Then
            ' Look for corresponding property
            Dim bookmarkContent As String
            On Error GoTo ErrHandler
            bookmarkContent = ActiveDocument.CustomDocumentProperties(bookmarkName & "_1").Value
            On Error GoTo 0
        
            isCitation = InStr(bookmarkContent, "CSL_CITATION") > 0 _
                         Or InStr(bookmarkContent, "ZOTERO_CITATION") > 0
            isBibliography = InStr(bookmarkContent, "CSL_BIBLIOGRAPHY") > 0 _
                             Or InStr(bookmarkContent, "ZOTERO_BIBLIOGRAPHY") > 0
            If isCitation Or isBibliography Then Exit For
        End If
LoopEnd:
    Next testBookmark
    Exit Sub
ErrHandler:
    Resume LoopEnd
End Sub

Sub ZoteroInitRibbon(ribbon As IRibbonUI)
    Set ribbonUI = ribbon
    ribbonUI.Invalidate

    ThisDocument.Saved = True
    Set theAppEventHandler.App = Word.Application
    initialized = True
End Sub

Sub ZoteroRibbonInsertCitation(button As IRibbonControl)
    Call ZoteroInsertCitation
End Sub

Sub ZoteroRibbonInsertBibliography(button As IRibbonControl)
    Call ZoteroInsertBibliography
End Sub

Sub ZoteroRibbonEditCitation(button As IRibbonControl)
    Call ZoteroEditCitation
End Sub

Sub ZoteroRibbonEditBibliography(button As IRibbonControl)
    Call ZoteroEditBibliography
End Sub

Sub ZoteroRibbonSetDocPrefs(button As IRibbonControl)
    Call ZoteroSetDocPrefs
End Sub

Sub ZoteroRibbonRefresh(button As IRibbonControl)
    Call ZoteroRefresh
End Sub

Sub ZoteroRibbonRemoveCodes(button As IRibbonControl)
    Call ZoteroRemoveCodes
End Sub

Sub ZoteroTabLabel(tb As IRibbonControl, ByRef returnedVal)
    Dim ver As Double
    ver = Val(Application.Version)
    If ver >= 15 And ver < 16 Then
        returnedVal = "ZOTERO"
    Else
        returnedVal = "Zotero"
    End If
End Sub

Sub ZoteroInsertCitationVisible(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Not selectionIsCitation
End Sub

Sub ZoteroEditCitationVisible(control As IRibbonControl, ByRef returnedVal)
    returnedVal = selectionIsCitation
End Sub

Sub ZoteroCitationEnabled(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Not selectionIsBibliography
End Sub

Sub ZoteroInsertBibliographyVisible(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Not selectionIsBibliography
End Sub

Sub ZoteroEditBibliographyVisible(control As IRibbonControl, ByRef returnedVal)
    returnedVal = selectionIsBibliography
End Sub

Sub ZoteroBibliographyEnabled(control As IRibbonControl, ByRef returnedVal)
    returnedVal = documentHasFields And Not selectionIsCitation
End Sub

Sub ZoteroRefreshEnabled(control As IRibbonControl, ByRef returnedVal)
    returnedVal = documentHasFields
End Sub
