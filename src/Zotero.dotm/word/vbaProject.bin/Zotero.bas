Attribute VB_Name = "Zotero"
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

Private Const CP_UTF8 = 65001
Private Const WM_COPYDATA = &H4A

#If VBA7 Then
    Type COPYDATASTRUCT
        dwData As LongPtr
        cbData As Long
        lpData As LongPtr
    End Type
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias _
        "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName _
        As String) As LongPtr
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias _
        "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal _
        wParam As Long, lParam As Any) As Integer
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" _
        (ByVal hwnd As LongPtr) As Boolean
    Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, _
        ByVal dwflags As Long, ByVal lpWideCharStr As LongPtr, _
        ByVal cchWideChar As Long, lpMultiByteStr As Any, _
        ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, _
        ByVal lpUsedDefaultChar As Long) As Long
#Else
    Type COPYDATASTRUCT
        dwData As Long
        cbData As Long
        lpData As Long
    End Type
    Private Declare Function FindWindow Lib "user32" Alias _
        "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName _
        As String) As Long
    Private Declare Function SendMessage Lib "user32" Alias _
        "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal _
        wParam As Long, lParam As Any) As Integer
    Private Declare Function SetForegroundWindow Lib "user32" _
        (ByVal hwnd As Long) As Boolean
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
    Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, _
        ByVal dwflags As Long, ByVal lpWideCharStr As Long, _
        ByVal cchWideChar As Long, lpMultiByteStr As Any, _
        ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, _
        ByVal lpUsedDefaultChar As Long) As Long
#End If

Public Sub ZoteroInsertCitation()
    Call ZoteroCommand("addCitation", True)
End Sub

Public Sub ZoteroInsertBibliography()
    Call ZoteroCommand("addBibliography", False)
End Sub

Public Sub ZoteroEditCitation()
    Call ZoteroCommand("editCitation", True)
End Sub

Public Sub ZoteroEditBibliography()
    Call ZoteroCommand("editBibliography", True)
End Sub

Public Sub ZoteroSetDocPrefs()
    Call ZoteroCommand("setDocPrefs", True)
End Sub

Public Sub ZoteroRefresh()
    Call ZoteroCommand("refresh", False)
End Sub

Public Sub ZoteroRemoveCodes()
    Call ZoteroCommand("removeCodes", False)
End Sub


Sub ZoteroCommand(cmd As String, bringToFront As Boolean)
    Dim cds As COPYDATASTRUCT
    #If VBA7 Then
        Dim ThWnd As LongPtr
    #Else
        Dim ThWnd As Long
    #End If
    Dim a$, args$, name$
    Dim i As Long
    Dim ignore As Long
    Dim sBuffer$
    Dim lLength As Long
    Dim buf() As Byte
    
    ' Try various names for Firefox
    Dim appNames(5)
    appNames(1) = "Zotero"
    appNames(2) = "Firefox"
    appNames(3) = "Browser"
    appNames(4) = "Minefield"
    For i = 1 To 4
        ThWnd = FindWindow(appNames(i) & "MessageWindow", vbNullString)
        If ThWnd <> 0 Then
            Exit For
        End If
    Next
    If ThWnd = 0 Then
        MsgBox ("Word could not communicate with Zotero. Please ensure Firefox is running and try again. If this problem persists, visit http://www.zotero.org/support/word_processor_plugin_troubleshooting")
        Exit Sub
    End If
    
    ' Allow Firefox to bring a window to the front
    If bringToFront Then Call SetForegroundWindow(ThWnd)
    
    ' Get path to active document
    If ActiveDocument.Path <> "" Then
        name$ = ActiveDocument.Path & Application.PathSeparator & ActiveDocument.name
    Else
        name$ = ActiveDocument.name
    End If
    
    ' Set up command line arguments
    name$ = Replace(name$, """", """""")
    args$ = "-silent -ZoteroIntegrationAgent WinWord -ZoteroIntegrationCommand " & cmd & " -ZoteroIntegrationDocument """ & name$ & """"
    a$ = "firefox.exe " & args$ & Chr$(0) & "C:\"
    
    ' Do some UTF-8 magic
    lLength = WideCharToMultiByte(CP_UTF8, 0, StrPtr(a$), -1, ByVal 0, 0, 0, 0)
    ReDim buf(lLength) As Byte
    Call WideCharToMultiByte(CP_UTF8, 0, StrPtr(a$), -1, buf(1), lLength, 0, 0)
    
    ' Send message to Firefox
    cds.dwData = 1
    cds.cbData = lLength
    cds.lpData = VarPtr(buf(1))
    i = SendMessage(ThWnd, WM_COPYDATA, 0, cds)
    
    ' Handle error
    If Err.LastDllError = 5 Then
        If Dir("C:\Program Files\Mozilla Firefox\firefox.exe") <> "" Then
            Call Shell("""C:\Program Files\Mozilla Firefox\firefox.exe"" " & args$, vbNormalFocus)
        ElseIf Dir("C:\Program Files (x86)\Mozilla Firefox\firefox.exe") <> "" Then
            Call Shell("""C:\Program Files (x86)\Mozilla Firefox\firefox.exe"" " & args$, vbNormalFocus)
        End If
    End If
End Sub

