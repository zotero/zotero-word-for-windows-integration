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
    Global ZotWnd As LongPtr
#Else
    Global ZotWnd As Long
#End If
Global IsZotero7 As Boolean

#If VBA7 Then
    Type COPYDATASTRUCT
        dwData As LongPtr
        cbData As Long
        lpData As LongPtr
    End Type
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias _
        "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName _
        As String) As LongPtr
    Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias _
        "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, _
        ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias _
        "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal _
        wParam As Long, lParam As Any) As Integer
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" _
        (ByVal hwnd As LongPtr) As Boolean
    Private Declare PtrSafe Function EnumThreadWindows Lib "user32" _
        (ByVal dwThreadId As Long, ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Boolean
    Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" _
        (ByVal hwnd As LongPtr, lpdwProcessId As Long) As Long
    Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" _
        (ByVal hwnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
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
    Private Declare Function FindWindowEx Lib "user32" Alias _
        "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, _
        ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
    Private Declare Function SendMessage Lib "user32" Alias _
        "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal _
        wParam As Long, lParam As Any) As Integer
    Private Declare Function SetForegroundWindow Lib "user32" _
        (ByVal hwnd As Long) As Boolean
    Private Declare Function EnumThreadWindows Lib "user32" _
        (ByVal dwThreadId As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean
    Private Declare Function GetWindowThreadProcessId Lib "user32" _
        (ByVal hwnd As Long, lpdwProcessId As Long) As Long
    Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" _
        (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
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

Public Sub ZoteroAddEditCitation()
    Call ZoteroCommand("addEditCitation", True)
End Sub

Public Sub ZoteroAddNote()
    Call ZoteroCommand("addNote", True)
End Sub

Public Sub ZoteroAddAnnotation()
    Call ZoteroCommand("addAnnotation", True)
End Sub

Public Sub ZoteroAddEditBibliography()
    Call ZoteroCommand("addEditBibliography", True)
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

Private Sub FindZoteroWindow()
    ZotWnd = 0
    #If VBA7 Then
        Dim ThWnd As LongPtr
    #Else
        Dim ThWnd As Long
    #End If
    ' Zotero 6 / FX60+
    ThWnd = FindWindow("ZoteroMessageWindow", vbNullString)
    If ThWnd <> 0 Then
        ZotWnd = ThWnd
        Exit Sub
    End If
    
    IsZotero7 = True
    
    ' Zotero 7 / FX102+
    Dim lpdwThreadId As Long
    ThWnd = FindWindow("MozillaWindowClass", vbNullString)
    Do While ThWnd <> 0
        lpdwThreadId = GetWindowThreadProcessId(ThWnd, 0)
        Call EnumThreadWindows(lpdwThreadId, AddressOf EnumWindowsCallback, ByVal 0&)
        If ZotWnd <> 0 Then
            Exit Do
        End If
        ThWnd = FindWindowEx(0, ThWnd, "MozillaWindowClass", vbNullString)
    Loop
End Sub

Function EnumWindowsCallback(ByVal hwnd As Long, ByVal lParams As Long) As Long ' {
    Dim windowClass As String * 256
    Dim retVal      As Long
    Dim zoteroPosition As Long
    Dim remoteWindowPosition As Long

    retVal = GetClassName(hwnd, windowClass, 255)
    windowClass = Left$(windowClass, retVal)
    zoteroPosition = InStr(windowClass, "Mozilla_zotero_")
    remoteWindowPosition = InStr(windowClass, "RemoteWindow")
    ' Looking for window name like `Mozilla_zotero_%profileName%_RemoteWindow`
    ' which is not configurable and used to be much simpler in Z6 - `ZoteroMessageWindow`
    If zoteroPosition <> 0 And remoteWindowPosition <> 0 Then
        ZotWnd = hwnd
        EnumWindowsCallback = False
    Else
        '
        ' Return true to indicate that we want to continue
        ' with the enumeration of the windows:
        '
        EnumWindowsCallback = True
    End If
End Function ' }


Sub ZoteroCommand(cmd As String, bringToFront As Boolean)
    Dim cds As COPYDATASTRUCT
    Dim a$, args$, name$, templateVersion$
    Dim i As Long
    Dim ignore As Long
    Dim sBuffer$
    Dim lLength As Long
    Dim buf() As Byte
    
    Call FindZoteroWindow
    If ZotWnd = 0 Then
        MsgBox ("Word could not communicate with Zotero. Please ensure Zotero is running and try again. If this problem persists, see https://www.zotero.org/support/word_processor_plugin_troubleshooting")
        Exit Sub
    End If
    
    ' Allow Firefox to bring a window to the front
    If bringToFront Then Call SetForegroundWindow(ZotWnd)
    
    ' Get path to active document
    If ActiveDocument.Path <> "" Then
        name$ = ActiveDocument.Path & Application.PathSeparator & ActiveDocument.name
    Else
        name$ = ActiveDocument.name
    End If
    
    templateVersion$ = 1
    
    ' Set up command line arguments
    name$ = Replace(name$, """", """""")
    args$ = "-silent -ZoteroIntegrationAgent WinWord -ZoteroIntegrationCommand " & cmd & " -ZoteroIntegrationDocument """ & name$ & """ -ZoteroIntegrationTemplateVersion " & templateVersion$
    a$ = "zotero.exe " & args$ & Chr$(0) & "C:\"
    
    If IsZotero7 Then
        ' With FX128+ WM_COPYDATA either has to be in UTF16 (native VBA string encoding)
        ' or the UTF8 WM_COPYDATA must start with Fire + Fox emojis. We cannot do that
        ' because the VBA editor doesn't support unicode at all.
        ' Either way this works with FX115 too which is now released
        ' so we do this for Zotero 7
        cds.dwData = 2
        cds.cbData = LenB(a$)
        cds.lpData = StrPtr(a$)
    Else
        ' Do some UTF-8 magic for Zotero 6
        lLength = WideCharToMultiByte(CP_UTF8, 0, StrPtr(a$), -1, ByVal 0, 0, 0, 0)
        ReDim buf(lLength) As Byte
        Call WideCharToMultiByte(CP_UTF8, 0, StrPtr(a$), -1, buf(1), lLength, 0, 0)
    
        cds.dwData = 1
        cds.cbData = lLength
        cds.lpData = VarPtr(buf(1))
    End If
    
    ' Send message to Firefox
    i = SendMessage(ZotWnd, WM_COPYDATA, 0, cds)
    
    ' Handle error
    If Err.LastDllError = 5 Then
        If Dir("C:\Program Files\Zotero\zotero.exe.exe") <> "" Then
            Call Shell("""C:\Program Files\Zotero\zotero.exe"" " & args$, vbNormalFocus)
        ElseIf Dir("C:\Program Files (x86)\Zotero\zotero.exe.exe") <> "" Then
            Call Shell("""C:\Program Files (x86)\Zotero\zotero.exe.exe"" " & args$, vbNormalFocus)
        End If
    End If
End Sub

