Attribute VB_Name = "AppEventHandler"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
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
' Inspired by code copyright (c) 2009-2012 Mendeley Ltd.
'
' Licensed under the Educational Community License, Version 1.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'
' http://www.opensource.org/licenses/ecl1.php
'
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.
'
' ***** END LICENSE BLOCK *****

Option Explicit

Public WithEvents App As Word.Application
Attribute App.VB_VarHelpID = -1

Private Sub App_WindowSelectionChange(ByVal Sel As Selection)
    Call UpdateRibbon
End Sub

Private Sub App_WindowActivate(ByVal Doc As Document, ByVal Wn As Window)
    Call UpdateRibbon
End Sub

Private Sub UpdateRibbon()
    If Not initialized Then
        Exit Sub
    End If
    
    On Error GoTo ErrHandler

    Dim isCitation As Boolean
    Dim isBibliography As Boolean
    Call getZoteroFieldAtSelection(isCitation, isBibliography)
    Dim hasFields As Boolean
    hasFields = documentHasZoteroFields()

    If isCitation <> selectionIsCitation Or _
       isBibliography <> selectionIsBibliography Or _
       hasFields <> documentHasFields Then
        selectionIsCitation = isCitation
        selectionIsBibliography = isBibliography
        documentHasFields = hasFields
        ribbonUI.Invalidate
    End If

ErrHandler:
    ThisDocument.Saved = True
End Sub


