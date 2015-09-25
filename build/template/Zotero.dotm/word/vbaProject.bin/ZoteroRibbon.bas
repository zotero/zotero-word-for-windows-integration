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

Sub ZoteroRibbonAddEditCitation(button As IRibbonControl)
    Call ZoteroAddEditCitation
End Sub

Sub ZoteroRibbonInsertBibliography(button As IRibbonControl)
    Call ZoteroInsertBibliography
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
