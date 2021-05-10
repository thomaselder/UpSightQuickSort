VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppVarsAccessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule
'@Folder("QuickSort.DataAccess")
Option Explicit

Private Const COL_APP_VAR_NAME As Long = 1
Private Const COL_APP_VAR_VALUE As Long = 2

Private wsAppVars As Worksheet
Private tblAppVars As ListObject
Private tblContentAppVars As Variant
Private appVarRows As Scripting.Dictionary

Private Sub Class_Initialize()
    
    Set wsAppVars = ThisWorkbook.Worksheets("AppVars")
    Set tblAppVars = wsAppVars.ListObjects(1)
    tblContentAppVars = tblAppVars.DataBodyRange
    SetAppVarRows

End Sub

Private Sub SetAppVarRows()

    Dim rowFirst As Long
    Dim rowLast As Long
    Dim rowCurrent As Long
    Dim colAppVarName As Long
    Dim colAppVarValue As Long
    
    Set appVarRows = New Scripting.Dictionary
    
    rowFirst = LBound(tblContentAppVars, 1)
    rowLast = UBound(tblContentAppVars, 1)
    
    colAppVarName = 1
    colAppVarValue = 2

    For rowCurrent = rowFirst To rowLast
        appVarRows.Add tblContentAppVars(rowCurrent, colAppVarName), rowCurrent + tblAppVars.HeaderRowRange.row
    Next rowCurrent

End Sub


Public Function GetAppVar(ByVal appVarName As String) As Variant

    Dim appVarValue As Variant

    If appVarName = "loadID" Then
        appVarValue = CLng(wsAppVars.Cells(appVarRows(appVarName), COL_APP_VAR_VALUE).Value2)
        wsAppVars.Cells(appVarRows(appVarName), COL_APP_VAR_VALUE) = appVarValue + 1
    End If

    GetAppVar = appVarValue

End Function
