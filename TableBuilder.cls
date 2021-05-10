VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("QuickSort.Utilities")
Option Explicit

Private tbl As ListObject
Private rowLast As Long
Private colLast As Long

Public Function CreateTable(ByRef ws As Worksheet) As ListObject

    SetLastActiveRow ws
    SetLastActiveColumn ws
    
    With ws
        Set tbl = .ListObjects.Add(xlSrcRange, _
                                   .Range(.Cells(1, 1), .Cells(GetLastActiveRow, GetLastActiveColumn)), _
                                   xlListObjectHasHeaders:=xlYes)
    End With
    
    Set CreateTable = tbl

End Function


Private Sub SetLastActiveRow(ByRef ws As Worksheet)

    rowLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

End Sub


Public Function GetLastActiveRow() As Long

    GetLastActiveRow = rowLast

End Function


Private Sub SetLastActiveColumn(ByRef ws As Worksheet)

    colLast = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

End Sub


Public Function GetLastActiveColumn() As Long

    GetLastActiveColumn = colLast

End Function


Private Sub Class_Initialize()

    rowLast = 0
    colLast = 0

End Sub
