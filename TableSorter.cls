VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableSorter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("QuickSort.Utilities")
Option Explicit


Public Sub SortTable(ByRef ws As Worksheet, ByRef table As ListObject, ByRef colSort As Long, Optional ByVal newSort As Boolean = True)

    Dim sortValues As Range
    Dim rowFirst As Long
    Dim rowLast As Long
    
    
    If table.DataBodyRange Is Nothing Then
        Exit Sub
    End If
    
    With table.DataBodyRange
        rowFirst = .row
        rowLast = .row + .Rows.Count - 1
    End With
    
    With ws
        Set sortValues = .Range(.Cells(rowFirst, colSort), .Cells(rowLast, colSort))
    End With
    
    With table.Sort
        With .SortFields
            If newSort Then
                .Clear
            End If
            
            .Add Key:=sortValues, _
                 SortOn:=xlSortOnValues, _
                 Order:=xlAscending, _
                 DataOption:=xlSortNormal
        End With
        
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    

End Sub

