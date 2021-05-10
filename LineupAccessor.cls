VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LineupAccessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("QuickSort.DataAccess")
Option Explicit

Private Const COL_AREA As Long = 1
Private Const COL_BAY As Long = 2
Private Const COL_DAY_DESTINATION As Long = 3
Private Const COL_TWILIGHT_DESTINATION As Long = 4
Private Const COL_NIGHT_DESTINATION As Long = 5

Public Function GetLineup(ByVal lineupName As String, ByVal sortName As String) As Scripting.Dictionary
    
    
    
    Dim lineupDict As Scripting.Dictionary
    Dim modelsFactory As DataModelsFactory
    Dim lineupTableContent As Variant
    Dim lineupBay As LineupModel
    
    Dim rowFirst As Long
    Dim rowLast As Long
    Dim rowCurrent As Long
    Dim colDestination As Long
    
    Set modelsFactory = New DataModelsFactory
    Set lineupDict = New Scripting.Dictionary
    
    lineupTableContent = GetLineupTableContent(lineupName)
    rowFirst = LBound(lineupTableContent, 1)
    rowLast = UBound(lineupTableContent, 1)
    colDestination = GetDestinationColumn(sortName)
    
    For rowCurrent = rowFirst To rowLast
        Set lineupBay = modelsFactory.BuildLineupModel
        
        With lineupBay
            .workArea = lineupTableContent(rowCurrent, COL_AREA)
            .bayNumber = lineupTableContent(rowCurrent, COL_BAY)
            .Destination = lineupTableContent(rowCurrent, colDestination)
        End With
        
        lineupDict.Add lineupBay.bayNumber, lineupBay
    Next
    
    Set GetLineup = lineupDict
    
End Function


Public Function GetLineupTableContent(ByVal worksheetName As String) As Variant

    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set ws = ThisWorkbook.Worksheets(worksheetName)
    Set tbl = ws.ListObjects(1)
    
    GetLineupTableContent = tbl.DataBodyRange.Value

End Function


Private Function GetDestinationColumn(ByVal sortName As String) As Long

    Dim colDestination As Long
    
    If sortName = "LEXKY-D" Then
        colDestination = COL_DAY_DESTINATION
    ElseIf sortName = "LEXKY-T" Then
        colDestination = COL_TWILIGHT_DESTINATION
    ElseIf sortName = "LEXKY-N" Then
        colDestination = COL_NIGHT_DESTINATION
    End If
    
    GetDestinationColumn = colDestination

End Function
