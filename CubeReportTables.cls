VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CubeReportTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("QuickSort")
Option Explicit

Private Const COL_PD As Long = 1
Private Const COL_PD_VOLUME As Long = 2
Private Const COL_PD_WALLS As Long = 3
Private Const COL_PD_PPW As Long = 4
Private Const COL_PD_UTILIZATION As Long = 5
Private Const COL_AREA As Long = 7
Private Const COL_AREA_VOLUME As Long = 8
Private Const COL_AREA_WALLS As Long = 9
Private Const COL_AREA_PPW As Long = 10
Private Const COL_AREA_UTILIZATION As Long = 11
Private Const COL_BAY As Long = 13
Private Const COL_BAY_VOLUME As Long = 14
Private Const COL_BAY_WALLS As Long = 15
Private Const COL_BAY_PPW As Long = 16
Private Const COL_BAY_UTILIZATION As Long = 17
Private Const COL_DESTINATION As Long = 19
Private Const COL_DESTINATION_VOLUME As Long = 20
Private Const COL_DESTINATION_WALLS As Long = 21
Private Const COL_DESTINATION_PPW As Long = 22
Private Const COL_DESTINATION_UTILIZATION As Long = 23
Private Const COL_TOP5 As Long = 25
Private Const COL_TOP5_VOLUME As Long = 26
Private Const COL_TOP5_WALLS As Long = 27
Private Const COL_TOP5_PPW As Long = 28
Private Const COL_TOP5_UTILIZATION As Long = 29
Private Const COL_MHN As Long = 31
Private Const COL_MHN_VOLUME As Long = 32
Private Const COL_MHN_WALLS As Long = 33
Private Const COL_MHN_PPW As Long = 34
Private Const COL_MHN_UTILIZATION As Long = 35
Private Const COL_EXPENSIVE As Long = 37
Private Const COL_EXPENSIVE_VOLUME As Long = 38
Private Const COL_EXPENSIVE_WALLS As Long = 39
Private Const COL_EXPENSIVE_PPW As Long = 40
Private Const COL_EXPENSIVE_UTILIZATION As Long = 41

Private models As DataModelsFactory
Private wsReportTables As Worksheet

Private Sub Class_Initialize()

    Set models = New DataModelsFactory
    Set wsReportTables = ThisWorkbook.Worksheets("CubeReportTables")

End Sub

Private Sub AddStatsTableRecords(ByVal statTotals As Scripting.Dictionary, _
                                 ByVal colSubject As Long, _
                                 ByVal colVolume As Long, _
                                 ByVal colWalls As Long, _
                                 ByVal colPPW As Long, _
                                 ByVal colUtilization As Long)


    Dim cubeStats As CubeStatsModel
    Dim statTotalsKey As Variant
    Dim rowInsert As Long
    
    rowInsert = 2
    
    For Each statTotalsKey In statTotals.Keys
        Set cubeStats = statTotals(statTotalsKey)
        
        With wsReportTables
            .Cells(rowInsert, colSubject) = cubeStats.Subject
            .Cells(rowInsert, colVolume) = cubeStats.Volume
            .Cells(rowInsert, colWalls) = cubeStats.Walls
            .Cells(rowInsert, colPPW) = cubeStats.PPW
            .Cells(rowInsert, colUtilization) = cubeStats.Utilization
        End With
        
        rowInsert = rowInsert + 1
    Next statTotalsKey

End Sub


Public Sub SetStatsTables(ByVal pdTotals As Scripting.Dictionary, _
                          ByVal areaTotals As Scripting.Dictionary, _
                          ByVal bayTotals As Scripting.Dictionary, _
                          ByVal destinationTotals As Scripting.Dictionary)

    ClearTables
    
    AddStatsTableRecords pdTotals, COL_PD, COL_PD_VOLUME, COL_PD_WALLS, COL_PD_PPW, COL_PD_UTILIZATION
    AddStatsTableRecords areaTotals, COL_AREA, COL_AREA_VOLUME, COL_AREA_WALLS, COL_AREA_PPW, COL_AREA_UTILIZATION
    AddStatsTableRecords bayTotals, COL_BAY, COL_BAY_VOLUME, COL_BAY_WALLS, COL_BAY_PPW, COL_BAY_UTILIZATION
    AddStatsTableRecords destinationTotals, COL_DESTINATION, COL_DESTINATION_VOLUME, COL_DESTINATION_WALLS, COL_DESTINATION_PPW, COL_DESTINATION_UTILIZATION


End Sub


Public Sub ClearTables()

    Dim tbl As ListObject
    Dim colFirst As Long
    Dim colLast As Long
    
    For Each tbl In wsReportTables.ListObjects
        If Not tbl.DataBodyRange Is Nothing Then
            With tbl
                With .DataBodyRange
                    colFirst = .Column
                    colLast = .Column + .Columns.Count - 1
                    .Clear
                End With
                .Resize Range(wsReportTables.Cells(.HeaderRowRange.row, colFirst), _
                              wsReportTables.Cells(.HeaderRowRange.row + 1, colLast))
            End With
        End If
    Next tbl

End Sub
