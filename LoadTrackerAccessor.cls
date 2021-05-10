VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoadTrackerAccessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("QuickSort.DataAccess")
Option Explicit

Private Const COL_LOAD_ID As Long = 2
Private Const COL_LOAD_DATE As Long = 3
Private Const COL_SORT As Long = 4
Private Const COL_AREA As Long = 5
Private Const COL_BAY As Long = 6
Private Const COL_DESTINATION As Long = 7
Private Const COL_ULD As Long = 8
Private Const COL_EQUIPMENT_TYPE As Long = 9
Private Const COL_START_PCT As Long = 10
Private Const COL_END_PCT As Long = 11
Private Const COL_VOLUME As Long = 12
Private Const COL_STATUS As Long = 13

Private modelsFactory As DataModelsFactory
Private wsDashboard As Worksheet
Private tblDashboard As ListObject

Public Sub CreateLoad(ByRef load As LoadModel)

    Dim rowTarget As Long
    
    If Not tblDashboard.DataBodyRange Is Nothing Then
        rowTarget = tblDashboard.DataBodyRange.row + tblDashboard.DataBodyRange.Rows.Count
    Else
        rowTarget = tblDashboard.HeaderRowRange.row + 1
    End If
    
    With wsDashboard
        .Cells(rowTarget, COL_LOAD_ID) = load.LoadID
        .Cells(rowTarget, COL_LOAD_DATE) = load.loadDate
        .Cells(rowTarget, COL_SORT) = load.Sort
        .Cells(rowTarget, COL_AREA) = load.Area
        .Cells(rowTarget, COL_BAY) = load.Bay
        .Cells(rowTarget, COL_DESTINATION) = load.Destination
        .Cells(rowTarget, COL_ULD) = load.trailerNumber
        .Cells(rowTarget, COL_EQUIPMENT_TYPE) = load.equipmentType
        .Cells(rowTarget, COL_START_PCT) = load.StartPercentage
        .Cells(rowTarget, COL_END_PCT) = load.EndPercentage
        .Cells(rowTarget, COL_VOLUME) = load.Volume
        .Cells(rowTarget, COL_STATUS) = "Active"
    End With
    
End Sub

Public Function ReadAllLoads() As Scripting.Dictionary
    
    Dim loadTrackers As Scripting.Dictionary
    Dim loadULD As String
    Dim rowFirst As Long
    Dim rowLast As Long
    Dim rowCurrent As Long
    
    Set loadTrackers = New Scripting.Dictionary
    
    With tblDashboard.DataBodyRange
        rowFirst = .row
        rowLast = .row + .Rows.Count - 1
    End With
    
    
    For rowCurrent = rowFirst To rowLast
        loadULD = wsDashboard.Cells(rowCurrent, COL_ULD).Value2
        loadTrackers.Add loadULD, ReadLoad(loadULD)
    Next rowCurrent
    
    Set ReadAllLoads = loadTrackers

End Function


Public Function ReadLoad(ByVal uld As String) As LoadModel

    
    Dim rowTarget As Long
    Dim load As LoadModel
    
    rowTarget = GetRowTarget(uld)

    If rowTarget <> 0 Then
        Set load = modelsFactory.BuildLoadModel
        
        With wsDashboard
            load.LoadID = .Cells(rowTarget, COL_LOAD_ID).Value2
            load.loadDate = .Cells(rowTarget, COL_LOAD_DATE).Value2
            load.Sort = .Cells(rowTarget, COL_SORT).Value2
            load.Area = .Cells(rowTarget, COL_AREA).Value2
            load.Bay = .Cells(rowTarget, COL_BAY).Value2
            load.Destination = .Cells(rowTarget, COL_DESTINATION).Value2
            load.trailerNumber = .Cells(rowTarget, COL_ULD).Value2
            load.equipmentType = .Cells(rowTarget, COL_EQUIPMENT_TYPE).Value2
            load.StartPercentage = .Cells(rowTarget, COL_START_PCT).Value2
            load.EndPercentage = .Cells(rowTarget, COL_END_PCT).Value2
            load.Volume = .Cells(rowTarget, COL_VOLUME).Value2
            load.Status = .Cells(rowTarget, COL_STATUS).Value2
        End With
    
    Else
        MsgBox "LoadTrackerAccessor.ReadLoad.uld not found " & uld, vbExclamation, "Application-Defined Error"
    End If

    Set ReadLoad = load

End Function






Public Sub DeleteAllLoads()
    
    Dim rowFirst As Long
    Dim rowLast As Long
    
    With tblDashboard
        If Not .DataBodyRange Is Nothing Then
            With .DataBodyRange
                rowFirst = .row
                rowLast = .row + .Rows.Count - 1
            End With
        Else
            With .HeaderRowRange
                rowFirst = .row
                rowLast = .row
            End With
        End If
    End With
    
    
    With wsDashboard
        .Range(.Cells(rowFirst, COL_LOAD_ID), .Cells(rowLast, COL_STATUS)).Clear
        tblDashboard.Resize .Range(.Cells(tblDashboard.HeaderRowRange.row, COL_LOAD_ID), .Cells(rowFirst, COL_STATUS))
        .Range(.Cells(rowFirst, COL_LOAD_ID), .Cells(rowLast, COL_STATUS)).Interior.Color = RGB(66, 66, 66)
        
    End With
    
End Sub


Private Function GetRowTarget(ByVal uld As String) As Long

    Dim found As Boolean
    Dim rowTarget As Long
    Dim rowFirst As Long
    Dim rowLast As Long
    Dim rowCurrent As Long
    
    found = False
    rowTarget = 0
    
    If Not tblDashboard Is Nothing Then
        With tblDashboard.DataBodyRange
            rowFirst = .row
            rowLast = .row + .Rows.Count - 1
        End With
        
        rowCurrent = rowFirst
        
        Do While Not found And rowCurrent <= rowLast
            If wsDashboard.Cells(rowCurrent, COL_ULD).Value2 = uld Then
                found = True
                rowTarget = rowCurrent
            End If
            
            rowCurrent = rowCurrent + 1
        Loop
    End If
    
    GetRowTarget = rowTarget

End Function


Private Sub Class_Initialize()

    Set modelsFactory = New DataModelsFactory
    Set wsDashboard = ThisWorkbook.Worksheets("Dashboard")
    Set tblDashboard = wsDashboard.ListObjects(1)

End Sub
