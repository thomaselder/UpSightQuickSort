VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoadUpdateForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("QuickSort")
Option Explicit


Private Const COL_BAY_EQUIPMENT_LINEUP As Long = 1
Private Const COL_ULD_EQUIPMENT_LINEUP As Long = 4
Private Const COL_EQUIPMENT_TYPE_EQUIPMENT_LINEUP As Long = 5
Private Const COL_DESTINATION_EQUIPMENT_LINEUP As Long = 8
Private Const COL_DESTINATION_SORT_EQUIPMENT_LINEUP As Long = 9
Private Const COL_AREA_LOAD_UPDATE As Long = 1
Private Const COL_BAY_LOAD_UPDATE As Long = 2
Private Const COL_ULD_LOAD_UPDATE As Long = 3
Private Const COL_EQUIPMENT_TYPE_LOAD_UPDATE As Long = 4
Private Const COL_DESTINATION_LOAD_UPDATE As Long = 5
Private Const COL_MEASURE_LOAD_UPDATE As Long = 6

Public Sub Build(ByVal equipmentLineupFilePath As String)
    
    Dim fso As FileSystemObject
    Dim accessFactory As DataAccessFactory
    Dim accessor As LineupAccessor
    Dim tblBuilder As TableBuilder
    Dim utilities As UtilitiesFactory
    Dim lineup As Scripting.Dictionary
    Dim lineupBay As LineupModel
    Dim wbEquipmentLineup As Workbook
    Dim wsEquipmentLineup As Worksheet
    Dim wbLoadUpdate As Workbook
    Dim wsLoadUpdate As Worksheet
    Dim tblEquipmentLineup As ListObject
    Dim tblContentEquipmentLineup As Variant
    Dim bayDestination As String
    Dim bayDestinationSort As String
    
    Dim rowFirstEquipmentLineup As Long
    Dim rowLastEquipmentLineup As Long
    Dim rowCurrentEquipmentLineup As Long
    
    Dim rowCurrentLoadUpdate As Long
       
    
    Set fso = New FileSystemObject
    Set accessFactory = New DataAccessFactory
    Set utilities = New UtilitiesFactory
    Set accessor = accessFactory.BuildLineupAccessor
    
    Application.ScreenUpdating = False
    
    If fso.FileExists(equipmentLineupFilePath) Then
        Set lineup = accessor.GetLineup("Lineup", "LEXKY-T")
        Set wbEquipmentLineup = Workbooks.Open(equipmentLineupFilePath)
        Set wsEquipmentLineup = wbEquipmentLineup.Worksheets(1)
        Set tblBuilder = utilities.BuildTableBuilder
        
        tblBuilder.CreateTable wsEquipmentLineup
        
        Set tblEquipmentLineup = wsEquipmentLineup.ListObjects(1)
        
        tblContentEquipmentLineup = tblEquipmentLineup.DataBodyRange
        
        Set wbLoadUpdate = Workbooks.Add
        Set wsLoadUpdate = wbLoadUpdate.Worksheets(1)
        
        SetHeadersLoadUpdate wsLoadUpdate
        
        rowFirstEquipmentLineup = LBound(tblContentEquipmentLineup, 1)
        rowLastEquipmentLineup = UBound(tblContentEquipmentLineup, 1)
        rowCurrentLoadUpdate = 2
        
        For rowCurrentEquipmentLineup = rowFirstEquipmentLineup To rowLastEquipmentLineup
            If lineup.Exists(CStr(tblContentEquipmentLineup(rowCurrentEquipmentLineup, COL_BAY_EQUIPMENT_LINEUP))) Then
                Set lineupBay = lineup(CStr(tblContentEquipmentLineup(rowCurrentEquipmentLineup, COL_BAY_EQUIPMENT_LINEUP)))
                
                With wsLoadUpdate
                    .Cells(rowCurrentLoadUpdate, COL_AREA_LOAD_UPDATE) = lineupBay.workArea
                    .Cells(rowCurrentLoadUpdate, COL_BAY_LOAD_UPDATE) = tblContentEquipmentLineup(rowCurrentEquipmentLineup, COL_BAY_EQUIPMENT_LINEUP)
                    
                    If IsUPSTrailer(tblContentEquipmentLineup(rowCurrentEquipmentLineup, COL_ULD_EQUIPMENT_LINEUP)) Then
                        .Cells(rowCurrentLoadUpdate, COL_ULD_LOAD_UPDATE) = "UPST" & tblContentEquipmentLineup(rowCurrentEquipmentLineup, COL_ULD_EQUIPMENT_LINEUP)
                    Else
                        .Cells(rowCurrentLoadUpdate, COL_ULD_LOAD_UPDATE) = tblContentEquipmentLineup(rowCurrentEquipmentLineup, COL_ULD_EQUIPMENT_LINEUP)
                    End If
                    
                    .Cells(rowCurrentLoadUpdate, COL_EQUIPMENT_TYPE_LOAD_UPDATE) = tblContentEquipmentLineup(rowCurrentEquipmentLineup, COL_EQUIPMENT_TYPE_EQUIPMENT_LINEUP)
                    bayDestination = tblContentEquipmentLineup(rowCurrentEquipmentLineup, COL_DESTINATION_EQUIPMENT_LINEUP)
                    bayDestinationSort = tblContentEquipmentLineup(rowCurrentEquipmentLineup, COL_DESTINATION_SORT_EQUIPMENT_LINEUP)
                    .Cells(rowCurrentLoadUpdate, COL_DESTINATION_LOAD_UPDATE) = bayDestination & "-" & bayDestinationSort
                End With
                
                rowCurrentLoadUpdate = rowCurrentLoadUpdate + 1
            End If
        Next rowCurrentEquipmentLineup
        
        wbEquipmentLineup.Close False
        tblBuilder.CreateTable wsLoadUpdate
        wbLoadUpdate.SaveAs ThisWorkbook.Path & "\Data\LoadUpdates\LoadUpdate_" & Format(Date, "yyyymmdd")
        wbLoadUpdate.Close
    Else
        Exit Sub
    End If

    Application.ScreenUpdating = True

End Sub


Private Function IsUPSTrailer(ByVal uld As String) As Boolean
    
   IsUPSTrailer = (Len(uld) = 6 And _
                  IsNumeric(uld) And _
                  (Left(uld, 1) <> "0" And Left(uld, 1) <> "2" And Left(uld, 1) <> "4"))

End Function


Private Sub SetHeadersLoadUpdate(ByRef ws As Worksheet)

    With ws
        .Cells(1, COL_AREA_LOAD_UPDATE) = "AREA"
        .Cells(1, COL_BAY_LOAD_UPDATE) = "BAY"
        .Cells(1, COL_ULD_LOAD_UPDATE) = "ULD"
        .Cells(1, COL_EQUIPMENT_TYPE_LOAD_UPDATE) = "EQUIP"
        .Cells(1, COL_DESTINATION_LOAD_UPDATE) = "DEST"
        .Cells(1, COL_MEASURE_LOAD_UPDATE) = "MEASURE"
    End With

End Sub

