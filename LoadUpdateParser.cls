VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoadUpdateParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("QuickSort.FileParsers")
Option Explicit


Public Function Parse(ByVal fileContent As Variant) As Scripting.Dictionary
    
    Const COL_AREA As Long = 1
    Const COL_BAY As Long = 2
    Const COL_TRAILER_NUMBER As Long = 3
    Const COL_EQUIPMENT_TYPE As Long = 4
    Const COL_DESTINATION As Long = 5
    Const COL_MEASURE As Long = 6
    
    Dim dataAccess As DataAccessFactory
    Dim dataModels As DataModelsFactory
    Dim equipmentTypes As EquipmentTypeLookup
    Dim loadUpdates As Scripting.Dictionary
    Dim loadUpdate As LoadUpdateModel
    
    Dim rowCurrent As Long
    
    Set dataAccess = New DataAccessFactory
    Set dataModels = New DataModelsFactory
    Set equipmentTypes = dataAccess.BuildEquipmentTypeLookup
    Set loadUpdates = New Scripting.Dictionary
    
    For rowCurrent = LBound(fileContent, 1) To UBound(fileContent, 1)
        Set loadUpdate = dataModels.BuildLoadUpdateModel
        
        With loadUpdate
            .Area = Trim(fileContent(rowCurrent, COL_AREA))
            .Bay = CStr(fileContent(rowCurrent, COL_BAY))
            .trailerNumber = Trim(fileContent(rowCurrent, COL_TRAILER_NUMBER))
            .equipmentType = Trim(fileContent(rowCurrent, COL_EQUIPMENT_TYPE))
            .Destination = Trim(fileContent(rowCurrent, COL_DESTINATION))
            .Measure = fileContent(rowCurrent, COL_MEASURE)
            If .Measure > 0 Then
                If equipmentTypes.GetEquipmentLength(.equipmentType) > 0 Then
                    .LoadPercent = 1 - (.Measure / equipmentTypes.GetEquipmentLength(.equipmentType))
                Else
                    MsgBox "Missing equipmentLength for equipment type: " & .equipmentType, vbCritical, "Application-Defined Error"
                End If
            End If
        End With
        
        If Not loadUpdates.Exists(loadUpdate.trailerNumber) Then
            loadUpdates.Add loadUpdate.trailerNumber, loadUpdate
        End If
    Next rowCurrent
    
    Set Parse = loadUpdates
    
End Function
