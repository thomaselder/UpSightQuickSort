VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoadArchiveAccessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("QuickSort.DataAccess")
Option Explicit

Private Const COL_LOAD_DATE As Long = 1
Private Const COL_SORT_ID As Long = 2
Private Const COL_LOAD_ID As Long = 3
Private Const COL_SORT As Long = 4
Private Const COL_AREA As Long = 5
Private Const COL_BAY As Long = 6
Private Const COL_DESTINATION As Long = 7
Private Const COL_ULD As Long = 8
Private Const COL_EQUIPMENT_TYPE As Long = 9
Private Const COL_START_PCT As Long = 10
Private Const COL_END_PCT As Long = 11
Private Const COL_RESP_PCT As Long = 12
Private Const COL_VOLUME As Long = 13
Private Const COL_WALLS As Long = 14
Private Const COL_PPW As Long = 15

Private utilities As UtilitiesFactory
Private sorter As TableSorter
Private models As DataModelsFactory
Private sortName As String
Private loadIDMax As Long

Public Sub SetSortName(ByVal sortAlphaSlic As String)

    sortName = sortAlphaSlic

End Sub

Private Sub Class_Initialize()
    
    Set utilities = New UtilitiesFactory
    Set sorter = utilities.BuildTableSorter
    Set models = New DataModelsFactory
    sortName = vbNullString
    
    
    
End Sub


Public Function DateQuery(ByVal loadDate As Date) As Scripting.Dictionary

     If sortName = vbNullString Then
        MsgBox "LoadArchiveAccesor.sortName Not Set", vbExclamation, "Application-Defined Error"
        Exit Function
    End If
    
    Dim wsArchive As Worksheet
    Dim tblArchive As ListObject
    Dim results As Scripting.Dictionary
    Dim load As LoadModel
    
    Dim rowFirstArchive As Long
    Dim rowLastArchive As Long
    Dim rowCurrentArchive As Long
    
    Set wsArchive = ThisWorkbook.Worksheets(GetArchiveName)
    Set tblArchive = wsArchive.ListObjects(1)
    
    If tblArchive.DataBodyRange Is Nothing Then
        MsgBox "LoadArchiveAccessor archive is empty", vbExclamation, "Application-Defined Error"
        Exit Function
    End If
    
    With tblArchive.DataBodyRange
        rowFirstArchive = .row
        rowLastArchive = .row + .Rows.Count - 1
    End With
    
    Set results = New Scripting.Dictionary
    
    For rowCurrentArchive = rowFirstArchive To rowLastArchive
        If CDate(wsArchive.Cells(rowCurrentArchive, COL_LOAD_DATE).Value2) = loadDate Then
            Set load = models.BuildLoadModel
            
            With load
                .LoadID = wsArchive.Cells(rowCurrentArchive, COL_LOAD_ID).Value2
                .loadDate = wsArchive.Cells(rowCurrentArchive, COL_LOAD_DATE).Value2
                .Sort = wsArchive.Cells(rowCurrentArchive, COL_SORT).Value2
                .Area = wsArchive.Cells(rowCurrentArchive, COL_AREA).Value2
                .Bay = wsArchive.Cells(rowCurrentArchive, COL_BAY).Value2
                .Destination = wsArchive.Cells(rowCurrentArchive, COL_DESTINATION).Value2
                .trailerNumber = wsArchive.Cells(rowCurrentArchive, COL_ULD).Value2
                .equipmentType = wsArchive.Cells(rowCurrentArchive, COL_EQUIPMENT_TYPE).Value2
                .StartPercentage = wsArchive.Cells(rowCurrentArchive, COL_START_PCT).Value2
                .EndPercentage = wsArchive.Cells(rowCurrentArchive, COL_END_PCT).Value2
                .Volume = wsArchive.Cells(rowCurrentArchive, COL_VOLUME).Value2
                .Walls = wsArchive.Cells(rowCurrentArchive, COL_WALLS).Value2
                .packagesPerWall = wsArchive.Cells(rowCurrentArchive, COL_PPW).Value2
            End With
            
            results.Add load.LoadID, load
        End If
    
    Next rowCurrentArchive
    
    Set DateQuery = results

End Function


Public Sub ArchiveLoads(ByRef loads As Scripting.Dictionary)
    
    Dim equipmentType As EquipmentTypeLookup
    
    Set equipmentType = New EquipmentTypeLookup
    
    If sortName = vbNullString Then
        MsgBox "LoadArchiveAccesor.sortName Not Set", vbExclamation, "Application-Defined Error"
        Exit Sub
    End If
    
    Dim wsArchive As Worksheet
    Dim tblArchive As ListObject
    
    Dim load As LoadModel
    Dim loadsKey As Variant
    Dim archiveName As String
    Dim rowInsert As Long
    
    archiveName = GetArchiveName
    
    Set wsArchive = ThisWorkbook.Worksheets(archiveName)
    Set tblArchive = wsArchive.ListObjects(1)
    
    If Not tblArchive.DataBodyRange Is Nothing Then
        rowInsert = tblArchive.DataBodyRange.row + tblArchive.DataBodyRange.Rows.Count
    Else
        rowInsert = tblArchive.HeaderRowRange.row + 1
    End If
    
    sorter.SortTable wsArchive, tblArchive, COL_LOAD_ID
    
    If Not tblArchive.DataBodyRange Is Nothing Then
        loadIDMax = wsArchive.Cells(tblArchive.DataBodyRange.row + tblArchive.DataBodyRange.Rows.Count - 1, COL_LOAD_ID).Value2
    Else
        loadIDMax = 0
    End If
    
    For Each loadsKey In loads.Keys
        Set load = loads(loadsKey)
        
        With load
            wsArchive.Cells(rowInsert, COL_LOAD_DATE) = .loadDate
            wsArchive.Cells(rowInsert, COL_SORT_ID) = 0
            wsArchive.Cells(rowInsert, COL_LOAD_ID) = loadIDMax + 1
            wsArchive.Cells(rowInsert, COL_SORT) = .Sort
            wsArchive.Cells(rowInsert, COL_AREA) = .Area
            wsArchive.Cells(rowInsert, COL_BAY) = .Bay
            wsArchive.Cells(rowInsert, COL_DESTINATION) = .Destination
            wsArchive.Cells(rowInsert, COL_ULD) = .trailerNumber
            wsArchive.Cells(rowInsert, COL_EQUIPMENT_TYPE) = .equipmentType
            wsArchive.Cells(rowInsert, COL_START_PCT) = .StartPercentage
            wsArchive.Cells(rowInsert, COL_END_PCT) = .EndPercentage
            wsArchive.Cells(rowInsert, COL_RESP_PCT) = .EndPercentage - .StartPercentage
            wsArchive.Cells(rowInsert, COL_VOLUME) = .Volume
            wsArchive.Cells(rowInsert, COL_WALLS) = ((.EndPercentage - .StartPercentage) * equipmentType.GetEquipmentLength(.equipmentType)) / 2
            wsArchive.Cells(rowInsert, COL_PPW) = .Volume / wsArchive.Cells(rowInsert, COL_WALLS)
        End With
        
        loadIDMax = loadIDMax + 1
        rowInsert = rowInsert + 1
    Next loadsKey
    
    

    

End Sub

Private Function GetArchiveName() As String

    Dim archiveName As String
    
    If sortName = "LEXKY-D" Then
        archiveName = "LoadArchiveDay"
    ElseIf sortName = "LEXKY-T" Then
        archiveName = "LoadArchiveTwilight"
    End If

    GetArchiveName = archiveName
    
End Function
