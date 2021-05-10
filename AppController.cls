VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("QuickSort")
Option Explicit

Private utilities As UtilitiesFactory
Private fileParsers As FileParserFactory
Private models As DataModelsFactory

Private lineupAccess As LineupAccessor
Private filePicker As FilePickerDialog
Private tblBuilder As TableBuilder
Private loadTrackerAccess As LoadTrackerAccessor
Private archiveAccess As LoadArchiveAccessor

Private isGoodState As Boolean

Private Sub Class_Initialize()

    Set utilities = New UtilitiesFactory
    Dim dataAccessors As DataAccessFactory
    Set dataAccessors = New DataAccessFactory
    Set fileParsers = New FileParserFactory
    Set lineupAccess = dataAccessors.BuildLineupAccessor
    Set filePicker = utilities.BuildFilePickerDialog
    Set tblBuilder = utilities.BuildTableBuilder
    Set loadTrackerAccess = dataAccessors.BuildLoadTrackerAccessor
    
    Set models = New DataModelsFactory
    Set archiveAccess = dataAccessors.BuildLoadArchiveAccessor
    
    
End Sub


Public Sub SaveToArchiveButtonClick_EventHandler()

    archiveAccess.SetSortName Sheet1.ActiveSortComboBox.Value
    archiveAccess.ArchiveLoads loadTrackerAccess.ReadAllLoads

End Sub


Public Sub ClearRecordsButtonClick_EventHandler()

    loadTrackerAccess.DeleteAllLoads

End Sub


Public Sub ESRUpdate_EventHandler(ByVal selectedSort As String, ByVal targetSortDate As Date)
    
    Const COL_ULD As Long = 8
    Const COL_VOLUME As Long = 12
    
    Dim parser As ESRParser
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Dim trailerNumbers As Variant
    Dim trailerVolume As Scripting.Dictionary
    
    Dim rowFirst As Long
    Dim rowLast As Long
    Dim rowCurrent As Long
    
    Dim uldKey As String
    
    Set parser = fileParsers.BuildESRParser
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    Set tbl = ws.ListObjects(1)
    
    With tbl.DataBodyRange
        rowFirst = .row
        rowLast = .row + .Rows.Count - 1
    End With
    
    trailerNumbers = ws.Range(ws.Cells(rowFirst, COL_ULD), ws.Cells(rowLast, COL_ULD))
    
    Set trailerVolume = parser.ExtractVolumeData(trailerNumbers, _
                                                  "LEXINGTON", _
                                                  "4059", _
                                                  Split(selectedSort, "-")(1), _
                                                  targetSortDate)
    
    For rowCurrent = rowFirst To rowLast
        uldKey = ws.Cells(rowCurrent, COL_ULD).Value2
        
        If trailerVolume.Exists(uldKey) Then
            ws.Cells(rowCurrent, COL_VOLUME) = trailerVolume(uldKey)
        End If
    
    Next rowCurrent

End Sub

Public Sub InitialUpdateButtonClick_EventHandler()

    
    Dim loadUpdates As Scripting.Dictionary
    
    Set loadUpdates = ReadLoadUpdateForm
    WriteLoadUpdatePercentages 10, loadUpdates

End Sub


Public Sub FinalUpdateButtonClick_EventHandler()
    
    Dim loadUpdates As Scripting.Dictionary
    
    Set loadUpdates = ReadLoadUpdateForm
    WriteLoadUpdatePercentages 11, loadUpdates

End Sub


Private Sub WriteLoadUpdatePercentages(ByVal pctType As Long, ByRef loadUpdates As Scripting.Dictionary)

    Const COL_ULD As Long = 8
    
    Dim loadUpdate As LoadUpdateModel
    Dim load As LoadModel
    Dim tblDashboard As ListObject
    Dim rowFirst As Long
    Dim rowLast As Long
    Dim rowCurrent As Long
    Dim loadUpdatesKey As Variant
    Dim loadFound As Boolean
    
    Set tblDashboard = Sheet1.ListObjects(1)
    
    If Not tblDashboard.DataBodyRange Is Nothing Then
        With tblDashboard.DataBodyRange
            rowFirst = .row
            rowLast = .row + .Rows.Count - 1
        End With
    Else
        rowFirst = tblDashboard.HeaderRowRange.row + 1
        rowLast = tblDashboard.HeaderRowRange.row + 1
    End If
    
    For Each loadUpdatesKey In loadUpdates.Keys
        loadFound = False
        rowCurrent = rowLast
        
        Set loadUpdate = loadUpdates(loadUpdatesKey)
        
        Do While Not loadFound And rowCurrent >= rowFirst
            If Sheet1.Cells(rowCurrent, COL_ULD) = loadUpdate.trailerNumber Then
                loadFound = True
            Else
                rowCurrent = rowCurrent - 1
            End If
        Loop
        
        If loadFound Then
            Sheet1.Cells(rowCurrent, pctType) = loadUpdate.LoadPercent
        Else
            If Not tblDashboard.DataBodyRange Is Nothing Then
                rowCurrent = tblDashboard.DataBodyRange.row + tblDashboard.DataBodyRange.Rows.Count
            Else
                rowCurrent = tblDashboard.HeaderRowRange.row
            End If
            
            Set load = models.BuildLoadModel
            
            With loadUpdate
                    'load.LoadID = appVars.GetAppVar("loadID")
                    load.Sort = Sheet1.ActiveSortComboBox.Value
                    load.loadDate = Format(Sheet1.SortDateComboBox.Value, "yyyy-mm-dd")
                    load.Area = .Area
                    load.Bay = .Bay
                    load.Destination = .Destination
                    load.trailerNumber = .trailerNumber
                    load.equipmentType = .equipmentType
                    
                    If pctType = 10 Then
                        load.StartPercentage = .LoadPercent
                    ElseIf pctType = 11 Then
                        load.EndPercentage = .LoadPercent
                    End If
                    
                    loadTrackerAccess.CreateLoad load
            End With
        End If
    Next loadUpdatesKey
    
End Sub

Private Function ReadLoadUpdateForm() As Scripting.Dictionary

    Dim picker As FilePickerDialog
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim tblContent As Variant
    Dim parser As LoadUpdateParser
    Dim loadUpdates As Scripting.Dictionary
    
    Dim fpath As String
    
    Set picker = New FilePickerDialog
    Set parser = New LoadUpdateParser
    
    With picker
        .OpenDialog
        If .FileSelected Then
            fpath = .fileNameFull
        Else
            Exit Function
        End If
    End With
    
    Set wb = Workbooks.Open(fpath)
    Set ws = wb.Worksheets(1)
    Set tbl = ws.ListObjects(1)
    tblContent = tbl.DataBodyRange
    
    Set loadUpdates = parser.Parse(tblContent)
    
    wb.Close False
    
    Set ReadLoadUpdateForm = loadUpdates

End Function

Public Sub LoadUpdateFormBuilderButtonClick_EventHandler()
    
    Dim fileNameEquipmentLineup As String
    Dim loadUpdate As LoadUpdateForm
    
    Set loadUpdate = New LoadUpdateForm
    
    With filePicker
        .OpenDialog
        If .FileSelected Then
            fileNameEquipmentLineup = .fileNameFull
        Else
            Exit Sub
        End If
    End With
    
    fileNameEquipmentLineup = UpgradeExcelFileVersion(fileNameEquipmentLineup)
    
    loadUpdate.Build fileNameEquipmentLineup
    


End Sub

Public Sub AuditTrailButtonClick_EventHandler()
    
    Dim wbAuditTrail As Workbook
    Dim wsAuditTrail As Worksheet
    Dim tblAuditTrail As ListObject
    Dim lineup As Scripting.Dictionary
    Dim parser As AuditTrailParser
    Dim loads As Scripting.Dictionary
    Dim loadsKey As Variant
    Dim load As LoadModel
    
    Dim tblContentAuditTrail As Variant
    Dim filePath As String
    Dim filePathNew As String
    
    Set parser = fileParsers.BuildAuditTrailParser
    Set lineup = lineupAccess.GetLineup("Lineup", Sheet1.ActiveSortComboBox.Value)
    Set parser.targetLineup = lineup
    
    filePath = GetFileFromUser
    
    If Not isGoodState Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    filePathNew = UpgradeExcelFileVersion(filePath)
    
    Set wbAuditTrail = Workbooks.Open(filePathNew)
    Set wsAuditTrail = wbAuditTrail.Worksheets(1)
    
    wsAuditTrail.Rows(2).Delete
    
    Set tblAuditTrail = tblBuilder.CreateTable(wsAuditTrail)
    tblContentAuditTrail = tblAuditTrail.DataBodyRange
    
    wbAuditTrail.Close False
    
    parser.Parse tblContentAuditTrail, _
                 CDate(Sheet1.SortDateComboBox.Value), _
                 Sheet1.ActiveSortComboBox.Value
    
    Set loads = parser.auditTrailLoads
    
    For Each loadsKey In loads
        Set load = loads(loadsKey)
        
        loadTrackerAccess.CreateLoad load
    Next loadsKey
    
    Application.ScreenUpdating = True
    
End Sub





Private Function UpgradeExcelFileVersion(ByRef filePath As String) As String
    
    Dim converter As FileConverter

    Set converter = utilities.BuildFileConverter

    With converter
        .ConvertToCsv filePath
        filePath = .fileNameNew
        .ConvertToXlsx filePath
        filePath = .fileNameNew
    End With
    
    UpgradeExcelFileVersion = filePath

End Function

Private Function GetFileFromUser() As String
    
    Dim picker As FilePickerDialog
    Dim filePath As String
    
    Set picker = utilities.BuildFilePickerDialog
    
    With picker
        .OpenDialog
        If .FileSelected Then
            filePath = .fileNameFull
            isGoodState = True
        Else
            isGoodState = False
            Exit Function
        End If
    End With

    GetFileFromUser = filePath

End Function



