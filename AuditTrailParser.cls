VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AuditTrailParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("QuickSort.FileParsers")
Option Explicit


Private Const COL_TIMESTAMP As Long = 2
Private Const COL_BAY_NUMBER As Long = 3
Private Const COL_TRAILER_NUMBER As Long = 4
Private Const COL_LOAD_MESSAGE As Long = 5
Private Const COL_MESSAGE_ID As Long = 6

Private loads As Scripting.Dictionary 'read-only property variable
Private auditLineup As Scripting.Dictionary



Public Property Get auditTrailLoads() As Scripting.Dictionary

    Set auditTrailLoads = loads

End Property


Public Property Get targetLineup() As Scripting.Dictionary

    Set targetLineup = auditLineup

End Property


Public Property Set targetLineup(ByRef lineupReference As Scripting.Dictionary)

    Set auditLineup = lineupReference

End Property


Public Sub Parse(ByRef tbl As Variant, ByVal reportDate As Date, ByVal sortAlphaSlic As String)
        
    Const MSG_ID_LOAD_UPDATE_USER As Long = 1013
    Const MSG_ID_LOAD_UPDATE_AUTO As Long = 1022
    Const MSG_ID_SPOT_REQUEST As Long = 1032
    
    Dim load As LoadModel
    Dim loadFactory As DataModelsFactory
    Dim utilities As UtilitiesFactory
    Dim equipmentLookup As EquipmentTypeLookup
    
    Dim updatePercent As Double
    
    Dim bayNumber As String
    Dim trailerNumber As String
    Dim msgId As Long
    
    Dim row As Long
    
    Set loadFactory = New DataModelsFactory
    Set utilities = New UtilitiesFactory
    Set loads = New Scripting.Dictionary
    Set equipmentLookup = utilities.BuildEquipmentTypeLookup

    For row = LBound(tbl, 1) To UBound(tbl, 1)
        bayNumber = tbl(row, COL_BAY_NUMBER)
        trailerNumber = Trim(tbl(row, COL_TRAILER_NUMBER))
        
        If auditLineup.Exists(bayNumber) And _
           trailerNumber <> vbNullString Then
            If IsUPSTrailer(trailerNumber) Then
                trailerNumber = "UPST" & trailerNumber
            End If
            
            If Not loads.Exists(trailerNumber) Then
                Set load = loadFactory.BuildLoadModel
                With load
                    .loadDate = reportDate
                    .Sort = sortAlphaSlic
                    .Area = auditLineup(bayNumber).workArea
                    .Bay = bayNumber
                    .trailerNumber = trailerNumber
                    .Destination = auditLineup(bayNumber).Destination
                    .equipmentType = equipmentLookup.GetEquipmentType(trailerNumber)
                End With
            Else
                Set load = loads(trailerNumber)
            End If
            
            msgId = tbl(row, COL_MESSAGE_ID)
            
            If msgId = MSG_ID_SPOT_REQUEST Then
                load.StartPercentage = 0
            ElseIf msgId = MSG_ID_LOAD_UPDATE_USER Or msgId = MSG_ID_LOAD_UPDATE_AUTO Then
                updatePercent = GetLoadUpdatePercent(tbl(row, COL_LOAD_MESSAGE))
                                
                If updatePercent > load.EndPercentage Then
                    load.EndPercentage = updatePercent
                End If
            End If
            
            If Not loads.Exists(trailerNumber) Then
                loads.Add load.trailerNumber, load
            Else
                Set loads(load.trailerNumber) = load
            End If
        End If
    Next row
        
    
End Sub

Private Function GetLoadUpdatePercent(ByVal eventMessage As String) As Double
    
    Const COL_LOAD_UPDATE As Long = 6
    
    Dim msgSplit() As String
    Dim loadUpdateMsg As String
    Dim loadPercentage As Double
    
    msgSplit = Split(eventMessage, " ")
    loadUpdateMsg = msgSplit(COL_LOAD_UPDATE)
                
    If Right(loadUpdateMsg, 1) = "," Then
        loadPercentage = (CDbl(Left(loadUpdateMsg, Len(loadUpdateMsg) - 2))) / 100
    ElseIf Right(loadUpdateMsg, 1) = "%" Then
        loadPercentage = (CDbl(Left(loadUpdateMsg, Len(loadUpdateMsg) - 1))) / 100
    Else
        loadPercentage = CDbl(loadUpdateMsg) / 100
    End If

    GetLoadUpdatePercent = loadPercentage

End Function


Private Function IsUPSTrailer(ByVal uld As String) As Boolean
    
   IsUPSTrailer = (Len(uld) = 6 And _
                  IsNumeric(uld) And _
                  (Left(uld, 1) <> "0" And Left(uld, 1) <> "2" And Left(uld, 1) <> "4"))

End Function


