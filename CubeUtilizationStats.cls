VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CubeUtilizationStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("QuickSort.Stats")
Option Explicit


Private models As DataModelsFactory

Private loads As Scripting.Dictionary
Private cubeStats As Scripting.Dictionary
Private isDataSet As Boolean

Private Enum CubeStatsType
    pdStats = 1
    areaStats = 2
    bayStats = 3
    destinationStats = 4
End Enum


Private Sub SetCubeStats(ByVal statsType As Long)

    Dim statsModel As CubeStatsModel
    Dim loadsKey As Variant
    Dim statsKey As Variant
    Dim load As LoadModel
    
     If Not isDataSet Then
        MsgBox "CubeUtilizationStats data is not set.", vbCritical, "Application-Defined Error"
        Exit Sub
    End If
    
    Set cubeStats = New Scripting.Dictionary
    
    For Each loadsKey In loads
        Set load = loads(loadsKey)
        
        statsKey = GetCubeStatsKey(load, statsType)
       
        If cubeStats.Exists(statsKey) Then
            With cubeStats(statsKey)
                .Volume = .Volume + load.Volume
                .Walls = .Walls + load.Walls
            End With
        Else
            Set statsModel = models.BuildCubeStatsModel
            
            With statsModel
                .Subject = statsKey
                .Volume = load.Volume
                .Walls = load.Walls
            End With
            
            cubeStats.Add statsKey, statsModel
        End If
    Next loadsKey

    For Each statsKey In cubeStats.Keys
        With cubeStats(statsKey)
            .PPW = .Volume / .Walls
            .Utilization = .PPW / 65
        End With
    Next statsKey
    
End Sub

Private Function GetCubeStatsKey(ByVal load As LoadModel, ByVal statsType As Long) As String

    Dim statsKey As String

    If statsType = CubeStatsType.pdStats Then
        statsKey = CStr(Left(load.Area, 1))
    ElseIf statsType = CubeStatsType.areaStats Then
        statsKey = load.Area
    ElseIf statsType = CubeStatsType.bayStats Then
        statsKey = CStr(load.Bay)
    ElseIf statsType = CubeStatsType.destinationStats Then
        statsKey = load.Destination
    End If

    GetCubeStatsKey = statsKey
    
End Function

Public Function GetSortTotals() As CubeStatsModel

    Dim statsModel As CubeStatsModel
    Dim load As LoadModel
    Dim loadsKey As Variant
    
    If Not isDataSet Then
        MsgBox "CubeUtilizationStats data is not set.", vbCritical, "Application-Defined Error"
        Exit Function
    End If
    
    Set statsModel = models.BuildCubeStatsModel
    
    For Each loadsKey In loads.Keys
        Set load = loads(loadsKey)
        
        With statsModel
            .Volume = .Volume + load.Volume
            .Walls = .Walls + load.Walls
        End With
    Next loadsKey
    
    With statsModel
        .PPW = .Volume / .Walls
        .Utilization = .PPW / 65
    End With
    
    Set GetSortTotals = statsModel
    
End Function


Public Function GetPrimaryDirectTotals() As Scripting.Dictionary

    SetCubeStats CubeStatsType.pdStats
    
    Set GetPrimaryDirectTotals = cubeStats
    
End Function


Public Function GetAreaTotals() As Scripting.Dictionary

    SetCubeStats CubeStatsType.areaStats

    Set GetAreaTotals = cubeStats

End Function


Public Function GetBayTotals() As Scripting.Dictionary

    SetCubeStats CubeStatsType.bayStats
    
    Set GetBayTotals = cubeStats

End Function


Public Function GetDestinationTotals() As Scripting.Dictionary

    SetCubeStats CubeStatsType.destinationStats
    
    Set GetDestinationTotals = cubeStats

End Function


Public Sub SetLoadData(ByVal cubeLoads As Scripting.Dictionary)

    Dim load As LoadModel

    If Not IsValidLoadData(cubeLoads) Then
        Exit Sub
    End If
    
    Set loads = cubeLoads
    
    isDataSet = True

End Sub


Private Function IsValidLoadData(ByRef cubeLoads As Scripting.Dictionary) As Boolean

    Dim isValid As Boolean
    
    isValid = True
    
    If cubeLoads Is Nothing Then
        MsgBox "CubeUtilizationStats.cubeLoads uninitialized", vbCritical, "Application-Defined Error"
        isValid = False
    ElseIf cubeLoads.Count = 0 Then
        MsgBox "CubeUtilizationStats.cubeLoads contains no data", vbExclamation, "Application-Defined Error"
        isValid = False
    ElseIf Not TypeName(cubeLoads.Items(0)) = "LoadModel" Then
        MsgBox "CubeUtilizationStats.cubeLoads contains invalid data type", vbCritical, "Application-Defined Error"
        isValid = False
    End If

    IsValidLoadData = isValid

End Function


Private Sub Class_Initialize()

    isDataSet = False
    Set models = New DataModelsFactory
    
End Sub
