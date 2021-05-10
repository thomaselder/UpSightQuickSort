Attribute VB_Name = "Module1"
'@Folder("QuickSort")
'@IgnoreModule
Option Explicit



Public Sub CreateSortOpReport()
    
    Dim stats As StatsFactory
    Dim dataAccess As DataAccessFactory
    Dim loadArchive As LoadArchiveAccessor
    Dim cubeStats As CubeUtilizationStats
    Dim reportTables As CubeReportTables
    Dim targetLoads As Scripting.Dictionary
    Dim sortStatTotals As CubeStatsModel
    Dim pdStatTotals As Scripting.Dictionary
    Dim areaStatTotals As Scripting.Dictionary
    Dim bayStatTotals As Scripting.Dictionary
    Dim destinationStatTotals As Scripting.Dictionary

    Set stats = New StatsFactory
    Set dataAccess = New DataAccessFactory
    Set cubeStats = stats.BuildCubeUtilizationStats
    Set loadArchive = dataAccess.BuildLoadArchiveAccessor
    
    loadArchive.SetSortName Sheet1.ActiveSortComboBox.Value
    Set targetLoads = loadArchive.DateQuery(#4/28/2021#)
    Set reportTables = stats.BuildCubeReportTables
    
    With cubeStats
        .SetLoadData targetLoads
        Set sortStatTotals = .GetSortTotals
        Set pdStatTotals = .GetPrimaryDirectTotals
        Set areaStatTotals = .GetAreaTotals
        Set bayStatTotals = .GetBayTotals
        Set destinationStatTotals = .GetDestinationTotals
    End With
    
    reportTables.SetStatsTables pdStatTotals, areaStatTotals, bayStatTotals, destinationStatTotals

End Sub
