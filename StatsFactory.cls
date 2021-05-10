VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StatsFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("QuickSort.Factories")
Option Explicit

Public Function BuildCubeUtilizationStats() As CubeUtilizationStats

    Set BuildCubeUtilizationStats = New CubeUtilizationStats

End Function

Public Function BuildCubeReportTables() As CubeReportTables

    Set BuildCubeReportTables = New CubeReportTables

End Function
