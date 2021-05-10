VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataModelsFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("QuickSort.Factories")
Option Explicit

Public Function BuildLineupModel() As LineupModel

    Set BuildLineupModel = New LineupModel

End Function


Public Function BuildLoadModel() As LoadModel
    
    Set BuildLoadModel = New LoadModel

End Function


Public Function BuildLoadUpdateModel() As LoadUpdateModel

    Set BuildLoadUpdateModel = New LoadUpdateModel

End Function


Public Function BuildCubeStatsModel() As CubeStatsModel

    Set BuildCubeStatsModel = New CubeStatsModel

End Function
