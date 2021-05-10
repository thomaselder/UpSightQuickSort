VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CubeStatsModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("QuickSort.Models")
Option Explicit

Private subjectName As String
Private netVolume As Long
Private wallCount As Double
Private packagesPerWall As Long
Private utilizationPercent As Double


Public Property Get Subject() As String

    Subject = subjectName

End Property


Public Property Let Subject(ByVal statSubject As String)

    subjectName = statSubject

End Property


Public Property Get Volume() As Long

    Volume = netVolume

End Property


Public Property Let Volume(ByVal packageVolume As Long)

    netVolume = packageVolume

End Property


Public Property Get Walls() As Double

    Walls = wallCount

End Property


Public Property Let Walls(ByVal packageWalls As Double)

    wallCount = packageWalls

End Property


Public Property Get PPW() As Long

    PPW = packagesPerWall

End Property


Public Property Let PPW(ByVal ppwCount As Long)

    packagesPerWall = ppwCount

End Property


Public Property Get Utilization() As Double

    Utilization = utilizationPercent

End Property


Public Property Let Utilization(ByVal utilizationResult As Double)

    utilizationPercent = utilizationResult

End Property
