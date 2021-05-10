VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoadModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("QuickSort.Models")
Option Explicit

Private loadIDNumber As Long
Private trailerLoadDate As Date
Private responsibleSort As String
Private parentWorkArea As String
Private parentBay As String
Private destinationFacility As String
Private trailer As String
Private equipmentCode As String
Private loadPercentageInitial As Double
Private loadPercentageFinal As Double
Private packageVolume As Long
Private wallsCount As Double
Private PPW As Long
Private recordStatus As String


Private Sub Class_Initialize()

    loadPercentageInitial = -1
    loadPercentageFinal = -1

End Sub


Public Property Get LoadID() As Long

        LoadID = loadIDNumber
    
End Property


Public Property Let LoadID(ByVal id As Long)

    loadIDNumber = id

End Property


Public Property Get loadDate() As Date

    loadDate = trailerLoadDate
    
End Property


Public Property Let loadDate(ByVal sortDate As Date)

    trailerLoadDate = sortDate

End Property


Public Property Get Sort() As String

    Sort = responsibleSort

End Property


Public Property Let Sort(ByVal alphaSlic As String)

    responsibleSort = alphaSlic

End Property


Public Property Get Area() As String

    Area = parentWorkArea

End Property


Public Property Let Area(ByVal workArea As String)

    parentWorkArea = workArea

End Property


Public Property Get Bay() As String

    Bay = parentBay

End Property


Public Property Let Bay(ByVal bayID As String)

    parentBay = bayID

End Property


Public Property Get Destination() As String



    Destination = destinationFacility

End Property


Public Property Let Destination(ByVal alphaSlic As String)

    destinationFacility = alphaSlic

End Property


Public Property Get trailerNumber() As String

    trailerNumber = trailer

End Property


Public Property Let trailerNumber(ByVal uld As String)

    trailer = uld

End Property


Public Property Get equipmentType() As String

    equipmentType = equipmentCode

End Property


Public Property Let equipmentType(ByVal trailerType As String)

    equipmentCode = trailerType

End Property


Public Property Get StartPercentage() As Double

    StartPercentage = loadPercentageInitial

End Property


Public Property Let StartPercentage(ByVal loadPercentage As Double)

    loadPercentageInitial = loadPercentage

End Property


Public Property Get EndPercentage() As Double

    EndPercentage = loadPercentageFinal

End Property


Public Property Let EndPercentage(ByVal loadPercentage As Double)

    loadPercentageFinal = loadPercentage

End Property


Public Property Get Volume() As Long

    Volume = packageVolume

End Property


Public Property Let Volume(ByVal netVolume As Long)

    packageVolume = netVolume

End Property


Public Property Get Walls() As Double

    Walls = wallsCount

End Property


Public Property Let Walls(ByVal wallsBuilt As Double)

    wallsCount = wallsBuilt

End Property


Public Property Get packagesPerWall() As Long

    packagesPerWall = PPW

End Property


Public Property Let packagesPerWall(ByVal ppwCount As Long)

    PPW = ppwCount

End Property


Public Property Get Status() As String

    Status = recordStatus

End Property


Public Property Let Status(ByVal currentStatus As String)

    recordStatus = currentStatus

End Property


