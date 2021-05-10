VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LineupModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("QuickSort.Models")
Option Explicit

Private Area As String
Private Bay As String
Private destinationFacility As String

Public Property Get workArea() As String

    workArea = Area

End Property


Public Property Let workArea(ByVal workAreaName As String)

    Area = workAreaName

End Property


Public Property Get bayNumber() As String

    bayNumber = Bay

End Property


Public Property Let bayNumber(ByVal bayID As String)

    Bay = bayID

End Property


Public Property Get Destination() As String

    Destination = destinationFacility

End Property


Public Property Let Destination(ByVal alphaSlic As String)

    destinationFacility = alphaSlic

End Property




