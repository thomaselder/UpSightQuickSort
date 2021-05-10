VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoadUpdateModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("QuickSort.Models")
Option Explicit

Private workArea As String
Private bayNumber As String
Private uld As String
Private trailerType As String
Private destinationFacility As String
Private trailerMeasurement As Double
Private trailerPercent As Double

Public Property Get Area() As String

    Area = workArea

End Property


Public Property Let Area(ByVal workAreaCode As String)

    workArea = workAreaCode

End Property


Public Property Get Bay() As String

    Bay = bayNumber

End Property


Public Property Let Bay(ByVal bayID As String)

    bayNumber = bayID

End Property

Public Property Get trailerNumber() As String

    trailerNumber = uld

End Property


Public Property Let trailerNumber(ByVal uldNumber As String)

    uld = uldNumber

End Property


Public Property Get equipmentType() As String

    equipmentType = trailerType

End Property


Public Property Let equipmentType(ByVal equipmentCode As String)

    trailerType = equipmentCode

End Property

Public Property Get Destination() As String

    Destination = destinationFacility

End Property


Public Property Let Destination(ByVal alphaSlic As String)

    destinationFacility = alphaSlic

End Property


Public Property Get Measure() As Double

    Measure = trailerMeasurement
    
End Property


Public Property Let Measure(ByVal feetMeasure As Double)

    trailerMeasurement = feetMeasure

End Property


Public Property Get LoadPercent() As Double

    LoadPercent = trailerPercent

End Property


Public Property Let LoadPercent(ByVal updatePercent As Double)

    trailerPercent = updatePercent

End Property
