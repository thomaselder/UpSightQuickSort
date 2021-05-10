VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UtilitiesFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("QuickSort.Factories")
Option Explicit

Public Function BuildFilePickerDialog() As FilePickerDialog

    Set BuildFilePickerDialog = New FilePickerDialog

End Function


Public Function BuildFileConverter() As FileConverter

    Set BuildFileConverter = New FileConverter

End Function


Public Function BuildTableBuilder() As TableBuilder

    Set BuildTableBuilder = New TableBuilder

End Function


Public Function BuildTableSorter() As TableSorter

    Set BuildTableSorter = New TableSorter

End Function


Public Function BuildEquipmentTypeLookup() As EquipmentTypeLookup

    Set BuildEquipmentTypeLookup = New EquipmentTypeLookup

End Function
