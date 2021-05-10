VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataAccessFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("QuickSort.Factories")
Option Explicit

Public Function BuildAppVarsAccessor() As AppVarsAccessor

    Set BuildAppVarsAccessor = New AppVarsAccessor

End Function


Public Function BuildLineupAccessor() As LineupAccessor

    Set BuildLineupAccessor = New LineupAccessor

End Function


Public Function BuildLoadTrackerAccessor() As LoadTrackerAccessor

    Set BuildLoadTrackerAccessor = New LoadTrackerAccessor

End Function


Public Function BuildEquipmentTypeLookup() As EquipmentTypeLookup

    Set BuildEquipmentTypeLookup = New EquipmentTypeLookup

End Function


Public Function BuildLoadArchiveAccessor() As LoadArchiveAccessor

    Set BuildLoadArchiveAccessor = New LoadArchiveAccessor

End Function
