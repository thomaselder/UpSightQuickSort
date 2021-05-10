VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileParserFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("QuickSort.Factories")
Option Explicit

Public Function BuildAuditTrailParser() As AuditTrailParser

    Set BuildAuditTrailParser = New AuditTrailParser

End Function


Public Function BuildESRParser() As ESRParser

    Set BuildESRParser = New ESRParser

End Function
