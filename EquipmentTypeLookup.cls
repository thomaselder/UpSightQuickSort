VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EquipmentTypeLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("QuickSort.Utilities")
Option Explicit

Private equipmentTypeLength As Scripting.Dictionary

Public Function GetEquipmentType(ByVal uld As String) As String

    Dim equipType As String
    
    If Left(uld, 4) = "UPST" Then
        If Left(Right(uld, 6), 1) = "1" Then
            equipType = "ZZZ"
        ElseIf Left(Right(uld, 6), 1) = "3" Or Left(Right(uld, 6), 1) = "5" Then
            equipType = "WW"
        ElseIf Left(Right(uld, 6), 1) = "6" Then
            equipType = "LC"
        ElseIf Left(Right(uld, 6), 2) = "72" Then
            equipType = "X"
        ElseIf Left(Right(uld, 6), 2) = "74" Then
            equipType = "F"
        ElseIf Left(Right(uld, 6), 1) = "8" Then
            equipType = "ZZZ"
        ElseIf Left(Right(uld, 6), 1) = "9" Then
            equipType = "XXX"
        End If
    ElseIf Left(uld, 4) = "CTLZ" Then
        equipType = "PMS"
    ElseIf Left(uld, 4) = "EJGZ" Then
        equipType = "PRT"
    ElseIf Left(uld, 4) = "MECZ" Then
        equipType = "PMS"
    ElseIf Left(uld, 4) = "ECDU" Then
        equipType = "PMC"
    ElseIf Left(uld, 4) = "MLNU" Then
        equipType = "PMC"
    ElseIf Left(uld, 4) = "MEQU" Then
        equipType = "PMC"
    ElseIf Left(uld, 4) = "MNLU" Then
        equipType = "PMC"
    ElseIf Left(uld, 4) = "BOWZ" Then
        equipType = "BOW"
    ElseIf Left(uld, 4) = "MEQZ" Then
        equipType = "BOW"
    ElseIf Left(uld, 1) = "U" And Len(uld) = 6 And IsNumeric(Right(uld, 5)) Then
        equipType = "BOW"
    ElseIf Left(uld, 4) = "MLNZ" Then
        equipType = "PMS"
    ElseIf Left(uld, 4) = "PCAZ" Then
        equipType = "PMS"
    ElseIf Left(uld, 4) = "PIFZ" Then
        equipType = "PPR"
    ElseIf Left(uld, 4) = "POTZ" Then
        equipType = "PPR"
    ElseIf Left(uld, 4) = "PRTZ" Then
        equipType = "PPR"
    ElseIf Left(uld, 4) = "PTLZ" Then
        equipType = "PPR"
    ElseIf Left(uld, 4) = "UMXU" Then
        equipType = "BCS"
    ElseIf Left(uld, 4) = "USAZ" Then
        equipType = "PMS"
    ElseIf Left(uld, 4) = "EJGZ" Then
        equipType = "PMS"
    ElseIf Left(uld, 4) = "CATZ" Then
        equipType = "PPR"
    ElseIf Left(uld, 4) = "EMHU" Then
        equipType = "BUP"
    ElseIf Left(uld, 4) = "MAR" Then
        equipType = "TMF"
    Else
        MsgBox "Equipment type for " & uld & " not found!", vbExclamation, "Application-Defined Error"
    End If
    
    GetEquipmentType = equipType


End Function


Public Function GetEquipmentLength(ByRef equipTypeCode As String) As Long

    Dim equipmentLength As Long

    equipmentLength = 0

    If equipmentTypeLength.Exists(equipTypeCode) Then
        equipmentLength = equipmentTypeLength(equipTypeCode)
    End If
    
    GetEquipmentLength = equipmentLength
    
End Function


Private Sub SetTrailerLengthLookup()

    Dim wsEquipmentTypes As Worksheet
    Dim tblEquipmentTypes As ListObject
    
    Dim tblContentEquipmentTypes As Variant
    
    Dim rowCurrent As Long
    Const COL_EQUIPMENT_TYPE As Long = 1
    Const COL_EQUIPMENT_TYPE_LENGTH As Long = 3
    
    Set wsEquipmentTypes = ThisWorkbook.Worksheets("EquipmentTypes")
    Set tblEquipmentTypes = wsEquipmentTypes.ListObjects(1)
    
    tblContentEquipmentTypes = tblEquipmentTypes.DataBodyRange
        
    Set equipmentTypeLength = New Scripting.Dictionary
    
    For rowCurrent = LBound(tblContentEquipmentTypes, 1) To UBound(tblContentEquipmentTypes, 1)
        If Not equipmentTypeLength.Exists(tblContentEquipmentTypes(rowCurrent, COL_EQUIPMENT_TYPE)) Then
            equipmentTypeLength.Add tblContentEquipmentTypes(rowCurrent, COL_EQUIPMENT_TYPE), _
                                    tblContentEquipmentTypes(rowCurrent, COL_EQUIPMENT_TYPE_LENGTH)
            
        End If
    Next rowCurrent
    

End Sub


Private Sub Class_Initialize()

    SetTrailerLengthLookup

End Sub
