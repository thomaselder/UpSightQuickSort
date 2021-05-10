VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("QuickSort.Utilities")
Option Explicit

Private fileNew As String

Public Property Get fileNameNew() As String

    fileNameNew = fileNew

End Property

Private Property Let fileNameNew(ByVal fileNameFull As String)

    fileNew = fileNameFull

End Property


Private Sub ConvertFile(ByVal fileNameFull As String, ByVal fileFormatType As Variant)

    Dim wb As Workbook
    Dim fso As FileSystemObject
    Dim fileExtensionSplit() As String
    
    Set fso = New FileSystemObject
    
    If Not fso.FileExists(fileNameFull) Then
        Exit Sub
    End If
    
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With
    
    Set wb = Workbooks.Open(fileNameFull)
    
    fileExtensionSplit = Split(fileNameFull, ".")
    
    If fileFormatType = xlCSV Then
        fileNameNew = fileExtensionSplit(0) & ".csv"
    ElseIf fileFormatType = xlWorkbookDefault Then
        fileNameNew = fileExtensionSplit(0) & ".xlsx"
    End If
    
    With wb
        .SaveAs fileNameNew, fileFormatType
        .Close True
    End With
    
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With
End Sub



Public Sub ConvertToCsv(ByVal fileNameFull As String)

    ConvertFile fileNameFull, xlCSV
    
End Sub


Public Sub ConvertToXlsx(ByVal fileNameFull As String)

    ConvertFile fileNameFull, xlWorkbookDefault

End Sub

