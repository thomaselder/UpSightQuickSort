VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FilePickerDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("QuickSort.Utilities")
Option Explicit

Private absoluteFileName As String
Private selected As Boolean


Private Sub Class_Initialize()

    absoluteFileName = vbNullString

End Sub


Public Property Get fileNameFull() As String

    fileNameFull = absoluteFileName

End Property


Public Property Get FileSelected() As Boolean

    FileSelected = selected
    
End Property


Public Sub OpenDialog()

    With Application.fileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        If .Show = -1 Then
            selected = True
            absoluteFileName = .SelectedItems(1)
        End If
    End With

End Sub
