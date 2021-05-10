VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ESRParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("QuickSort.FileParsers")
Option Explicit

Private Const ESR_URL As String = "https://igate.inside.ups.com/IGateway/apps/index.cfm?action=IntlOps.Home.Main&strFuncName15543=ULDActivity&strFuncAction15543=ULDActivity&intReportLocationID=6485183463415122666&intAutoNavigate=1&strULD="
Private Const COL_SORT_DATE As Long = 0
Private Const COL_ORIGIN_CENTER As Long = 1
Private Const COL_RESP_SORT As Long = 2
Private Const COL_LOOSE_VOLUME As Long = 10
Private Const COL_BAGS_LINKED As Long = 11
Private Const COL_SYSTEM_MSG As Long = 2
Private Const ROW_SYSTEM_MSG As Long = 7

Private originFacility As String
Private responsibleSort As String
Private loosePackages As Long
Private bags As Long


'Read-Only`
Public Property Get looseVolume() As Long

    looseVolume = loosePackages

End Property


'Read-Only
Public Property Get bagsLinked() As Long

    bagsLinked = bags

End Property


Public Function ExtractVolumeData(ByVal trailerNumbers As Variant, _
                                  ByVal facilityName As String, _
                                  ByVal facilitySlic As String, _
                                  ByVal respSort As String, _
                                  ByVal targetDate As Date) As Scripting.Dictionary

On Error GoTo EH

    
    
    Dim ie As InternetExplorerMedium
    Dim htmlDoc As HTMLDocument
    Dim element As Variant
    Dim elements As Object
    Dim trailerVolume As Scripting.Dictionary
    Dim i As Long
    Dim wsDashboard As Worksheet
    Dim trailerNumber As String
    
    Set wsDashboard = ThisWorkbook.Worksheets("Dashboard")
    
    SetOriginFacility facilityName, facilitySlic
    SetResponsibleSort respSort
    Set ie = New InternetExplorerMedium
    Set trailerVolume = New Scripting.Dictionary
    
    With ie
        .Visible = False
        For i = LBound(trailerNumbers, 1) To UBound(trailerNumbers, 1)
            wsDashboard.Cells(ROW_SYSTEM_MSG, COL_SYSTEM_MSG) = CStr(Round((i / UBound(trailerNumbers, 1)) * 100, 2)) & "% Complete. Searching for trailer " & trailerNumbers(i, 1)
            trailerNumber = trailerNumbers(i, 1)
            
            If trailerNumber <> "EMPTY" Then
                Application.ScreenUpdating = False
                
                .navigate ESR_URL & trailerNumbers(i, 1) 'ToDO: Validate trailerNumber
                
                Do While .readyState <> READYSTATE_COMPLETE Or ie.Busy
                    DoEvents
                Loop
                
                Set htmlDoc = .document
                Set elements = htmlDoc.getElementsByTagName("tr")
                
                loosePackages = 0
                bags = 0
                
                For Each element In elements
                    If IsDataBodyRow(element.rowIndex) And IsTargetRecord(element) Then
                        If element.Cells(COL_SORT_DATE).innerText = Format(targetDate, "m/d/yy") Then
                            SetVolumeData element
                            Exit For
                        End If
                    End If
                Next
                
                If Not trailerVolume.Exists(trailerNumbers(i, 1)) Then
                    trailerVolume.Add trailerNumbers(i, 1), loosePackages + bags
                Else
                    Debug.Print "ESRScraper duplicate trailer number: " & trailerNumbers(i, 1)
                End If
                
                Application.ScreenUpdating = True
            End If
        Next i
        .Quit
    End With
    
    Application.ScreenUpdating = True
    
    Set ExtractVolumeData = trailerVolume
    
    Exit Function
EH:
    
    Debug.Print Err.Number & " " & Err.Description & " " & Err.Source
    Err.Clear
    
  

End Function


Private Sub SetVolumeData(ByVal htmlTableRow As Variant)

    loosePackages = CLng(htmlTableRow.Cells(COL_LOOSE_VOLUME).innerText)
    bags = CLng(htmlTableRow.Cells(COL_BAGS_LINKED).innerText)

End Sub


Private Sub SetOriginFacility(ByVal facilityName As String, ByVal facilitySlic As String)

    originFacility = "(" & facilitySlic & ") " & facilityName

End Sub


Private Sub SetResponsibleSort(ByVal respSort As String)

    responsibleSort = respSort

End Sub


Private Function IsTargetRecord(ByVal htmlTableRow As Object) As Boolean

    IsTargetRecord = (Trim(htmlTableRow.Cells(COL_ORIGIN_CENTER).innerText) = originFacility) And _
                     (Trim(htmlTableRow.Cells(COL_RESP_SORT).innerText) = responsibleSort)
    
End Function


Private Function IsDataBodyRow(ByVal rowIndex As Long) As Boolean

    IsDataBodyRow = (rowIndex > 1)

End Function


'@Ignore ProcedureNotUsed
Private Function ESRIsOpen() As Boolean

    Const ESR_PAGE_TITLE As String = "Operations - Home - Enterprise Scan Reports on the iGate Framework"
    
    Dim shellApp As Object
    
    Dim ieCount As Variant
    Dim isOpen As Boolean
    Dim windowTitle As String
    
    Dim x As Variant
    
    Set shellApp = CreateObject("Shell.Application")
    
    ieCount = shellApp.Windows.Count
    isOpen = False
    
    Do While x < ieCount And Not isOpen
        '@Ignore UnhandledOnErrorResumeNext
        On Error Resume Next
        windowTitle = shellApp.Windows(x).document.Title
        
        If windowTitle = ESR_PAGE_TITLE Then
            isOpen = True
        End If
        
        x = x + 1
    Loop
    
    ESRIsOpen = isOpen

End Function



