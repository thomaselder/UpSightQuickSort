Attribute VB_Name = "Module2"
Option Explicit

Sub DashTableFormat()
Attribute DashTableFormat.VB_ProcData.VB_Invoke_Func = " \n14"
'
' DashTableFormat Macro
'

'
    ActiveWindow.SmallScroll Down:=-111
    Range("Table1[[#Headers],[LOAD_ID]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "0"
    With Selection
        .HorizontalAlignment = xlRight
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 1
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("Table1[[#Headers],[LOAD_DATE]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "yyyy-mm-dd"
    With Selection
        .HorizontalAlignment = xlRight
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 1
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("Table1[[#Headers],[SORT]]").Select
    ActiveWindow.SmallScroll Down:=105
    Range("Table1[[#All],[SORT]:[AREA]]").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 1
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("Table1[[#Headers],[BAY]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection
        .HorizontalAlignment = xlRight
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 1
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("Table1[[#Headers],[DESTINATION]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("Table1[[#All],[DESTINATION]:[EQUIPMENT]]").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 1
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("Table1[[#Headers],[START_PCT]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("Table1[[#All],[START_PCT]:[END_PCT]]").Select
    Selection.NumberFormat = "0.00"
    With Selection
        .HorizontalAlignment = xlRight
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 1
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("Table1[[#Headers],[NET_VOLUME]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "0"
    With Selection
        .HorizontalAlignment = xlRight
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 1
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("Table1[[#Headers],[STATUS]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection
        .HorizontalAlignment = xlLeft
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 1
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("Table1[[#Headers],[LOAD_ID]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    With Selection.Font
        .Name = "Courier New"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    ActiveWindow.SmallScroll Down:=120
    Range("K128").Select
End Sub
