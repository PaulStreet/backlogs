Sub MD_Changes()
' MD_Changes Macro

	'This section of code initializes a string variable for your current workbook filename.
	Dim originalWB As String
	originalWB = ActiveWorkbook.Name
	
	'Original code done by Shawn "P&B" Boyer
    Cells.Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Columns("C:C").Select
    Selection.TextToColumns Destination:=Range("C1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(4, 9)), TrailingMinusNumbers:=True
    Range("A2").Select
    ChDir _
        "C:\Users\boyers\Documents\Broadload\Moving Orders\Extensions\Master Data Changes\2014 MD changes"
    Workbooks.Open Filename:= _
        "C:\Users\boyers\Documents\Broadload\Moving Orders\Extensions\Master Data Changes\2014 MD changes\Template to request Laminate MD changes.xlsx"
    Rows("1:7").Select
    Selection.Copy
    Windows("originalWB").Activate
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
    Range("A7:Q7").Select
    Application.CutCopyMode = False
    Selection.AutoFilter
    Cells.Select
    Cells.EntireColumn.AutoFit
    Cells.EntireColumn.AutoFit
    Cells.EntireColumn.AutoFit
    Range("E8").Select
    ActiveWindow.FreezePanes = True
    Windows("Template to request Laminate MD changes.xlsx").Activate
    Range("K8:Q8").Select
    Selection.Copy
    Windows("originalWB").Activate
    Range("K8").Select
    ActiveSheet.Paste
    Windows("Template to request Laminate MD changes.xlsx").Activate
    ActiveWindow.Close
    Range("E8").Select
End Sub
