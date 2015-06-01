Sub Z_ORDERRAT()
'Code written by N99610 on 6/1/15
'This section executes code necessary to integrate the order rat and backlog files into one file.

    'This section deletes any FinishedGoods worksheets in preparation for creation of a new sheet.
    Application.DisplayAlerts = False
    On Error Resume Next
    Sheets("FinishedGoods").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    'This section creates a copy of the AllOrders worksheet and names it as a temp file.
    Dim FGTemp As Worksheet
    Sheets("AllOrders").Copy After:=Sheets("AllOrders")
    Set FGTemp = ActiveSheet
    FGTemp.Name = "FinishedGoodsTemp"

    'This important piece of code stores the total number of rows in the All Orders tab.
    Dim lastRowAO As Long
    lastRowAO = Range("A" & Rows.Count).End(xlUp).Row
    
    'The below code creates a helper column based on material number in order to identify finished goods.
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "FERTorHALB"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[-1]>100000000,RC[-1]<200000000),""FERT"",""HALB"")"
    Range("E2").Select
    Selection.AutoFill Destination:=Range("E2:E" & lastRowAO)
    Range("E2:E" & lastRowAO).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    'This section copies and pastes only the finished goods into a new worksheet and names it FinishedGoods.
    ActiveSheet.Range("$A$1:$FB$" & lastRowAO).AutoFilter Field:=5, Criteria1:="FERT"
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("FinishedGoodsTemp").Select
    Sheets.Add
    ActiveSheet.Name = "FinishedGoods"
    ActiveSheet.Paste
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.AutoFilter
    Cells.Select
    With Selection
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit
    Columns("E:E").Select
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
    
    'This section deletes the temp FinishedGoodsTemp sheet.
    Application.DisplayAlerts = False
    On Error Resume Next
    Sheets("FinishedGoodsTemp").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    'This section downloads the order rat information to the backlog.
    'To Be Completed
    
    'This code process the order rat information in the backlog.
    'To Be Completed
    
    'This section sets up the final information pertaining order rat and gives order rat options.
    'To Be Completed

End Sub
