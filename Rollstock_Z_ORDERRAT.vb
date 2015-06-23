Sub Z_ORDERRAT_FGSORT()
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
    ActiveSheet.Range("$A$1:$CC$" & lastRowAO).AutoFilter Field:=5, Criteria1:="FERT"
    Cells.EntireColumn.Hidden = False
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
    
End Sub

Sub Z_ORDERRAT_IMPORT()
'This function downloads the Order Rat file from the SAS server, formats it and adds a worksheet to the backlog with the data.
'Coded by N99610 on 6/1/15

    'This section deletes current Order Rat sheet if it exists.
    Sheets("Order Rat").Select
    Cells.ClearContents
    'Application.DisplayAlerts = False
    'On Error Resume Next
    'Sheets("Order Rat").Delete
    'Application.DisplayAlerts = True
    'On Error GoTo 0
    
    'This section imports the Order Rat file from the Duncan SAS server as coordinated with CSIM Team.
    Workbooks.OpenText Filename:="PATH OMITTED"

    'This section counts the number of rows and stores it as a variable.
    Dim lastRowORI As Long
    lastRowORI = Range("A" & Rows.Count).End(xlUp).Row

    'This section does minor formatting to the Order Rat source file.
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "On-Hand"
    Range("O1").Select
    ActiveCell.FormulaR1C1 = "Avg Shipment"
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "Days Between Shipments"
    Columns("A:B").Select
    Selection.Insert Shift:=xlToRight
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "key"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "DOH"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=RC[2]&RC[5]"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=RC[7]/(RC[15]/RC[16])"
    Range("A2:B2").AutoFill Destination:=Range("A2:B" & lastRowORI)
    Columns("A:B").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:B").EntireColumn.AutoFit
    Columns("A:B").EntireColumn.AutoFit
    Columns("B:B").Select
    Selection.Replace What:="#DIV/0!", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Columns("B:B").Select
    Application.CutCopyMode = False
    Selection.NumberFormat = "#,##0"
    Columns("A:AJ").Select
    Selection.AutoFilter
    Selection.Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
    
    'This creates an Order Rat worksheet and copies over Order Rat source data
    Dim templateName As String
    templateName = ActiveWindow.Caption
    Windows(templateName).Activate
    'Sheets.Add.Name = "Order Rat"
    'Worksheets("Order Rat").Move After:=Sheets("FinishedGoods")
    Windows("Order Rat.csv").Activate
    Cells.Select
    Selection.Copy
    Windows(templateName).Activate
    Sheets("Order Rat").Select
    Cells.Select
    Range("D1").Activate
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
        
    ' this closes the order rat.csv file
    Windows("Order Rat.csv").Activate
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True

End Sub

Sub Z_ORDERRAT_FORMULAS()
'This macro copies over formulas from the formula sheet to the new FinishedGoods sheet.
'Coded by N99610 on 6/1/15

    Sheets("Formulas").Select
    Range("CC1:CR2").Select
    Selection.Copy
    Sheets("FinishedGoods").Select
    Range("CC1").Select
    ActiveSheet.Paste

    Dim lastRowFO As Long
    lastRowFO = Range("A" & Rows.Count).End(xlUp).Row
    Range("CC2:CR2").AutoFill Destination:=Range("CC2:CR" & lastRowFO)
    Range("A1").Select
    
    Application.CutCopyMode = False
    Range("CJ2").Select
    Selection.Copy
    Range("CJ2:CJ" & lastRowFO).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

End Sub

Sub Z_ORDERRAT_FORMAT()
' This macro formats the finished goods sheet for viewing.
'Coded by N99610 on 6/1/15

    Sheets("FinishedGoods").Select
    Columns("CC:CR").Select
    Columns("CC:CR").EntireColumn.AutoFit
    Cells.Select
    Range("CB1").Activate
    Cells.EntireColumn.AutoFit
    Range("CE1").Select
    Columns("CC:CC").ColumnWidth = 9.86
    Columns("CD:CD").ColumnWidth = 10.86
    Columns("CE:CE").ColumnWidth = 8.29
    Columns("CC:CC").EntireColumn.AutoFit
    Columns("CD:CD").EntireColumn.AutoFit
    Columns("CE:CE").EntireColumn.AutoFit
    Columns("CF:CF").EntireColumn.AutoFit
    Columns("CG:CG").EntireColumn.AutoFit
    Columns("CH:CH").EntireColumn.AutoFit
    Columns("CI:CI").EntireColumn.AutoFit
    Columns("CJ:CJ").EntireColumn.AutoFit
    Columns("CK:CK").ColumnWidth = 11.71
    Columns("CL:CL").ColumnWidth = 11.43
    Columns("CM:CM").ColumnWidth = 8.29
    Columns("CN:CN").ColumnWidth = 13
    Columns("CN:CN").ColumnWidth = 19.71
    Columns("CO:CO").ColumnWidth = 12.43
    Columns("CP:CP").ColumnWidth = 8.14
    Columns("CQ:CQ").ColumnWidth = 10.71
    Columns("CR:CR").ColumnWidth = 10.43
    Columns("CF:CF").ColumnWidth = 9.71
    Columns("CG:CG").ColumnWidth = 7.57
    Cells.Select
    Range("CB1").Activate
    Cells.EntireRow.AutoFit
    Range("CH7").Select
    Range("CR1").Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.AutoFilter
    Selection.AutoFilter
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("CC6").Select
    Columns("CC:CC").ColumnWidth = 10.14
    Columns("CC:CC").ColumnWidth = 12.29
    Columns("CC:CC").ColumnWidth = 15.14
    Columns("CC:CC").ColumnWidth = 16.29
    Cells.Select
    Range("CB1").Activate
    Cells.EntireRow.AutoFit
    Rows("1:1").Select
    Range("CB1").Activate
    With Selection
        .HorizontalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("CC:CC").EntireColumn.AutoFit
    Range("A1").Select
End Sub
