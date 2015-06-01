Sub Z_ORDERRAT_1()
'This section has undergone significant code revisions by N99610 on 6/1/2015
'Changes
'	Insertion of lookup keys and Days OH information has been drastically streamlined
'	Absolute values for number of rows have been replaced by a variable count of rows
'	The function will now automatically import the CSIM Order Rat file without human interference
'	The file will automatically avoid saving over the CSIM Order Rat CSV and will now dump clipboard at end of function
'	The aged tubing section seems non functional and code has been deprecated

    Workbooks.OpenText Filename:="PATH_OMITTED"

    'New Code by N99610 on 6/1/2015
    'This section counts the number of rows and stores it as a variable.
    Dim lastRowOR As Long
    lastRowOR = Range("A" & Rows.Count).End(xlUp).Row
    'End New Code
    
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
	'New code by N99610 on 6/1/2015
    Range("A2:B2").AutoFill Destination:=Range("A2:B" & lastRowOR)
    'End New code
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
    
    'this section copies the order rat.xls to the order rat template
    Windows("Order Rat.csv").Activate
    Cells.Select
    Selection.Copy
    Windows("Backlogtemplate-NEW3.xls").Activate
    Sheets("Order Rat").Select
    Cells.Select
    Range("D1").Activate
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
	Application.CutCopyMode = False
        
    ' this closes the order rat.xls file
    Windows("Order Rat.csv").Activate
	'New code by N99610 on 6/1/2015
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True

'   This section formats the Aged tab so the lookups will work.
'	As of 6/1/2015 this section was deprecated by N99610
    'Windows("Backlogtemplate-NEW3.xls").Activate
    'Sheets("AgedBags").Select
    'Rows("1:5").Select
    'Range("F1").Activate
    'Selection.Delete Shift:=xlUp
    'Cells.Select
    'Range("F1").Activate
    'Selection.Interior.ColorIndex = xlNone
    'Columns("i:i").Select
    'Selection.TextToColumns Destination:=Range("i1"), DataType:=xlDelimited, _
        'TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        'Semicolon:=False, Comma:=False, Space:=False, Other:=False, OtherChar _
        ':="/", FieldInfo:=Array(1, 2), TrailingMinusNumbers:=True
    'Cells.Select
    'Range("F1").Activate
    'Selection.Sort Key1:=Range("i2"), Order1:=xlAscending, Header:=xlGuess, _
        'OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        'DataOption1:=xlSortNormal    

End Sub
