Sub BACKLOG2_AUTOLATES()
'This macro automates much of the late order review.
'Automates the input of pre-review late numbers.
'Automates the review of backorders and even changes dates.
'Orders still need to be reviewed for duplicates, development, etc.
'Recommend saving before running . . .
'Coded by N99610 on 6/28/15

	'This section of code automates the entering of lates before review.
    Sheets("LATE ORDERS").Select
    
    Range("U6").Select
    ActiveSheet.PivotTables("PivotTable5").PivotCache.Refresh
    Range("S1").Select
    Selection.Copy
    
    Sheets("Lates Trend").Select
    Dim lastRowTrends As Long
    Dim nextRowTrends As Long
    lastRowTrends = Range("A" & Rows.Count).End(xlUp).Row
    nextRowTrends = lastRowTrends + 1
    
    Range("BY" & nextRowTrends).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Sheets("LATE ORDERS").Select
    Range("S3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Lates Trend").Select
    Range("CA" & nextRowTrends).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False


	'This section of code makes a bunch of helper columns to look at original order information
	'calculate a percentage that is still in MIMI and suggest that an order is a backorder if
	'less than 50% of the original order is in MIMI and the order is not one of many production orders
	'and if the order is not a SFG.
    Sheets("AllOrders").Select
    Columns("BD:BH").Select
    Selection.EntireColumn.Hidden = False

	Dim lastRowAO As Long
    lastRowAO = Range("A" & Rows.Count).End(xlUp).Row

    Range("CC1").Select
    ActiveCell.FormulaR1C1 = "OrigOrd Qty"
    Range("CC2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-22],SalesOrder!C3:C13,10,)"
    Range("CC2").Select
    Range("CC2").AutoFill Destination:=Range("CC2:CC" & lastRowAO)
    
    Range("CD1").Select
    ActiveCell.FormulaR1C1 = "OrigOrd UOM"
    Range("CD2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-23],SalesOrder!C3:C13,11,)"
    Range("CD2").Select    
    Range("CD2").AutoFill Destination:=Range("CD2:CD" & lastRowAO)

    Range("CE1").Select
    ActiveCell.FormulaR1C1 = "OrigOrd BOM"
	Range("CE2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-1]=0,0,RC81/(VLOOKUP(RC4&""/""&RC82&""/""&RC7&""/""&RC1,UOM!C1:C2,2,FALSE)))"
	Range("CE2").Select    
    Range("CE2").AutoFill Destination:=Range("CE2:CE" & lastRowAO)
    
	Range("CF1").Select
    ActiveCell.FormulaR1C1 = "Percent Left"
	Range("CF2").Select
    ActiveCell.FormulaR1C1 = "=IF(LEFT(RC[-80],1)=""2"",""SFG"",RC[-78]/RC[-1])"
	Range("CF2").Select    
    Range("CF2").AutoFill Destination:=Range("CF2:CF" & lastRowAO)
    
	Range("CG1").Select
    ActiveCell.FormulaR1C1 = "Possible Backorder"
    Range("CG2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]<0.5,""Y"",""N"")"
	Range("CG2").Select    
    Range("CG2").AutoFill Destination:=Range("CG2:CG" & lastRowAO)
    
	Range("CH1").Select
    ActiveCell.FormulaR1C1 = "# of PrdOrds Per SOLI"
    Range("CH2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-2]=""SFG"",RC[-2],COUNTIF(C[-27],RC[-27]))"
	Range("CH2").Select    
    Range("CH2").AutoFill Destination:=Range("CH2:CH" & lastRowAO)
    
	Range("CI1").Select
    ActiveCell.FormulaR1C1 = "New Date"
    Range("CI2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[-2]=""Y"",RC[-1]=1,RC[-32]=""Late""),RC[-34]+14,RC[-34])"
	Range("CI2").Select    
    Range("CI2").AutoFill Destination:=Range("CI2:CI" & lastRowAO)
    
	Range("CJ1").Select
    ActiveCell.FormulaR1C1 = "Changed?"
	Range("CJ2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]=RC[-35],"""",""Changed"")"
	Range("CJ2").Select    
    Range("CJ2").AutoFill Destination:=Range("CJ2:CJ" & lastRowAO)
    
    Columns("CI:CI").Select
    Selection.NumberFormat = "mm/dd/yy;@"
    Range("CC1:CJ1").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
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
        .WrapText = False
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
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("CC1").Select	
	
    Columns("BE:BG").Select
    Selection.EntireColumn.Hidden = True
	
	'This code copies over the dates and deletes the helper columns.  Comment out if you want to look at the helper columns.
    Columns("CI:CI").Select
    Selection.Copy
    Range("BA1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("BA1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "ADJ"
    Range("BA2").Select

    Columns("CC:CJ").Select
    Selection.Delete Shift:=xlToLeft
	
End Sub
