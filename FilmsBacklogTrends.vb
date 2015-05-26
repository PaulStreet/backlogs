Sub BACKLOG3()
' Based off of Bags BACKLOG4

'this section pastes special all values on the all plants tab
    Sheets("AllOrders").Select
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    'this section copies the AllOrders tab to a csv file for OMITTED
    Sheets("AllOrders").Select
    Sheets("AllOrders").Copy

    Rows("1:1").Select
    Range("AQ1").Activate
    Selection.Delete Shift:=xlUp
    ActiveWorkbook.SaveAs Filename:= _
        "PATH_OMITTED", _
        FileFormat:=xlCSV, CreateBackup:=False
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    
    ' this section updates the pivot tables
    '
    Sheets("By Week").Select
    Range("C37").Select
    ActiveSheet.PivotTables("PivotTable1").PivotCache.Refresh
            
    ' This Section captures today's backlog numbers as trend data for future comparisons
    ' Requested by TIM ROGERS on 5/13/15
    ' Coded by N99610 on 5/26/15
    Sheets("Total BL Trend").Select
    Rows("1:1").Select
    Selection.Copy
    Columns("A").Find("", Cells(Rows.Count, "A")).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=TODAY()"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    Dim lastRowFB As Long
    Dim nextLastRowFB As Long
    lastRowFB = Range("D" & Rows.Count).End(xlUp).Row
    nextLastRowFB = lastRowFB - 1
    Range("E" & nextLastRowFB).AutoFill Destination:=Range("E" & nextLastRowFB & ":E" & lastRowFB)
    Range("F" & nextLastRowFB).AutoFill Destination:=Range("F" & nextLastRowFB & ":F" & lastRowFB)
    Application.CutCopyMode = False
    Sheets("Daily Summary").Select
    Range("A1").Select
            
    ' This Section captures today's late MIMI orders as trend data for future comparisons
    ' Requested by HARVEYJ on 5/8/15
    ' Coded by N99610 on 5/13/15
    Sheets("MIMI Late Trends").Select
    Rows("1:1").Select
    Selection.Copy
    Columns("A").Find("", Cells(Rows.Count, "A")).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=TODAY()"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Sheets("Daily Summary").Select
    Range("A1").Select
            
End Sub
