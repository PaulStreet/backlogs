 Sub BACKLOG3()
 
    'Some information has been removed, only late trend information is available from this file.
 
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
