Sub Z_TRENDS()
' This macro copies down the the Lates Trends for today to the historical table.
' Requested by Shawn Boyer on 6/23
' Coded by N99610 on 6/23

    Sheets("Trend ").Select
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
    
    Dim lastRowTrends As Long
    Dim nextLastTrends As Long
    lastRowTrends = Range("D" & Rows.Count).End(xlUp).Row
    nextLastTrends = lastRowTrends - 1
    Range("K" & nextLastTrends).AutoFill Destination:=Range("K" & nextLastTrends & ":K" & lastRowTrends)
    Range("L" & nextLastTrends).AutoFill Destination:=Range("L" & nextLastTrends & ":L" & lastRowTrends)
    Application.CutCopyMode = False

End Sub
