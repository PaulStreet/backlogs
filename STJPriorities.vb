Sub STJPriorities()
' STJ Extrusion Priorities for 2015 Busy Season.

    'Selects the AllOrders worksheet
    Sheets("AllOrders").Select
    Range("A1").Select
    
    'This section grabs the last row from the AllOrders worksheet
    Dim lastRowAO As Long
    lastRowAO = Range("A" & Rows.Count).End(xlUp).Row

    'Filters for STJ orders on Extrusion Lines 1 thru 3
    ActiveSheet.Range("$A$1:$ES$" & lastRowAO).AutoFilter Field:=1, Criteria1:="1111"
    ActiveSheet.Range("$A$1:$ES$" & lastRowAO).AutoFilter Field:=2, Criteria1:=Array( _
        "EXLA01", "EXLA02", "EXLA03"), Operator:=xlFilterValues
    Cells.Select
    Selection.Copy
    Sheets("AllOrders").Select
    Sheets.Add
    ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    'Formats and renames the STJPriorities Worksheet
    ActiveSheet.Name = "STJPriorities"
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Application.CutCopyMode = False
    Selection.AutoFilter
    Cells.Select
    Cells.EntireColumn.AutoFit
    Rows("1:1").Select
    With Selection
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A1").Select
    
    'This section creates separate XLS files of the AllOrders
    Application.DisplayAlerts = False
    Dim tempWB As Workbook
    Sheets("STJPriorities").Copy
    Set tempWB = ActiveWorkbook
    tempWB.SaveAs Filename:= _
        "\\DUNFS01\Duncan\Logistics and Operations\Central Planning's Reporting Hub\Rollstock\Rollstock Backlog\STJPrioritiesCurrent.xls", _
        FileFormat:=56, CreateBackup:=False
    tempWB.Save
    tempWB.Close
    Sheets("STJPriorities").Delete
    Application.DisplayAlerts = True
    
    'This section removes all filters on the AllOrders worksheet
    Sheets("AllOrders").Select
    Range("A1").Select
    ActiveSheet.ShowAllData
    
End Sub
