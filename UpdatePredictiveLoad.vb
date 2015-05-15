Sub UpdatePredictiveLoad()

    'This section adds an estimated weight and estimated feet column to the AllPlants tab.
    Sheets("AllPlants").Select
    Sheets("AllPlants").Range("CP3:CP20000").ClearContents
    Sheets("AllPlants").Range("CQ3:CP20000").ClearContents
    Dim lastRowAP As Long
    lastRowAP = Range("CO" & Rows.Count).End(xlUp).Row
    Range("CP2").AutoFill Destination:=Range("CP2:CP" & lastRowAP)
    Range("CQ2").AutoFill Destination:=Range("CQ2:CQ" & lastRowAP)
    Range("CP2").Select
    
    'This section adds a bag type to the press section by conducting a lookup in the Press tab.
    Sheets("Press").Select
    Sheets("Press").Range("AZ3:AZ20000").ClearContents
    Dim lastRowPress As Long
    lastRowPress = Range("AY" & Rows.Count).End(xlUp).Row
    Range("AZ2").AutoFill Destination:=Range("AZ2:AZ" & lastRowPress)
    Range("AZ2").Select
    
    'This section goes to each worksheet and refreshes the pivot tables.
    'It also sets the top left cell as the active cell as it cycles through.
    Dim PT As PivotTable
    Dim WS As Worksheet
    For Each WS In ThisWorkbook.Worksheets
        For Each PT In WS.PivotTables
          PT.RefreshTable
          Range("A1").Select
        Next PT
    Next WS
    
    'This section simply selects the CPM ES tab as the current worksheet at the end.
    Sheets("CPM ES").Select
    
End Sub
