Sub BACKLOG5()

'This section creates separate XLS files of the AllPlants and Press Tabs.
'The purpose of this code is to generate source files for the Predictive Load and Backlog and the Aged Inventory Eliminator.
'Coded by N99610 on 5/27/15
    'This section copies the AllPlants tab and saves it the SVL Share Drive.
    Application.DisplayAlerts = False
    Dim tempWB As Workbook
    Sheets("AllPlants").Copy
    Set tempWB = ActiveWorkbook
    tempWB.SaveAs Filename:= _
        "PATH OMITTED", _
        FileFormat:=56, CreateBackup:=False
    tempWB.Save
    tempWB.Close
    'This section copies the Press tab and saves it the SVL Share Drive.
    Sheets("Press").Copy
    Set tempWB = ActiveWorkbook
    tempWB.SaveAs Filename:= _
        "PATH OMITTED", _
        FileFormat:=56, CreateBackup:=False
    tempWB.Save
    tempWB.Close
    Application.DisplayAlerts = True
'End Code Added on 5/27/15

End Sub
