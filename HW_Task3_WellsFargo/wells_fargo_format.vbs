Attribute VB_Name = "Module1"
' Steps:
' 1. Loop through every worksheet and select the state contents.
' 2. Copy the state contents and paste it into the Combined_Data tab

Sub WellsFargo_HWTask3()
    
    ' Add a sheet named "Combined Data"
    Sheets.Add.Name = "Combined_Data"
    
    ' Move created sheet to be first sheet
    Sheets("Combined_Data").Move Before:=Sheets(1)
    
    ' Specify the location of the combined sheet
    Set combined_sheet = Worksheets("Combined_Data")
    

    ' Loop through all sheets
    For Each ws In Worksheets
    
        ' Grabbed the WorksheetName
        WorksheetName = ws.Name
        ' MsgBox WorksheetName

        ' Find the last row of the combined sheet after each paste
        ' Add 1 to get first empty row
        lastRow = combined_sheet.Cells(Rows.Count, "A").End(xlUp).Row + 1

        ' Find the last row of each worksheet
        ' Subtract one to return the number of rows without header
        ' HINT: Use similar logic as above, but this time subtract a row
        ' ****************
        lastRowState = ws.Cells(Rows.Count, "A").End(xlUp).Row - 1
        ' ****************
        
        ' Copy the contents of each state sheet into the combined sheet
        ' ****************
        ' Copy the contents of each state sheet into the combined sheet
        combined_sheet.Range("A" & lastRow & ":G" & ((lastRowState - 1) + lastRow)).Value = ws.Range("A2:G" & (lastRowState + 1)).Value
        ' ****************
        
    Next ws

    ' Copy the headers from sheet 1
    combined_sheet.Range("A1:G1").Value = Sheets(2).Range("A1:G1").Value
    
    ' Autofit to display data
    combined_sheet.Columns("A:G").AutoFit
    
End Sub
