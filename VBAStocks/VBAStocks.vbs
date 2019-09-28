Attribute VB_Name = "Module1"
' Create a script that will loop through all the stocks for one year for each run and take the following information.

  ' The ticker symbol.

  ' Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  ' The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  ' The total stock volume of the stock.

'You should also have conditional formatting that will highlight positive change in green and negative change in red.

' ---------------------------------------------------------------------------------------


Sub stockchallenge():
    
  For Each ws In Worksheets
    
    ' Set Dimensions
    Dim total As Double
    Dim ticker As String
    Dim opening As Double
    Dim closing As Double
    Dim Change As Double
    Dim percentChange As Double
    Dim j As Integer

    ' Start total out at zero
    total = 0
    
    ' get the row number of the last row with data
    RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    ' get the number of rows for each ticker
    tickerRowCount = 0

    ' Set title row for the summary table
    j = 2
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    For i = 2 To RowCount

        ' If ticker changes then print results
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            ' Stores results in variable
            ticker = ws.Cells(i, 1).Value
            closing = ws.Cells(i, 6).Value
            opening = ws.Cells(i - tickerRowCount, 3).Value
            Change = closing - opening
            
            If opening <> 0 Then
                percentChange = Change / opening
            Else
                percentChange = 0
            End If
            
            ' Add to the total volume
            total = total + ws.Cells(i, 7).Value
            
            ' Print ticker symbol in the respective column for the `ticker`
            ws.Range("I" & j).Value = ticker
            
            ' Print total in the respective column for the `total stock volume`
            ws.Range("L" & j).Value = total
            ws.Range("J" & j).Value = Change
            ws.Range("K" & j).Value = percentChange
            ws.Range("K" & j).NumberFormat = "0.00%"

            
            ' Conditional formating: highlight positive change in green and negative change in red.
            If ws.Range("J" & j).Value > 0 Then
                ws.Range("J" & j).Interior.ColorIndex = 4
            Else
                ws.Range("J" & j).Interior.ColorIndex = 3
            End If
                
            ' Reset Total
            total = 0

            ' Move to next row
            j = j + 1
            
            ' Reset ticker row count
            tickerRowCount = 0

        ' Else keep adding to the total volume
        Else
            total = total + ws.Cells(i, 7).Value
            tickerRowCount = tickerRowCount + 1

        End If

    Next i
    
    ' Get the number of rows in the summary table
    sumRowCount = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    ' Set dimensions to find the row with greatest increase, decrease, and total stock volume
    Dim maxIncreaseRow As Integer
    Dim maxDecreaseRow As Integer
    Dim maxTotalRow As Integer
    
    ' Set title row for the table with greatest values
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    

    ' Iterate through the summary table
    
    'Start iteration from row 2 in summary table
    maxIncreaseRow = 2
    maxDecreaseRow = 2
    maxTotalRow = 2
    
    For k = 2 To sumRowCount
        
        ' Compare the current maximum value with the value in a new row, update the maximum if the value in the new row is larger
        If ws.Cells(maxIncreaseRow, 11).Value < ws.Cells(k, 11).Value Then
            maxIncreaseRow = k
        End If
        
        If ws.Cells(maxDecreaseRow, 11).Value > ws.Cells(k, 11).Value Then
            maxDecreaseRow = k
        End If
        
        If ws.Cells(maxTotalRow, 12).Value < ws.Cells(k, 12).Value Then
            maxTotalRow = k
        End If
        
    Next k
    
    ' Store the maximum values
    maxIncreaseVal = ws.Cells(maxIncreaseRow, 11).Value
    maxDecreaseVal = ws.Cells(maxDecreaseRow, 11).Value
    maxTotalVal = ws.Cells(maxTotalRow, 12).Value
    
    ' Format and input values to the maximum value table
    ws.Range("P2").Value = ws.Range("I" & maxIncreaseRow).Value
    ws.Range("P3").Value = ws.Range("I" & maxDecreaseRow).Value
    ws.Range("P4").Value = ws.Range("I" & maxTotalRow).Value
    ws.Range("Q2").Value = maxIncreaseVal
    ws.Range("Q3").Value = maxDecreaseVal
    ws.Range("Q4").Value = maxTotalVal
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"
    
    ws.Columns("A:Q").AutoFit
    
Next ws

End Sub

