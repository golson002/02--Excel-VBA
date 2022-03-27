Attribute VB_Name = "Module1"
Public Sub Stocks()

For Each ws In Worksheets

    'Define worksheet name
    Dim Worksheetname As String
    Worksheetname = ws.Name
    
    'Name the headers for the output table and autofit the columns to the correct length
    ws.Range("J1") = "Ticker"
    ws.Range("K1") = "Yearly Change"
    ws.Range("L1") = "Percent Change"
    ws.Range("M1") = "Total Stock Volume"
    ws.Range("J:M").EntireColumn.AutoFit
    
   'Define a start location in the output table for the first stock's ticker symbol, yearly change, percent change, and total volume. We want to start in row 2 since our headers are in row 1
    Dim SummaryTableRow As Integer
    SummaryTableRow = 2
    
    'Find the last row of raw data in the sheet so we do not have to manually count no. of rows and our code works for each sheet
    Dim lrow As Long
    lrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Define a starting row for each loop- necessary for calculating yearly change, percent change, total volume
    Dim startrow As Long
    startrow = 2
    
    'Define a variable i (and specifiy its data type) that our code will run through until we reach the last row (i reaches lrow)
    Dim i As Long
    
    'Define variable names and data types for the ticker symbol, yearly change, percent change, and total volume of each stock
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As LongLong
    
    'Loop through all rows on the sheet, starting at row 2 and ending at lrow, and find the unique ticker symbols and calculate the yearly change, percent change, and total volume for those stocks
    For i = 2 To lrow
    
        'If the value in the cell directly below the current cell selection is not the same as the value in the current selection, then
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'Write the value in the current row (i), column 1 (the ticker symbol) in the output table
            Ticker = ws.Cells(i, 1).Value
            ws.Range("J" & SummaryTableRow).Value = Ticker
            
            'Calculate the yearly change as the difference between the value in the current row (i), column 6 and the value in the starting row for the loop, column 3. This is the stock closing price at the end of the year - the opening price at the beginning of the year
            YearlyChange = ws.Cells(i, 6).Value - ws.Cells(startrow, 3).Value
            ws.Range("K" & SummaryTableRow).Value = YearlyChange
             
            'Calculate the percent change in the year for the stock as the difference between the value in the current row (i), column 6 and the value in the starting row for the loop, column 3, divided by the value in the starting row for the loop, column 3. (close price - open price)/open price
            PercentChange = (ws.Cells(i, 6).Value - ws.Cells(startrow, 3).Value) / (ws.Cells(startrow, 3).Value)
            ws.Range("L" & SummaryTableRow).Value = Format(PercentChange, "Percent")
            
            'Calculate the total volume in the year for the stock as the sum of the range between the value in the current row (i), column 7 and the value in the starting row for the loop, column 7
            TotalVolume = Application.WorksheetFunction.Sum(Range(ws.Cells(i, 7), ws.Cells(startrow, 7)))
            ws.Range("M" & SummaryTableRow).Value = TotalVolume
            
                'Use conditional formatting to color the cell green if the yearly change was positve and red if the yearly change was negative
                If ws.Range("K" & SummaryTableRow).Value >= 0 Then
                    ws.Range("K" & SummaryTableRow).Interior.ColorIndex = 4
                    Else
                    ws.Range("K" & SummaryTableRow).Interior.ColorIndex = 3
                End If
                
            'Add one to the row we left off at in the raw data to reset our starting row for the next loop and capture the next stock's ticker symbol
            startrow = i + 1
            
            'Add a row to the output table so we do not overwrite the data we just added
            SummaryTableRow = SummaryTableRow + 1
            
        End If
    
    Next i
    
    'Format column K so all Yearly Change numbers are defined to the same decimal point
     ws.Columns("K:K").NumberFormat = "0.00"
    
    'Name the headers for the greatest change (+/-) and greatest total volume table
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    
    'Find the last row in the output table we created so we do not have to manually count no. of rows and our code works for each sheet
    Dim lrow2 As Long
    lrow2 = Cells(Rows.Count, 10).End(xlUp).Row
    
    'Define a variable j (and specifiy its data type) that our code will run through until we reach the last row (j reaches lrow2)
    Dim j As Long
    
    'Define variable names and data types for the greatest percent increase, greatest percent decrease, and greatest total volume values from the output table
    Dim Greatest_PercentIncrease As Double
    Dim Greatest_PercentDecrease As Double
    Dim Greatest_TotalVolume As LongLong
    
    'Assign a value to each of our defined variables so excel has a value to start comparing to when it is looping through the rows in the output table
    Greatest_PercentIncrease = ws.Cells(2, 12).Value
    Greatest_PercentDecrease = ws.Cells(2, 12).Value
    Greatest_TotalVolume = ws.Cells(2, 13).Value

    'Loop through all rows on the sheet, starting at row 2 and ending at lrow2, and find the stock with the greatest percent increase. (lrow2 is different from lrow, which is how excel knows to only look through the output table and not the entire sheet)
    For j = 2 To lrow2
    
        'Find the stock with the greatest percent increase by comparing the value in the current row (j), column 12 to the value we set for Greatest_PercentIncrease
        If ws.Cells(j, 12).Value > Greatest_PercentIncrease Then
            'If the value in the current cell is greater than the Greatest_PercentIncrease value, then reset the Greatest_PercentIncrease value to equal the value in the current cell
            Greatest_PercentIncrease = ws.Cells(j, 12).Value
            ws.Cells(2, 16).Value = ws.Cells(j, 10).Value
            ws.Cells(2, 17).Value = Format(ws.Cells(j, 12).Value, "Percent")
            
            'Otherwise, Greatest_PerecentIncrease still equals the value in cell(2,12).
            Else
            Greatest_PercentIncrease = Greatest_PercentIncrease
        
        End If
            
    Next j
    
     'Loop through all rows on the sheet, starting at row 2 and ending at lrow2, and find the stock with the greatest percent decrease. (lrow2 is different from lrow, which is how excel knows to only look through the output table and not the entire sheet)
    For j = 2 To lrow2
    
        'Find the stock with the greatest percent decrease by comparing the value in the current row (j), column 12 to the value we set for Greatest_PercentDecrease
       If ws.Cells(j, 12).Value < Greatest_PercentDecrease Then
            'If the value in the current cell is less than the Greatest_PercentDecrease value, then reset the Greatest_PercentDecrease value to equal the value in the current cell
            Greatest_PercentDecrease = ws.Cells(j, 12).Value
            ws.Cells(3, 16).Value = ws.Cells(j, 10).Value
            ws.Cells(3, 17).Value = Format(ws.Cells(j, 12).Value, "Percent")
            
            'Otherwise, Greatest_PerecentDecrease still equals the value in cell(2,12).
            Else
            Greatest_PercentDecrease = Greatest_PercentDecrease
            
        End If
        
     Next j
            
    'Loop through all rows on the sheet, starting at row 2 and ending at lrow2, and find the stock with the greatest total volume. (lrow2 is different from lrow, which is how excel knows to only look through the output table and not the entire sheet)
    For j = 2 To lrow2
    
        'Find the stock with the greatest total volume by comparing the value in the current row (j), column 13 to the value we set for Greatest_TotalVolume
        If ws.Cells(j, 13).Value > Greatest_TotalVolume Then
            'If the value in the current cell is greater than the Greatest_TotalVolume value, then reset the Greatest_TotalVolume value to equal the value in the current cell
            Greatest_TotalVolume = ws.Cells(j, 13).Value
            ws.Cells(4, 16).Value = ws.Cells(j, 10).Value
            ws.Cells(4, 17).Value = ws.Cells(j, 13).Value
                
            'Otherwise, Greatest_TotalVolume still equals the value in cell(2,13).
            Else
            Greatest_TotalVolume = Greatest_TotalVolume
        
        End If
    
    Next j
        
    'Autofit the column widths to fit the data entered from the conditionals and loops.
    ws.Range("O:Q").EntireColumn.AutoFit
    
'Move to the next worksheet with the following year's data
Next ws

End Sub

