Sub AllStockCalculationsforAllSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If WorksheetFunction.CountA(ws.UsedRange) > 1 Then   ' Check if the worksheet has data
            Call SumandLargestTradeVolsByGroup(ws)  ' Call all necessary subroutines for the current worksheet
            Call CalcOpenCloseDiffsByGroup(ws)
        End If
    Next ws 'move to next worksheet
End Sub
Sub SumandLargestTradeVolsByGroup(ws As Worksheet) 'subroutine to calculate the total trading volume for each stock over the course of a year and the largest trading volume for all stocks in a year
    
    Dim lastRow As Long
    Dim currentStockSymbol As String
    Dim sumTradingVol As Double
    Dim outputRow As Long
    Dim greatestVolSymbol As String
    Dim greatestVol As Double
   
    
    'set the global column and row headers for data output to the appropriate titles for subroutine
    ws.Range("I1").Value = "Ticker"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    
    ' initialize global variables
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row 'Sets upper bound of our loop to the number of the last row in column A in the worksheet that contains data (starting from the bottom)
    currentStockSymbol = ws.Cells(2, "A").Value ' Assuming the stock symbol data starts from row 2 - read in that stock symbol string into a variable
    sumTradingVol = 0 'Set the inital trading volume amount to zero to begin summing
    outputRow = 2 'output the calculated values starting in row 2
    greatestVol = 0
    greatestVolSymbol = ""
    
    ' Loop through all rows in the worksheet that contain data to calculate trading volume totals
    
    For i = 2 To lastRow
            
            If ws.Cells(i, "A").Value = currentStockSymbol Then     ' Add the value in column G to the trading volume sum if the stock symbols match
                sumTradingVol = sumTradingVol + ws.Cells(i, "G").Value
                
            Else          ' If the current stock symbol does NOT match the symbol in column G, output the trading volume sum for the current stock symbol next to the corresponding stock symbol string, compare its total to the stored greatest total see if its the new greatest, and then reset the sum variable to day one value for new stock symbol; increment output row
                
                If sumTradingVol > greatestVol Then  'compare the trade volume sum from that particular stock symbol to if it is greater than the last stored greatest trade vol
                      greatestVol = sumTradingVol
                      greatestVolSymbol = currentStockSymbol
                End If
                    
                ws.Cells(outputRow, "L").Value = sumTradingVol
                ws.Cells(outputRow, "I").Value = currentStockSymbol
                currentStockSymbol = ws.Cells(i, "A").Value
                sumTradingVol = ws.Cells(i, "G").Value
                outputRow = outputRow + 1
            
            End If
    
    Next i
    
    ' Output the trading volume sum for the last stock symbol in the list and output the greatest trading volume total and its corresponding symbol
    
    ws.Cells(outputRow, "L").Value = sumTradingVol
    ws.Cells(outputRow, "I").Value = currentStockSymbol
    ws.Range("P4").Value = greatestVolSymbol
    ws.Range("Q4").Value = greatestVol
   
End Sub
Sub CalcOpenCloseDiffsByGroup(ws As Worksheet) 'subroutine to calculate total open to close price difference for the year and % change for the year

    Dim lastRow As Long
    Dim dateString As String
    Dim yearPart As String
    Dim monthPart As String
    Dim dayPart As String
    Dim specificDayOpen As String
    Dim specificMonthOpen As String
    Dim specificMonthClose As String
    Dim specificDayClose As String
    Dim openPriceYear As Double
    Dim closePriceYear As Double
    Dim percentChangeYear As Double
    Dim greatestPercentInc As Double
    Dim greatestPercentDec As Double
    Dim greatestPercentIncSymbol As String
    Dim greatestPercentDecSymbol As String
    Dim finalYearDiff As Double
   
   'set global output row and column headings to appropriate title for the subroutine
   
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    
     ' initialize global variables
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row 'Sets upper bound of our loop to the number of the last row in column A in the worksheet that contains data (starting from the bottom)
    currentStockSymbol = ws.Cells(2, "A").Value ' Assuming the stock symbol data starts from row 2 - read in that stock symbol string into a variable
    outputRow = 2 'output the calculated values starting in row 2
    greatestPercentIncSymbol = ""
    greatestPercentDecSymbol = ""
    greatestPercentInc = 0
    greatestPercentDec = 0
    
    'set strings for the dates we need to match to find opening and closing prices
    
    specificMonthOpen = "01" 'opening month value for string matching
    specificDayOpen = "02" 'opening day value for string matching
    specificMonthClose = "12" 'closing month value for string matching
    specificDayClose = "31" 'closing month value for string matching
    
    'remove any old formatting from all rows that contain data in columns J and K
    
    columnLetter = "J"
    Set ColumnRange = ws.Range(columnLetter & "1:" & columnLetter & ws.Cells(ws.Rows.Count, columnLetter).End(xlUp).Row)
    ColumnRange.ClearFormats
     
    columnLetter = "K"
    Set ColumnRange = ws.Range(columnLetter & "1:" & columnLetter & ws.Cells(ws.Rows.Count, columnLetter).End(xlUp).Row)
    ColumnRange.ClearFormats
    
    For i = 2 To lastRow  'start loop to calculate yearly open to close price differences
  
         If ws.Cells(i, "A").Value = currentStockSymbol Then 'make sure we haven't moved to a new stock symbol to start a new price change calculation
 
            dateString = ws.Cells(i, "B").Value  'parse the dates into strings so we can string match to find the opening and closing prices
            yearPart = Left(dateString, 4)
            monthPart = Mid(dateString, 5, 2)
            dayPart = Right(dateString, 2)
                   
                 If dayPart = specificDayOpen And monthPart = specificMonthOpen Then  ' Check that the date cell contains the specific day and month for each stock symbol (in this case the opening day and closing day) and set open and close day trade price variables to cell values
                       openPriceYear = ws.Cells(i, "C").Value
                 ElseIf dayPart = specificDayClose And monthPart = specificMonthClose Then
                       closePriceYear = ws.Cells(i, "F").Value
                 End If
         
         Else
         
            finalYearDiff = closePriceYear - openPriceYear ' Output the yearly open to close difference for the current stock symbol to a variable and then output that variable next to the corresponding stock symbol in the output row
            ws.Cells(outputRow, "J").Value = finalYearDiff
    
      If ws.Cells(outputRow, "J").Value < 0 Then   'conditional format the change value cell to red if negative or green if positive (or zero)
         ws.Cells(outputRow, "J").Interior.ColorIndex = 3
      ElseIf ws.Cells(outputRow, "J").Value >= 0 Then
         ws.Cells(outputRow, "J").Interior.ColorIndex = 4
      End If
     
      If openPriceYear <> 0 Then 'check for non-zero starting price to circumvent divide by zero errors during % change calculation
         percentChangeYear = ((closePriceYear - openPriceYear) / openPriceYear) * 100
      Else
         percentChangeYear = 0
      End If
      
      ws.Cells(outputRow, "K").Value = percentChangeYear & "%" 'output the %change value with % symbol and sign into the appropriate row in the output section
      
         If ws.Cells(outputRow, "K").Value < 0 Then
         ws.Cells(outputRow, "K").Interior.ColorIndex = 3
         ElseIf ws.Cells(outputRow, "K").Value >= 0 Then
         ws.Cells(outputRow, "K").Interior.ColorIndex = 4
         End If
 
      If percentChangeYear > greatestPercentInc Then  'compare the trade volume sum from that particular stock symbol to if it is greater than the last stored greatest trade vol
          greatestPercentInc = percentChangeYear
          greatestPercentIncSymbol = currentStockSymbol
      ElseIf percentChangeYear < greatestPercentDec Then
          greatestPercentDec = percentChangeYear
          greatestPercentDecSymbol = currentStockSymbol
      End If
   
    'reset the opening price variable and output row to correspond to new stock symbol
  
          currentStockSymbol = ws.Cells(i, "A").Value
          openPriceYear = ws.Cells(i, "C").Value
          outputRow = outputRow + 1
         
   End If

Next i

       finalYearDiff = closePriceYear - openPriceYear
       ws.Cells(outputRow, "J").Value = finalYearDiff  'output the last stock symbol in the list's yearly open to close price difference to the final difference variable; output that variable in the last output row

      If ws.Cells(outputRow, "J").Value < 0 Then   'conditional format the last stock symbol in the list's yearly open to close difference output from the list
          ws.Cells(outputRow, "J").Interior.ColorIndex = 3
      ElseIf ws.Cells(outputRow, "J").Value >= 0 Then
          ws.Cells(outputRow, "J").Interior.ColorIndex = 4
      End If
      
      If openPriceYear <> 0 Then
         percentChangeYear = ((closePriceYear - openPriceYear) / openPriceYear) * 100 ' calculate and output the percent change for the last stock symbol in the list
      Else
         percentChangeYear = 0
      End If
       
       ws.Cells(outputRow, "K").Value = percentChangeYear & "%"
      
      If ws.Cells(outputRow, "K").Value < 0 Then                 'conditional format the last stock symbol in the list's yearly open to close percent difference output from the list
         ws.Cells(outputRow, "K").Interior.ColorIndex = 3
      ElseIf ws.Cells(outputRow, "K").Value >= 0 Then
         ws.Cells(outputRow, "K").Interior.ColorIndex = 4
      End If
      
      If percentChangeYear > greatestPercentInc Then  'check to see if last stock in list is the highest or lowest % change
            greatestPercentInc = percentChangeYear
            greatestPercentIncSymbol = currentStockSymbol
      ElseIf percentChangeYear < greatestPercentDec Then
            greatestPercentDec = percentChangeYear
            greatestPercentDecSymbol = currentStockSymbol
      End If
 
 ws.Range("P2").Value = greatestPercentIncSymbol   'output the greatest % increase and decrease for the year
 ws.Range("P3").Value = greatestPercentDecSymbol
 ws.Range("Q2").Value = greatestPercentInc & "%"
 ws.Range("Q3").Value = greatestPercentDec & "%"
 
 columnLetter = "J" ' Format the yearly difference column to have two decimal places even when a trailing zero is present; do this for every row in the column that contains data
 NumberFormat = "0.00"
 Set ColumnRange = ws.Range(columnLetter & "1:" & columnLetter & ws.Cells(ws.Rows.Count, columnLetter).End(xlUp).Row)
 ColumnRange.NumberFormat = NumberFormat

End Sub