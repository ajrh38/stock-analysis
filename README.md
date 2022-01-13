# Stock-analysis by Arturo Rodriguez # 

## Project Overview ##
The following https://github.com/ajrh38/stock-analysis/blob/main/VBA_challenge.xlsm.xlsm provides an automated analysis of stock performance during 2017 and 2018. The overall analysis is achieved through Visual basic automation fetching data from Excel database and calculation key indicators that represent performance from the begning till the end of the year. The VBA automation has been refactored from the original file https://github.com/ajrh38/stock-analysis/blob/main/green_stocks.xlsm to improve the performance time during the calculation.

## Results ##

Results present "ENPH" Ticker(stock) and "Run" as the best available options for investment. Both show positive return and high daily volume of transactions making them atractive investments. The following image show a visual representation of the gain in 2018. image <img src="https://github.com/ajrh38/stock-analysis/blob/main/Ticker_Gains_2018.PNG">

## Summary ##

In terms of automation refactoring the code resulted in the code running aproximate 6 times faster than the original code. This is shown comparing both Timestamps from the run. original code is shown here
<img src="https://github.com/ajrh38/stock-analysis/blob/main/2017%20Before%20Refactoring.PNG"> versus the refactored outcome 
<img src="https://github.com/ajrh38/stock-analysis/blob/main/VBA_Challenge_2017.png">. 
This is the result of changes like creating a ticker index, and using arrays.

### Advantages and disadvantages ###
This tell us there are advantages to refactore codes like improving execution times, simplifying codes, positiblity to clean and add bettter comments, however it is important to understand the disavantages of refactoring an existing code, it can be dificult to understand if the commenting is bad, does not start from a clean slate and can import inneficient coding, therefore sometimes it is better to write the code from scratch.

Back up code utilized in this analysis
  '''
  Sub AllStocksAnalysisRefactored()
    
    
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    
 
    
    Tickerindex = 0
    
    

    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
     tickerVolumes(Tickerindex) = tickerVolumes(Tickerindex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
         If Cells(i, 1).Value = tickers(Tickerindex) And Cells(i - 1, 1).Value <> tickers(Tickerindex) Then
            tickerStartingPrices(Tickerindex) = Cells(i, 6).Value
        End If
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
           If Cells(i, 1).Value = tickers(Tickerindex) And Cells(i + 1, 1).Value <> tickers(Tickerindex) Then
            tickerEndingPrices(Tickerindex) = Cells(i, 6).Value
         End If
            

            '3d Increase the tickerIndex.
               If Cells(i, 1).Value = tickers(Tickerindex) And Cells(i + 1, 1).Value <> tickers(Tickerindex) Then
                Tickerindex = Tickerindex + 1
            End If
            
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub '''
