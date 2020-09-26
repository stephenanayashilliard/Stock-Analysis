# Stock-Analysis

## Project Overview
My client, Steve, needed help with his analysis of what stocks were worth investing in.  Although, our client is well versed in the use of Excel, it was determined that using VBA and automating the anaylsis process would better serve his purposes.

## Results
When I began to code for my client's project, I initially started out writing a simple if/then statement so that my client could run an analysis specifically analysing DAQO stocks based on year, their Total Daily Volume and the stocks return.
   
   - #### Original if/then statment for "DAQO" Analysis
    Sub DQAnalysis()

    Worksheets("DQ Analysis").Activate
    
    Range("A1").Value = "DAQO(Ticker:DQ)"

    'Create a header row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    Worksheets("2018").Activate

    'set intiial volume to zero
    totalvolume = 0

    Dim startingPrice As Double
    Dim endingPrice As Double

    'find the number of rows to loop over
    'rowend code taken from https://stackoverglow.com/questions/18088729/row-count
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    'loop over all the rows
    For i = 2 To RowCount

        If Cells(i, 1).Value = "DQ" Then
    
            'increase totalVolume by the value in the current row
            totalvolume = totalvolume + Cells(i, 8).Value
    
        End If
     
        If Cells(i, 1).Value = "DQ" And Cells(i - 1, 1).Value <> "DQ" Then
    
            'set starting price
            startingPrice = Cells(i, 6).Value
    
         End If
    
        If Cells(i, 1).Value = "DQ" And Cells(i + 1, 1).Value <> "DQ" Then
            
            'set ending price
            endingPrice = Cells(i, 6).Value

    End If
The analysis revealed that "DAQO" had been performing poorly over the last year. The client requested the ability to analyse all stocks over multiple years. To accomplish this  the program refered to as "AllstockAnalysis" was created.

#### Sub AllStocksAnalyis
Sub AllStocksAnalsys()
    Dim starttime As Single
    Dim endtime As Single
    
    yearvalue = InputBox("what year would you like to run the analysis on:")
    
        starttime = Timer
        
'1)Format the output sheet on the "All Stocks Analysis" Worksheet.

    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearvalue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
'2) Initialize an array of all tickers.
    
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
        
'3a) Initialize variables for the starting price and ending price

    Dim startingPrice As Double
    Dim endingPrice As Double
    
'3b) Activate the data worksheet

    Worksheets(yearvalue).Activate
    
'3c) Find the number of rows to loop over.
     
     RowCount = Cells(Rows.Count, "A").End(xlUp).Row

'4) Loop through the tickers

    For i = 0 To 11
        Ticker = tickers(i)
        totalvolume = 0
        
'5) Loop through rows in the data

    Worksheets(yearvalue).Activate
    For j = 2 To RowCount
    
'5a) Find total volume for the current ticker
    
    If Cells(j, 1).Value = Ticker Then
    
        totalvolume = totalvolume + Cells(j, 8).Value
        
    End If
    
'5b) Find starting price for the current ticker.

    If Cells(j, 1).Value = Ticker And Cells(j - 1, 1).Value <> Ticker Then
    
            'set starting price
            startingPrice = Cells(j, 6).Value
            
    End If
    
'5c)  Find ending price for current ticker

    If Cells(j, 1).Value = Ticker And Cells(j + 1, 1).Value <> Ticker Then
            
            'set ending price
            endingPrice = Cells(j, 6).Value

    End If
    
Next j

'6) Output the data for the current ticker
    
    
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = Ticker
    Cells(4 + i, 2).Value = totalvolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    
Next i

To further aid the client's ability to analyse the data easily, further coding was done to allow for formatting the data.

#### Formatting the Data

Sub formatAllStocksAnalysisTable()

    'formatting

    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.00%"


    Range("A1").Font.Bold = True
    Range("A1").Font.Underline = xlUnderlineStyleSingle
    Range("A1").Font.Italic = True

    Range("A4:A15").Font.Bold = True

    Columns("B").AutoFit

    Worksheets("all stocks analysis").Activate

    datarowstart = 4
    datarowend = 15
    For i = datarowstart To datarowend

    If Cells(i, 3) > 0 Then
    
        'color the cell green
        Cells(i, 3).Interior.Color = vbGreen
    
    ElseIf Cells(i, 3) < 0 Then
    
        'color the cell red
        Cells(i, 3).Interior.Color = vbRed
        
    Else
    
        'Clear the cell color
        Cells(i, 3).Interior.Color = xlNone
        
    End If
   

Next i
    
End Sub

Because the client needs is often working with financial clients, the client requested to the ability to be able to see how fast the program was producing the desired data.  A program measuring the results was creatied with the following outcomes.

##### Report Time for 2017
[Analysis for 2017](https://github.com/stephenanayashilliard/Stock-Analysis/blob/master/Greenstock%202017.png)

##### Report Time for 2018
[Analysis for 2018](https://github.com/stephenanayashilliard/Stock-Analysis/blob/master/Greenstock%202018.png)

As you can see the run times for 2017 and 2018 were 0.453125 seconds and 0.5859375 seconds respectively.  The client then asked if it would be possible to have the formula's run even faster.  This was especially important as more stock date will be added in the future.  To accomplice this,  The code would have to be refactored by switching the nesting order of the loops and using arrays.  

#### The Refactored Code
  'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
        
        Dim tickerIndex As Integer
        tickerIndex = 0
    

    '1b) Create three output arrays
        
        Dim tickerVolumes(11) As Long
        Dim tickerStartingPrices(11) As Single
        Dim tickerEndingPrices(11) As Single
        
        
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For i = 0 To 11
    tickerVolumes(i) = 0
    
    Next i
    
    ' If the next row’s ticker doesn’t match, increase the tickerIndex.
    
    '2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
    '3a) Increase volume for current ticker
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    

    '3b) Check if the current row is the first row with the selected tickerIndex.
    'If  Then
     
     If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    'End If
    End If
    
  
     
        
    '3c) check if the current row is the last row with the selected ticker
    'If  Then
     
     If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
     
    '3d Increase the tickerIndex.

    tickerIndex = tickerIndex + 1
            
            
    'End If
    End If

    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
    For i = 0 To 11
    
        
    Worksheets("All Stocks Analysis").Activate
    tickerIndex = 1
    Cells(i + 4, 1).Value = tickers(i)
    Cells(i + 4, 2).Value = tickerVolumes(i)
    Cells(i + 4, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1

Above is the portion of the code that was refractored.  The refractored code produced the following run times:

##### The Refactored Run time for 2017
[VBA_Challenge_2017](https://github.com/stephenanayashilliard/Stock-Analysis/blob/master/VBA_Challenge_2017.png)

##### The Refactored Run time for 2018
[VBA Challenge 2018](https://github.com/stephenanayashilliard/Stock-Analysis/blob/master/VBA_Challenge_2018.png)

The new runtimes were 0.453125 seconds for 2017 and 0.09375  seconds 2018.

## Summary
The final project did provide the automated analysis the client requested.   However, when you compare the original programing to the refactored program, the process the refactored report for 2017 was actually slight over 7 seconds slower than the original program.  Only the refractored 2018 report showed any increase in efficiency.

### Advantages and Disadvantages of Refactoring Code
### Comparison of the Advantages and Disadvantages between the Original and Refactored VBA Script.
