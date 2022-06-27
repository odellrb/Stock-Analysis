# Stock-Analysis
Repository address: https://github.com/odellrb/Stock-Analysis.git
## Overview of Project
The purpose of the project was to help Steve compile the results of a dozen different stocks over two years.
Steve wanted to help his parnets with makeing some good choices on stocks that have yielded consistant returns.
We needed to refactor the code to help the program run quicker when Steve would anylyze a thousand stocks instead of 
just twelve. The end goal of refactoring the code was to help the program run more efficeintly.


## Results
After refactoring the code, the program runs about 22 milliseconds faster than before. Before refactoring the program would run at 29 millisecond where
as now it runs on average at 07 milliseconds. The code below gave the processor less work with simplier commands to execute the program.

![VBA_Speed Box_2017](/Resources/VBA_Challenge_2017.png)





![VBA_Speed_Box_2018](/Resources/VBA_Challenge_2018.png)

    
    '1a) Create a ticker Index
          tickerIndex = 0
    
    '1b) Create three output arrays
         Dim tickerVolumes(11) As Long
         Dim tickerStartingPrices(11) As Single
         Dim tickerEndingPrices(11) As Single

    '2a) Create a for loop to initialize the tickerVolumes to zero.
     'If the next row’s ticker doesn’t match, increase the tickerIndex.
           For i = 0 To 11
           tickerVolumes(i) = 0
    
     Next i

     '2b) Loop over all the rows in the spreadsheet.
           For i = 2 To RowCount

        '3a) Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
    
    '3c) check if the current row is the last row with the selected ticker
    'If  Then
           If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
           tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
     End If

        '3d Increase the tickerIndex.
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If

     Next i

     '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
          For i = 0 To 11
    
        Worksheets("All Stocks Analysis").Activate
          Cells(4 + i, 1).Value = tickers(i)
          Cells(4 + i, 2).Value = tickerVolumes(i)
          Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
    Next i

## Summary
The advatages of refactoring code makes the program more oraganized and run more efficeintly and puts less strain on the processor. The disadavantage of refactoring code is that we have to go back in and touch something that already works to begin with. When we go back into code theres always a chance that 
we will end up fat fingering a few lines of code or spending more time on researching and coming up with ideas how we can stream line our code which can take time when you factoring possible human error and tinkering. 



