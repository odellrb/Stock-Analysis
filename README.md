# Stock-Analysis
Repository address: https://github.com/odellrb/Kickstarter-anaylsis-.git
## Overview of Project
The purpose of the project was to help Louise understand trends based on launch dates and funding from Kickstarter campaigns, to help 
her launch her play "Fever".

## Results
We had to isolate down to the most successful dates and determine which monetary goals where more realistic 
than others. Trying to determine proper launch date and the right goal amount were the challenges.

![2017_Speed_box](/Resources/VBA_Challenge_2017.png)


![2018_Speed_box](/Resources/VBA_Challenge_2018.png)

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

- What are two conclusions you can draw about the Outcomes based on Launch Date?

  May and June are the best times to launch. April and especially October were not good
  times to launch.
- What can you conclude about the Outcomes based on Goals?

  A goal of 5,000 dollars or less had the best results for success.
- What are some limitations of this dataset?

  The limitations of this dataset do not paint the whole picture to be succussful. We should look into 
  other variables to get a better idea of how to possibly crowdfund more money and backers.
- What are some other possible tables and/or graphs that we could create?

  We could use Pie charts, bar graphs, and possibly a Box and whisker chart if one is interested enough for it.
