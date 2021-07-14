# Stocks-analysis

## Overview of Project
Steve’s parents have a great interest in green energy productions, so they are going to invest all their money into Daqo New Energy Corporation, a company that makes silicon wafers for solar panels. However, they haven’t done any research about green energy stocks. To help them make a great decision on investing, I have provided the analysis report to display the entire stocks performance. Furthermore, the main purpose of this project is to improve efficiency of analysis process by refactoring the original Microsoft Excel VBA code.

## Results
1. Below is the refactored code with instructions. Here are some differences from the original code:
- create the arrays for three output value: tickerVolumes, tickerStartingPrices and tickerEndingPrices (1b) 
- create a variable, tickerIndex (1a) to access the correct index for the four different arrays and get the value of tickerVolumes, tickerStartingPrices and tickerEndingPrices for each ticker (2a)
- Loop through the arrays and output the data (4)

```VBA
'1a) Create a ticker Index
    tickerIndex = 0
       
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12), tickerEndingPrices(12) As Single

    '2a) Create a for loop to initialize the tickerVolumes to zero.
    For j = 0 To 11 
    tickerVolumes(j) = 0 
    Next j

'2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
        
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1) = tickers(tickerIndex) And Cells(i - 1, 1) <> tickers(tickerIndex) Then

        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value

        End If

        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i, 1) = tickers(tickerIndex) And Cells(i + 1, 1) <> tickers(tickerIndex) Then

        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

        '3d Increase the tickerIndex.
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
```
2. Stocks performance between 2017 and 2018


###### 2017 stock performance versus 2018 stock performance
<img src="https://github.com/BessHung/Stocks-analysis/blob/31b7918bbc6afa0a202afd316f00775a7ed9403c/Resources/stock_performance_2017.png" width=35% height=35%> <img src="https://github.com/BessHung/Stocks-analysis/blob/31b7918bbc6afa0a202afd316f00775a7ed9403c/Resources/stock_performance_2018.png" width=35% height=35%>

3. Execution times of the original script and the refactored script.

###### 2017 original script versus 2017 refactored script
<img src="https://github.com/BessHung/Stocks-analysis/blob/68794e5e88b88cc386912eb12dbe57e8014cedf1/Resources/VBA_original_2017.png" width="400" height="200"> <img src="https://github.com/BessHung/Stocks-analysis/blob/68794e5e88b88cc386912eb12dbe57e8014cedf1/Resources/VBA_Challenge_2017.png" width="400" height="200">


###### 2018 original script versus 2018 refactored script
<img src="https://github.com/BessHung/Stocks-analysis/blob/68794e5e88b88cc386912eb12dbe57e8014cedf1/Resources/VBA_original_2018.png" width="400" height="200"> <img src="https://github.com/BessHung/Stocks-analysis/blob/68794e5e88b88cc386912eb12dbe57e8014cedf1/Resources/VBA_Challenge_2018.png" width="400" height="200">

## Summary
