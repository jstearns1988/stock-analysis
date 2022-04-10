# Analysis of Stock Performances Using VBA

## Overview of Stock Analysis 
To assist Steve in his research of the stock market in 2017 and 2018 for his first clients (his parents) after graduating with his finance degree, an analysis on the performance of several different stocks was performed. This analysis was originally created to determine the returns for Daqo (DQ) and how actively it was traded in 2018. Through the DQ Analysis, a 63% drop in 2018 was discovered, and Steve had the data to present to his clients the more successful options available. I then refactored the VBA script for a faster, more optimal performance.

### Refactored VBA Code for Stock Analysis
```
   '1a) Create a ticker Index
    tickerIndex = 0

'1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

''2a) Create a for loop to initialize the tickerVolumes to zero.
' If the next row’s ticker doesn’t match, increase the tickerIndex.
    For I = 0 To 11
        tickerVolumes(I) = 0
        tickerStartingPrices(I) = 0
        tickerEndingPrices(I) = 0
    Next I

''2b) Loop over all the rows in the spreadsheet.
    For I = 2 To RowCount

    '3a) Increase volume for current ticker
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(I, 8).Value
    
    '3b) Check if the current row is the first row with the selected tickerIndex.
    'If  Then
    If Cells(I, 1).Value = tickers(tickerIndex) And Cells(I - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(I, 6).Value
    End If
    
    '3c) check if the current row is the last row with the selected ticker
    'If  Then
     If Cells(I, 1).Value = tickers(tickerIndex) And Cells(I + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(I, 6).Value
     End If

        '3d Increase the tickerIndex.
         If Cells(I, 1).Value = tickers(tickerIndex) And Cells(I + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If

Next I

'4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
For I = 0 To 11
    
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + I, 1).Value = tickers(I)
    Cells(4 + I, 2).Value = tickerVolumes(I)
    Cells(4 + I, 3).Value = tickerEndingPrices(I) / tickerStartingPrices(I) - 1
    
Next I
```


## Results

### Stock Performance in 2017
In 2017, all stocks included in the analysis had a positive return, aside from TERP, who took a 7.2% loss.
![2017 Analysis](https://github.com/jstearns1988/stock-analysis/blob/main/All%20Stocks2017.png)
 
### Stock Performance in 2018
In 2018, all stocks had negative returns except for ENPH and RUN, who both returned at over 80%. It is to be noted RUN had a significant percentage jump in returns at 78.45%. While also still having a positive return, the percentage for ENPH dropped from 129.5% to 81.9%.
![2018 Analysis](https://github.com/jstearns1988/stock-analysis/blob/main/All%20Stocks2018.png)

### Execution Times of Original Script

#### 2017
![2017 Execution Time](https://github.com/jstearns1988/stock-analysis/blob/main/2017%20Performance%20Original%20Code.png)

#### 2018
![2018 Execution Time](https://github.com/jstearns1988/stock-analysis/blob/main/2018%20Performance%20Original%20Code.png)

### Execution Times of Refactored Script

#### 2017
![2017 Refactored](https://github.com/jstearns1988/stock-analysis/blob/main/VBA_Challenge_2017.png)

#### 2018
![2018 Refactored](https://github.com/jstearns1988/stock-analysis/blob/main/VBA_Challenge_2018.png)

## Summary

### Advantages or Disadvantages of Refactoring the Code
#### Advantages 
Refactoring the code to loop through all the data one time is advantageous to the effiency of the script running time. Seen above in the 2017 and 2018 original and refactored script images, running the VBA scripts for both 2017 and 2018 improved from approximately .28 seconds to .08 seconds.
A slight disadvantage to refactoring the code would be the potential of altering the results we originally achieved if any variables were forgotten when re-writing. This could be fixed easily with debugging due to the cleaner code, but is still an inconvenience.


### How these Pros and Cons Apply to Refactoring the Original VBA Script
A less obvious efficiency is the ease of use repurposing this code if Steve wants to add more stocks to analyze, or if other clients request a similar service. Before the refactoring, there were multiple different subroutines to achieve the same result. This could be more challenging if anyone needs to alter the code in the future, potentially causing bugs that take extra time to resolve. Refactoring the code and explaining each step taken in the comments makes it more simple to read and edit.
