# Module 2: VBA of Wall Street

## Overview of the Analysis

In this assignment, Stock Value data for 12 companies from 2017 and 2018 were analyzed to guide investing. Although the code worked well for a small number of companies, it needed to be refactored to make the VBA script run faster with larger datasets. 

## Results

The majority of companies provided higher returns on their stocks in 2017 rather than in 2018.

### 2017 vs. 2018 Stock Analysis

Almost all of the companies (11/12) analyzed in this project yielded positive Returns on their stocks in 2017. See Figure 1.

![Figure 1](/resources/Stock_Data_2017.png "Figure 1: Stock Data 2017")

TERP was the only company with a negative return in 2017 (-7.2%). SEDG yielded the greatest return (+184.5%)

In 2018, the majority of companies (10/12) yielded negative returns. See Figure 2. 

![Figure 2](/resources/Stock_Data_2018.png "Figure 2: Stock Data 2018")

DQ had the lowest return (-62.6%). Only ENPH and RUN yielded positive returns (81.9% and 84.0%, respectively).

### Original vs. Refactored Execution Times

The refactored code showed a significantly improved run time compared to the original code. 

The original code uses nested for-loops to iterate through the rows in the dataset and calculate total volume, starting price, and ending price for each stock "ticker."

```
For i = 0 To 11
    ticker = tickers(i)
    totalVolume = 0
    
    'loop through rows in the data
    Worksheets(yearValue).Activate
    For j = 2 To RowCount
        
        'get total volume for current ticker
        If Cells(j, 1).Value = ticker Then
            totalVolume = totalVolume + Cells(j, 8).Value
        
        'get starting price for current ticker
        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            startingPrice = Cells(j, 6).Value

        'get ending price for current ticker
        If Cells(j + 1, 1).value <> ticker And Cells(j, 1).Value = ticker Then
            endingPrice = Cells(j, 6).Value
        
        End If
    Next j

    'output data
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

Next i
```

The refactored accomplishes the same task using a single for loop, improving the code's efficiency. 

```
'Use a for-loop to initialize ticker volumes to zero
For i = 0 to 11
    tickerVolumes(i) = 0
    tickerStartingPrices(i) = 0
    tickerEndingPrices(i) = 0
Next i

'Loop through all rows in the spreadsheet. 
For i = 2 To RowCount
    
    'Increase volume for current ticker
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
    'Check if the current row is the first row with the selected tickerIndex.
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then 
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    End If   
            
    'check if the current row is the last row with the selected ticker. 'If the next row's ticker doesn't match, increase the tickerIndex.
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then 
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value    
    End If    

    'Increase the tickerIndex. 
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then 
            tickerIndex = tickerIndex + 1
    End If
    
Next i 
```

Using a single for-loop instead of a nested for-loop, the refactored code runs more efficiently than the original. For the 2017 dataset, the original code took 0.615 seconds to run, while the refactored code took 0.141 seconds, showing a 77% improvement in runtime. See Figures 3 and 4 for 2017's Original and Refactored Runtimes.

![Figure 3](/resources/Original_MB_2017.png "Figure 3: Original Runtime, 2017")

![Figure 4](/resources/VBA_Challenge_2017.png "Figure 4: Refactored Runtime, 2017")

For 2018, the original code took 0.631 seconds to run, while the refactored code took 0.172 seconds, showing a 73% improvement in runtime. See Figures 5 and 6 for 2018's Original and Refactored Runtimes.

![Figure 5](/resources/Original_MB_2018.png "Figure 5: Original Runtime, 2018")

![Figure 6](/resources/VBA_Challenge_2018.png "Figure 6: Refactored Runtime, 2018")

## Data Summary and Recommendations

There are numerous pros and cons of refactoring code.

### Advantages and Disadvantages of Refactoring Code, Generally

According to Ionos Digital Guide, Refactored Code is often simplified, understandable, and expandable, allowing greater efficiency and functionality. However, the process of refactoring can introduce bugs into the code, defeating the purpose of refactoring. Refactored code may also be less intuitive and require a higher level of understanding. [1]

### Advantages and Disadvantages of Refactoring this VBA Scripts 

The major 70+% improvement in runtime exemplifies the advantage of refactoring this script: efficiency. The dataset used in this challenge contained 11 tickers to track; a larger dataset with thousands of stocks would have taken significantly longer for the original code to compute. Although the refactored code removes the nested for-loop, it is slightly less intuitive for a reader. The refactoring process was also error-prone and took significant effort and time to complete.

## Citations

[1] IONOS Inc. (2020, September 29). Refactoring: How to improve source code. IONOS Digital Guide. Retrieved December 9, 2022, from https://www.ionos.com/digitalguide/websites/web-development/what-is-refactoring/ 