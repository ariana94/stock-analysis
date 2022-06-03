# VBA Stock Analysis

## Overview
The intent of this analysis project is to analyze stocks from 2017 and 2018 to find out what stocks are worth investing in.
We started off looking into one stock, "DQ", that turned out to not be a great investment. Now we are able to monitor twelve different stocks to find better ones to invest in. Once we achieved getting the VBA code to do what we wanted, we aimed to make it more efficient by refacrtoring it.

---

## The Results
### The Original Macro

In our original code, we did a loop for each of the twelve ticker symbols.
```
 For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       
       '5) loop through rows in the data
       Sheets(yearValue).Activate
       
       For j = 2 To RowCount
           '5a) Get total volume for current ticker
               If Cells(j, 1).Value = ticker Then
     'increase totalVolume
             totalVolume = totalVolume + Cells(j, 8).Value
            End If
            
            '5b) get starting price for current ticker
    If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
    'set starting price
             startingPrice = Cells(j, 6).Value
            End If
            
            '5c) get ending price for current ticker
    If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
    'set ending price
             endingPrice = Cells(j, 6).Value
            End If


       Next j
       '6) Output data for current ticker
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    
Next i
```
It ran at a speed around 0.51 seconds for 2017 and 0.52 seconds for 2018

![Original run time 2017](https://user-images.githubusercontent.com/19378130/171759364-e0b2b866-afcd-40e9-ab19-df2cef0d51ef.png)
![Original run time2018](https://user-images.githubusercontent.com/19378130/171759425-60132c9f-81b3-42d3-ab2d-33227fb90a60.png)


### The Refactored Macro

In our refactored macro, we instructed it to only read through all the stock data one time.
```
For i = 0 To 11
    tickerVolumes(i) = 0
    tickerStartingPrices(i) = 0
    tickerEndingPrices(i) = 0
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            
        End If
        
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            End If
            

            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
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
```
    
This showed an improvement in the speed, taking only 0.078 seconds for 2017 and 2018
    
![VBA_Challenge_2017](https://user-images.githubusercontent.com/19378130/171759660-b2881357-db7b-4931-a3f6-7256a33f7f56.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/19378130/171759666-1a8e180c-d03a-4a0a-aa66-61efa1953ca4.png)

## In Conclusion
### General Pros and Cons to Refactoring
In general, refactoring code is a great way to tidy things up and organize your macros. This could potentially make them more efficient by reducing resources used and improving the overall performance. However, it can be time consuming to figure out and implement the best way to refactor your code. If the code in question isn't something that is intended to be used constantly, it may not be as worth it to invest the manpower into refactoring it just to shave off a few micro-seconds.

### Advantages and Disadvantages of Our Refactored Code
Refactoring our code did show a fairly drastic improvment in speed, about seven times faster than the original. Following and understanding the macro is also easier without having a clunky nested for loop. The biggest disadvantage that I see to refactoring our code is that it took so much time compared to how much time it saved.
