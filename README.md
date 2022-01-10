# Stocks_Analysis

## Overview of Project

### Purpose
- The purpose of this project is the help Steve gather information on all of his stock volumes and stock return. We're using VBA writing code that can run through all stocks data over the last few years, then group them into a table that can easily show the result. Considering there will be more stocks added to the data in the future, we also refactoring the code that will be taking fewer steps, using less memory, and making the VBA script run faster for the same set of data.

## Analysis and Result

- Compare stock performance between 2017 and 2018, as well as teh execution times of the original scipt and refactored script.

 ### Orginal Script: 
    
    In the original script, we're using two FOR loops to run the stock data, this will take more time. We first create a For loop for tickers to loop through different 12 tickers ("AY", "CSIQ", "DQ", etc), one by one. Then create another For loop inside of the Ticker FOR loop to look through each row and see if it matches the current ticker. 
    
    For example, for i=0, we first looking at ticker(0) = "AY". Then we loop through all rows from row 2 to the last row to see if the cell value matches "AY". If match, we adding teh totalVolume. If the current cells is matched as "AY" but the previous one is not, then we record down the close price as starting price. If the current cells are matched as "AY" but the next cell does not match, then we record down the close price as the ending price. Now Ticker FOR Loop going to i=1, ticker(1) = "CSIQ" then we looped through all rows again to see if cell match to "CSIQ". Concludes that each time we're looking for a ticker's name, we need to loop through the whole data again. In the original script, we will need to loop through whole data 12 times for each stock.

    In the original script, we put down the value on the "All Stocks Analysis" sheet for of current ticker(i) before we started to look through the next ticker name.

        For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        
        '(5)loop through rows in the data
         'Activate data work sheet
         Worksheets(yearValue).Activate
        For j = 2 To RowCount
            
            '(5a)get total volumne for current ticker
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
            
            '(5b)get starting price for current ticker
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                startingPrice = Cells(j, 6).Value
            End If
            
            '(5c)get ending price for current ticker
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                endingPrice = Cells(j, 6).Value
            End If
            
        Next j
        'output result for current ticker
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    Next i

### Refactor Script:
    In the refactor script, we use variable tickerIndex instead of using a for a loop. This will help eliminate a big amount of time to run datas and get results. In this script, we determined the value while we run the result in each row, and use tickerIndex to help us organize the result of each ticker.
    
    For example, we first start with tickerIndex as 0. For i = 2, we first check on row 2. We record the volumes of the current ticker on this row and assign as tickerVolumes(0), which means we putting down volumes for "AY" ticker. Then we check if this current row was the first select of current ticker(0) "AY". If not the first select, then we check if this current row was the last selected of the current ticker(0) "AY". If nor first select and nor last select, then we going to next row i = 1, and check all the if statement again. If we found out the current row is the last select of current ticker(0) = "AY", then we change the tickerIndex as 1. Then, starting from the next row, all ticker volume records will be assigned to tickerVolume(1), which was the "CSIQ" ticker. We're now recording all volume for "CSIQ" ticker until we figure out the last row of "CSIQ"

    In Refactor, we put all values on the "All Stocks Analysis" sheet at the end after we run through all the rows, and figured the value for all tickers.


    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
       'Example if we are looking for 2018: loopping through different ticker from row 2 till row end.
        'tickerVolumes(0) = tickerVolumes(0) + Volume value on Column H
        'tickerVolumes(0) = keep adding next Volume Value on Column H
        'TickerIndex starting in equal to 0
        'TickerVolumes(0) is all volumes of "AY"
        'TickerVolumes(1) is all volumes of "CSIQ" , etc
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

               
        '3b) Check if the current row is the first row with the selected tickerIndex.
        
            'Example for 2018: we check that cells(1,1) is not "AY"(tickers(0)) and cell(2,1) is "AY" (tickers(0))
            'Now we know the starting price is 21.14 (column F, close price in 2018-01-02)
            
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        'Example for 2018:
        'then we keep going through, if cell(i-1 ,1)(row,column) is also "AY"
        'then tickerVolumes(0) = keep adding the volumes until the last row of "AY" ticker
        
        '3c) check if the current row is the last row with the selected ticker
        
            'Example for 2018:
            'if (i+1,1)next row value in column 1 is not "AY", but current row is "AY"
            'then set the tickerEndingprice(0) equal to the last close price on column F as "AY"
            
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            '3d Increase the tickerIndex.
            'after we went through the last row of "AY"
            'We now want to looks at new ticker. so that we're change tickerIndex as 1 use for new ticker "CSIQ"
            'after went through all last row of "CSIQ"
            'We then using tickerIndex as 2 for new ticker "DQ", etc
            tickerIndex = tickerIndex + 1

        End If

    
    Next i
    
     For i = 0 To 11
        
        'working on All Stocks Analysis worksheet
        Worksheets("All Stocks Analysis").Activate
        
        'Row 4, column A is "AY"
        'Row 4+1, column A is "CSIQ", etc
        'Row 6, column A is "DQ"
        Cells(4 + i, 1).Value = tickers(i)
        
        'Row 4, column B is totalvolume of "AY" (tickerVolume(0))
        'Row 5, column B is totalvalume of "CSIQ" (tickerVolume(1))
        'Row 6, column B is totalvolume of "DQ" (tickerVolumne(2)), etc
        Cells(4 + i, 2).Value = tickerVolumes(i)
        
        'Row 4, Column C is return of each stock after last date of closing
        'tickerStartPrices(0) and tickerEndingPrices(0) was close price for "AY"
        'tickerStartPrices(1) and tickerEndingPrices(1) was close price for "CSIQ", etc
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i

### Compares on running time
    - 2017 running time
        ![original script 2017](https://github.com/helen3121433/Stocks_Analysis/blob/main/Resources/Original_2017.PNG)
        ![Refactoring script 2017](https://github.com/helen3121433/Stocks_Analysis/blob/main/VBA_Challenge_2017.PNG)

    - 2018 running time
        ![original script 2018](https://github.com/helen3121433/Stocks_Analysis/blob/main/Resources/Original_2018.PNG)
        ![Refactoring scipt 2018](https://github.com/helen3121433/Stocks_Analysis/blob/main/VBA_Challenge_2018.PNG)

## Summary

- What are the advantages or disadvantages of refactoring code?
    - The advantage of refactoring code is that this helps the script run faster in a small amount of data, and uses less memory. It helps organize the code and make it easier to understand. The disadvantage was that data must be sorted by and grouped in a certain way. If we have a large amount of data, it could be risky. And it may be difficult to understand the code and cause more bugs.

- How do these pros and cons apply to refactoring the orginal VB script
    - Pros were that this does help us approve the execution time to run the script. Cons was that the data must be sorted by ticker name. TickerIndex will +1 every time if the ticker name of the current row was different than the ticker name of the next row. This could cause a mess when the ticker name was not sorted by and group. When I refactored the original VBA script, it took me some time to debug. 

