Sub AllStocksAnalysisRefactored()

    'We want to see how long it took to run all data
    Dim startTime As Single
    Dim endTime  As Single

    'InputBox, allow user to input value, since we only have two tab of 2017 and 2018,
    'if we input other value will runs error as out of range
    yearValue = InputBox("What year would you like to run the analysis on?")

    'set a timer function, which allow to start the clock / To start timing
    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    'Dimension tickers until 12 as string
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
    'Cells(Rows.Count,"A") goest to the bottome cell in column A
    '.End(xlUp) pressing End and then Up arrow in Excel. Which will go to the last cell with data in column A.
    '.End(xlUp)Use this to remove from the bottom of the sheet to the last row of data
    '.Row returns the row number
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    'Determine tickerValumes as long datatype (can hold values from -2,147,483,648 to 2,147,483,648)
    'Determine tickerStartingPrices and tickerEndingPrices as Single datatype (one floating point)
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    'reset all tickerVolumes as 0
    For i = 0 To 11
    
        tickerVolumes(i) = 0
         
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    'Activate yearvalue worksheet
    Worksheets(yearValue).Activate
    
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
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
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
    
    'Formatting
    'working on All Stocks Analysis
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        'if percentage or number is in postive, then make cell fill with green
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        'else mark cell with red
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    'endTime is click to end the timing/ stop the clock
    endTime = Timer
    
    'shows us how long we took to ran all data.
    'time when we stop the clock - time when we start the clock = how long it took
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
