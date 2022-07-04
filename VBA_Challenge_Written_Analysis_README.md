# VBA Challenge

## 1. Overview of Project

### Purpose and background: 
The purpose of this Excel VBA challenge project was to apply our knowledge of writting VBA codes 
to refactor the "All Stocks Analysis" macro used earlier in this module to more quickly gather 
the 2017 and 2018 stock data for 12 different ticker stocks'"Total Daily Volume" and "Return". 

## 2. Results:

### Description of Analysis with screenshots and code: First, I downloaded  the challenge starter 
code.vbs file provided in the project description. This file already contained many of the codes 
needed to write the Refactored All Stocks Analysis macro. What I needed to create were a tickerIndex
which was used to access the stock ticker index for the three output arrays we created in this 
challenge. The three arrays created were tickerVolumes, tickerStartingPrices, and tickerEndingPrices. 
I then set the arrays data types according to the instructions. I created two different for loops, 
one to initialize the tickerVolumes to zero and the other to loop over all the rows in the spreadsheet. 
Inside the for loops, I wrote a script to increase the current tickerVolumes variable using the 
tickerIndex variable as the index. I created if-then statements to check the rows and assign the 
current starting price to the tickerStartingPrices value. I repeated this similar process to check 
and assign the closing prices to the tickerEndingPrices variable. Last, I used a for loop to loop 
through the arrays (tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices) to output 
the Ticker, Total Daily Volume, and Return for each year. I ran my refactored code and checked to 
ensure my results matched the images given in the instructions, which they did. My code ran much 
faster than the original "All Stocks Analysis Macro" and formatted the results green for return gains 
and red for returned losses. 

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
    tickerIndex = 0

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
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If Then
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

    End Sub



## 3. Summary:

### Advantages and disadvantages of refactoring code
Some advantages of refactoring code as stated in the introduction of this challenge include: 
a more efficient code, fewer steps in writing and executing the code, making the code easier 
to use and read, and faster analysis outputs. A disadvantage of refactoring code would be if
the file/application is too big. 

### Advatages and disadvantages of the original and refactored VBA script
One advantage of the refactored code is that I was able to gather the analysis for all the stocks
in 2017 and 2018 much faster than the original VBA script. Another advantage of the refactored 
script is that the formatting code for the positive and negative returns was already included in
the refactored code, and not a seperate macro like the original script. From what I gather, there
were no disadvantages of refactoring the All Stocks Analysis VBA script. 