# Green Stock Analysis (Module 2 Challenge)
## Project Overview
Steve wants to analyze a green stock for his parents to determine if it is worth investing in. In order to this purpose, the Visual Basic Application(VBA) in Excel will be used to calculate the stock's total daily volume and annual return. After that, the rest 11 green stocks are also needed to find how the first one compared to them. Then, Steve will know what would be the best option for his parent according to the summary.

Here is the link for Green Stock Excel file: [Green_Stock](https://github.com/cffhr99/Module2-Challenge/raw/main/VBA_Challenge.xlsm)

## Purpose
The purpose of this project was to find an efficient way to analyze multiple stocks by VBA. After analyzing these 12 different stocks, it was obivous that there was a more efficient method to analyze the data. In order to do that, it is necessary to refactor the code and determine the refactored code more efficent.

Here is the link for refactored code :[VBA_Challenge.vbs](https://github.com/cffhr99/Module2-Challenge/raw/main/VBA_challenge.vbs)

## Results

### Refactoring the code
In order to do more efficient, the first job is to switch the nesting order of the for loops. There are four new arrays created, which are *tickerIndex*, *tickerVolumes*, *tickerStartingPrices* and *tickerEndingPrices*.
#### Refactored Code

    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    Dim tickerIndex As Single
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    tickerVolumes(i) = 0
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickIndex) + Cells(i, 8).Value
        

       
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            

            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        tickerIndex = i
        Cells(i + 4, 1).Value = tickers(tickerIndex)
        Cells(i + 4, 2).Value = tickerVolumes(tickerIndex)
        Cells(i + 4, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
        
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


#### Old Code
    '3a) Initialize variables for starting price and ending price
     Dim startingPrice As Single
     Dim endingPrice As Single
     '3b) Activate data worksheet
     Worksheets("2017").Activate
     '3c) Get the number of rows to loop over
     RowCount = Cells(Rows.Count, "A").End(xlUp).Row

     '4) Loop through tickers
       For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       '5) loop through rows in the data
       Worksheets("2018").Activate
       For j = 2 To RowCount
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If

           '5c) get ending price for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If
       Next j
       '6) Output data for current ticker
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

     Next i
   
      endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
     End Sub


These variables assigned the *tickerVolumes*, *tickerStartingPrices* and *tickerEndingPrices* to each ticker symbol before interating through the dataset. Therefore, the analysis would run more faster than the nested for loops.

### Run-time for Each Method on 2017 and 2018
Below pictures are the run-times from old code and refactored code.

#### Refactored Code run-times
![new 2017](https://github.com/cffhr99/Module2-Challenge/blob/main/Resources/VBA_Challenge_2017.png?raw=true)
![new_2018](https://github.com/cffhr99/Module2-Challenge/blob/main/Resources/VBA_Challenge_2018.png?raw=true)

#### Old Code run-times
![old_2017](https://github.com/cffhr99/Module2-Challenge/blob/main/Resources/2017_old.png?raw=true)
![old_2018](https://github.com/cffhr99/Module2-Challenge/blob/main/Resources/2018_old.png?raw=true)

Based on the run-times, it is obvious that the refactored code runs about 0.5 seconds faster than the original code, which means the refactored code is more efficient.

## Summary
1) The major advantage of refactoring code is improve the efficent of the analysis. However, the disadvantage of refactoring code is that the analysis will become more instability. Since the original code works fine, the analysis will be unusable if the refactoring work has some bugs. Therefore, refactoring code is always necessary when it will improve the analysis a lot and the original code is saved well.
2) The advantage of refactoring code in VBA is that you can compare the refacoring code and original code. Side by side windows can make the refactoring process more easier. However, the disadvantage is that the VBA requires a strong understanding of syntax and logic algorithm. VBA does not show a live window results so syntax is necessary. And a good understanding of logic alogrithm will help the refactored code more efficient.
