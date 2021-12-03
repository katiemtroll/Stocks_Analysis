# Stock Analysis

## Overview of Project

This workbook was created to help a stock trader evaluate the best stocks to invest in. 

### Purpose

Steve will utilize this Excel workbook to determine which of 12 different stocks would be the best option for his parents to invest in. Steve’s parents initially wanted to invest in DAQO (DQ) stock. 


## Results

### DQ Analysis
Steve’s parents initially wanted to invest in DAQO (DQ) stock. An analysis was conducted on DQ’s 2018 stock to determine if the annual return indicated that it would be a good investment. The analysis showed that in 2018, DQ had a negative return of -62.6%. With the low return, Steve decided he wanted an analysis of 12 stocks over two years to help his parents select a better option.

### All Stocks Analysis
The first iteration of analyzing the 12 stocks for the years 2017 and 2018 provided promising results for Steve to choose a stock to invest in (see analysis results below). 

    *All Stocks Analysis (2017)*
![Screen Shot 2021-12-02 at 8 40 21 PM](https://user-images.githubusercontent.com/94259442/144536542-d78a12d3-68fd-47bb-ae68-2db5b643b8ab.png)

    *All Stocks Analysis (2018)*
![Screen Shot 2021-12-02 at 8 40 40 PM](https://user-images.githubusercontent.com/94259442/144537065-4d00b4b5-37ab-4c87-82b8-cd668b332bf0.png)

However, as Steve began to think of future uses for this workbook, it was decided that the code should be refactored to make the analyses more efficient.

### Refactored All Stocks Analysis
The initial iteration of analyses ran in 0.2773 seconds and 0.2734 seconds in 2017 and 2018, respectively.

<img width="272" alt="Module Run Time 2017" src="https://user-images.githubusercontent.com/94259442/144537190-f19e9d3a-7fef-49d7-8d92-16db4aec317a.png">

<img width="272" alt="Module Run Time 2018" src="https://user-images.githubusercontent.com/94259442/144537207-dd7fc6a9-e7e2-4b01-ad92-b6b5fb220344.png">

With the following refactoring steps, it should be anticipated that the run time decreases.

#### Refactoring
The code was refactored in a couple of places (see notes below) to increase efficiencies in the general structure and overall run time. 
1. A variable called `TickerIndex` was created and set to zero before looping through the data. This variable was used to access the correct index within each of the arrays in the code.

2. Arrays were created for `tickers`, `TickerVolumes`, `TickerStartingPrices`, and `TickerEndingPrices`. The output arrays provided a simpler, cleaner solution for analysis output. `TickerIndex` was used to access the stock ticker index for each of these arrays. The code looped through the stock data utilizing the arrays and output the results to the “All Stocks Analysis” tab.

Refer to the **Code** subsection to view the initial code compared to the refactored code.

## Summary & Conclusion

After the code was refactored, it can be noted that the run time is faster at 0.078125 seconds and 0.078125 seconds for 2017 and 2018, respectively. 

  *All Stocks Analysis (2017) - New Run Time*
<img width="274" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/94259442/144537650-1ad8e008-994f-4d2f-b878-c08b92222b47.png">

  *All Stocks Analysis (2018) - New Run Time*
<img width="277" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/94259442/144537657-1c67082d-c6a2-496e-8830-160940b431c1.png">


In general, key advantages of refactoring the code are increased quality and readability, and the code should run more efficiently. The risk of refactoring code is that you could break the code if you refactor incorrectly, or if you don’t have a strong understanding of the syntax of the language. In these instances, it may be better not to refactor the code, or to make sure that the original code is saved and changes are made little by little.

It is difficult to say if the time required to refactor this VBA code was justified by the changes in output. The change in run times was minimal between the initial and refactored codes, and the initial code may have likely been acceptable as-is. However, because we refactored the code, it should be much easier to return to the code to make future changes because of the improved structure and quality of the code.

## Code
### Initial Code
'Sub AllStocksAnalysis()
' Activate the sheet where we want to place the returned data
Worksheets("All Stocks Analysis").Activate
    
' Set formatting/variable naming conventions
    ' Create headers
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    'Format Headers & Cells
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B:B").NumberFormat = "#,##0"
    Range("C:C").NumberFormat = "0.00%"
    Columns("A:C").ColumnWidth = 15
    Range("A:C").HorizontalAlignment = xlCenter

'Determine total run time of analysis
Dim StartTime As Single
Dim EndTime As Single

'Creat input for any year instead of defined year.
Do
YearValue = InputBox("What year would you like to run the analysis on?" & Chr(10) & "Box will re-open with invalid entry.")
    Dim wsExists As Boolean
    For i = 1 To Worksheets.Count
        If Worksheets(i).Name = YearValue Then
            wsExists = True
        ElseIf StrPtr(YearValue) = 0 Then
            Exit Sub
        End If
    Next i
Loop Until wsExists = True
    
    'Update year for title based on input box entry
    Cells(1, 1).Value = "All Stocks (" & YearValue & ")"
    
    StartTime = Timer
    
    'Set up our array of tickers (this prevents having to re-write code for each string).
    Dim tickers(11) As String

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
        
    ' Determine start/end of DQ data and use the data accordingly
    Dim StartingPrice As Single
    Dim EndingPrice As Single
    
' Activate the sheet from which we will run our analysis
Worksheets(YearValue).Activate
    ' note that we don't have to use Dim because we are assigning specific values. Could use Dim; would need to if only assigning as a variable without value here.
    RowStart = 2
    'rowEnd code taken from https://stackoverflow.com/questions/18088729/row-count-where-data-exists
    RowEnd = Cells(Rows.Count, "A").End(xlUp).Row

    ' loop through the tickers
    For i = 0 To 11
        Ticker = tickers(i)
        'move total volume inside so that it resets for each ticker
        TotalVolume = 0
        
            'loop through the data for each ticker type
            Worksheets(YearValue).Activate
        
            For j = RowStart To RowEnd
                ' If the cell value is equal to the ticker array
                If Cells(j, 1).Value = Ticker Then 'then set the total volume for the ticker to previous total volume plus the new volume of that ticker cell
                    TotalVolume = TotalVolume + Cells(j, 8).Value
                End If
                
                'Now check if it's the start or ending price
                If Cells(j - 1, 1).Value <> Ticker And Cells(j, 1).Value = Ticker Then
                    StartingPrice = Cells(j, 6).Value
                 End If
                If Cells(j + 1, 1).Value <> Ticker And Cells(j, 1).Value = Ticker Then
                    EndingPrice = Cells(j, 6).Value
                End If
                
            Next j
        
    ' Return back to the worksheet where we're placing data and include the updates inside the loop so that it changes for each i.
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = Ticker
    Cells(4 + i, 2).Value = TotalVolume
        ' this is the yearly return calculation
    Cells(4 + i, 3).Value = (EndingPrice / StartingPrice) - 1
    
        'Color formatting
        ColorStart = 4
        ColorEnd = Cells(Rows.Count, "A").End(xlUp).Row
        
        For k = ColorStart To ColorEnd
    
            If Cells(k, 3) > 0 Then
                Cells(k, 3).Interior.Color = vbGreen
                ElseIf Cells(k, 3) < 0 Then
                    Cells(k, 3).Interior.Color = vbRed
            Else
                Cells(k, 3).Interior.Color = xlNone
            End If
        Next k
        
    Next i

    EndTime = Timer
    MsgBox "This code ran in " & Format((EndTime - StartTime), "#,##0.0000") & " seconds for the year " & (YearValue) & "."
    
End Sub
'

### Refactored Code
'
Sub AllStocksAnalysisRefactored()

    Dim StartTime As Single
    Dim EndTime  As Single

    'Creat input box for year selection; box will re-open with invalid entry; selecting "Cancel" will exit the subroutine (close the input box).
    Do
    YearValue = InputBox("What year would you like to run the analysis on?" & Chr(10) & "Box will re-open with invalid entry.")
        Dim wsExists As Boolean
        For i = 1 To Worksheets.Count
            If Worksheets(i).Name = YearValue Then
                wsExists = True
            ElseIf StrPtr(YearValue) = 0 Then
                Exit Sub
            End If
        Next i
    Loop Until wsExists = True

    StartTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    'Set cell A1 to identify the year of the analysis
    Range("A1").Value = "All Stocks (" + YearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    'Initialize array of all tickers
    Dim tickers(11) As String
    
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

    'Activate the data worksheet
    Worksheets(YearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
        
    '1a) Create the ticker index variable and set to zero
    TickerIndex = 0
 
    
    '1b) Create three output arrays; we know the array is 12 b/c it needs to be the same size as the ticker array
    Dim TickerVolumes(11) As Long
    Dim TickerStartingPrices(11) As Single
    Dim TickerEndingPrices(11) As Single
    
    '2a) Create a for loop to initialize the TickerVolumes to zero.
    For i = 0 To 11
            TickerVolumes(i) = 0
    Next i
    
        '2b) Create a for loop that will loop over all of the rows in the spreadsheet.
        For j = 2 To RowCount
                
        '3a) Increase volume for the current ticker
                TickerVolumes(TickerIndex) = TickerVolumes(TickerIndex) + Cells(j, 8).Value
                 
         '3b) 'Check if the current row is the first row with that TickerIndex
            If Cells(j - 1, 1).Value <> tickers(TickerIndex) And Cells(j, 1).Value = tickers(TickerIndex) Then
                'store new value
                TickerStartingPrices(TickerIndex) = Cells(j, 6).Value
            End If
        '3c) Check if the current row is the last row with the selected ticker
            If Cells(j + 1, 1).Value <> tickers(TickerIndex) And Cells(j, 1).Value = tickers(TickerIndex) Then
               'store new value
                TickerEndingPrices(TickerIndex) = Cells(j, 6).Value
            End If
            '3d) Increase the Ticker Index if it's the last row with the index
             If Cells(j + 1, 1).Value <> tickers(TickerIndex) And Cells(j, 1).Value = tickers(TickerIndex) Then
                'store new value
                TickerIndex = TickerIndex + 1
            End If
        Next j

    '4) Loop through the arrays to output the Ticker, Totaly Daily Volume, and Return.
    For k = 0 To 11
        Worksheets("All Stocks Analysis").Activate
            TickerIndex = k
            Cells(4 + k, 1).Value = tickers(TickerIndex)
            Cells(4 + k, 2).Value = TickerVolumes(TickerIndex)
            Cells(4 + k, 3).Value = (TickerEndingPrices(TickerIndex) / TickerStartingPrices(TickerIndex)) - 1
    Next k
        
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    DataRowStart = 4
    DataRowEnd = 15

    For m = DataRowStart To DataRowEnd
        
        If Cells(m, 3) > 0 Then
            
            Cells(m, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(m, 3).Interior.Color = vbRed
            
        End If
        
    Next m
    
    EndTime = Timer
    
    MsgBox ("The refactored code ran in " & (EndTime - StartTime) & " seconds for the year " & (YearValue) & "." & Chr(10) & "To run this analysis again,select the 'Run Analysis' button.")

End Sub
'
