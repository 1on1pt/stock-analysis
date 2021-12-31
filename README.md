# Stock Analysis Tool Developement
Developing a stock analysis tool using Visual Basic for Applications (VBA) in Excel.
## Overview of Project
A good friend of mine, Steve, recently graduated with his finance degree and his proud parents have become his first clients.  His parents are passionate about green energy, so they have invested all their funds in DAQO (DQ).  Steve, keeping his parent's best interests first, wants to perform an analysis on DQ, as well as looking at several other green energy stocks to assure that their portfolio is well diversified.  He has fairly good handle on Excel and its functionality, but is asking for assistance in completing his stock analysis to assure accuracy and efficiency.  Steve will learn of the power and functionality of VBA, which in turn ultimately helps him accurately and efficiently analyze not only a small porfolio, but the entire stock market.

### Purpose
The purpose of this project is to macro-enable the Excel workbook so that VBA can be used to develop a tool that will accurately and efficiently analyze stocks.  Initially, an analysis was performed on DQ determining the stock's total daily volume and return for 2018.  Then an analyis of 11 additional green stocks was completed to find those with the best returns.  But the original code did not appear to be the most efficient code.  The original code was *refactored* with the idea of taking fewer steps, using less memory, and improving logic so the analysis would be more efficient.  The outcome of this project looks to determine if the refactored code *is* more efficient than the original code.

## Results
### Original Code
The *original code* contained a "nested loop", which ultimately resulted in additional steps and using more memory in determing the output of **Total Daily Volume** and **Return**.  See below the original code with the nested loop.
Sub AllStocksAnalysis()

    Dim startTime As Single
    Dim endTime As Single
    
    yearValue = InputBox("What year would you like to run the analysis on?")
    
        startTime = Timer
        
    '1) Format the output sheet on the "All Stocks Analysis" worksheet.

    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
        
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

'2) Initialize an array of all tickers.

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

'3a) Initialize variables for the starting price and ending price.

Dim startingPrice As Single
Dim endingPrice As Single

'3b) Activate the data worksheet.

    Sheets(yearValue).Activate

'3c) Find the number of rows to loop over.

RowCount = Cells(Rows.Count, "A").End(xlUp).Row

'4) Loop through the tickers.

For i = 0 To 11

    ticker = tickers(i)
    totalVolume = 0

    '5) Loop through rows in the data.
    
    Sheets(yearValue).Activate
        For j = 2 To RowCount
    
    '5a) Find total volume for the current ticker.
    
    If Cells(j, 1).Value = ticker Then
        
            totalVolume = totalVolume + Cells(j, 8).Value
            
    End If
               
    '5b) Find starting price for the current ticker.
    
    If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
    
        startingPrice = Cells(j, 6).Value
        
    End If
            
    '5c) Find ending price for the current ticker.
    
    If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
    
        endingPrice = Cells(j, 6).Value
        
    End If
    
    Next j
    
'6) Output the data for the current ticker.

    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    
Next i

    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    
End Sub

## Summary
