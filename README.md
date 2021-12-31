# Stock Analysis Tool Developement
Developing a stock analysis tool using Visual Basic for Applications (VBA) in Excel.
## Overview of Project
A good friend of mine, Steve, recently graduated with his finance degree and his proud parents have become his first clients.  His parents are passionate about green energy, so they have invested all their funds in DAQO (DQ).  Steve, keeping his parent's best interests first, wants to perform an analysis on DQ, as well as looking at several other green energy stocks to assure that their portfolio is well diversified.  He has fairly good handle on Excel and its functionality, but is asking for assistance in completing his stock analysis to assure accuracy and efficiency.  Steve will learn of the power and functionality of VBA, which in turn ultimately helps him to accurately and efficiently analyze not only a small porfolio, but the entire stock market.

### Purpose
The purpose of this project is to *refactor* the *original* code used in VBA to determine if the tool that was developed that will be more efficient in analyzing stocks.  Initially, an analysis was performed on DQ determining the stock's total daily volume and return for 2018.  Then an analyis of 11 additional green stocks was completed to find those with the best returns.  But the original code did not appear to be the most efficient code.  The original code was *refactored* with the idea of taking fewer steps, using less memory, and improving logic so the analysis would be more efficient.  The outcome of this project looks to determine if the refactored code *is* more efficient than the original code.

## Results
### Original Code
The *original code* contained a "nested loop", which ultimately resulted in additional steps and using more memory when determining the output of **Total Daily Volume** and **Return**.  The nested loop begins with j = 2 To RowCount and ends with Next j.  See below the original code with the nested loop.

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
    
 The outcome of the the stock performance for 2017 and execution time is as follows:
 All Stocks (2017)		
		
Ticker	Total Daily Volume	Return
AY	136,070,900	8.94%
CSIQ	310,592,800	33.07%
DQ	35,796,200	199.45%
ENPH	221,772,100	129.52%
FSLR	684,181,400	101.31%
HASI	80,949,300	25.84%
JKS	191,632,200	53.87%
RUN	267,681,300	5.55%
SEDG	206,885,200	184.47%
SPWR	782,187,000	23.07%
TERP	139,402,800	-7.21%
VSLR	109,487,900	50.00%
![image](https://user-images.githubusercontent.com/94148420/147826108-b58eb726-2d56-42da-b047-0ca9718ca2f4.png)

![Green_Stocks_2017_Original](https://user-images.githubusercontent.com/94148420/147826160-da2bc5b6-cf10-4ca2-9e08-dae63e3593c6.PNG)

And the results for 2018:

All Stocks (2018)		
		
Ticker	Total Daily Volume	Return
AY	83,079,900	-7.28%
CSIQ	200,879,900	-16.34%
DQ	107,873,900	-62.60%
ENPH	607,473,500	81.92%
FSLR	478,113,900	-39.71%
HASI	104,340,600	-20.66%
JKS	158,309,000	-60.53%
RUN	502,757,100	83.95%
SEDG	237,212,300	-7.75%
SPWR	538,024,300	-44.59%
TERP	151,434,700	-5.00%
VSLR	136,539,100	-3.54%
![image](https://user-images.githubusercontent.com/94148420/147826325-89b3c89e-285d-45e0-b04c-bd9b1197bb30.png)

![Green_Stocks_2018_Original](https://user-images.githubusercontent.com/94148420/147826351-6a832545-11b4-422e-b9d3-82291611a867.PNG)


### Refactored Code
To improve the logic and eliminate the "nested loop", 4 arrays were used:

1. Dim tickers(12) As String
2. ReDim tickerVolumes(12) As Long
3. ReDim tickerStartingPrices(12) As Single
4. ReDim tickerEndingPrices(12) As Single

The variable tickerIndex was used to match the ticker symbol via the ticker array with tickerVolumes, tickerStartingPrices, and ticker EndingPrices.  See refactored code below.

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
    Dim tickerIndex As Single
    tickerIndex = 0
    
    '1b) Create three output arrays
    ReDim tickerVolumes(12) As Long
    ReDim tickerStartingPrices(12) As Single
    ReDim tickerEndingPrices(12) As Single
       
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    tickerVolumes(i) = 0
    
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    Worksheets(yearValue).Activate
    For i = 2 To RowCount
           
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
                 
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
         
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
              
        'End If
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        'End If
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
  
## Summary
