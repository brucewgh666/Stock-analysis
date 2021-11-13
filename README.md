# Stock-analysis


Overview of Project:

Our project is mainly focus on the performance of 12 new sustainable energy stocks in previous year 2017 and 2018. We investigate 12 stocks everyday performance and accumulate daily volume together in both 2017 and 2018 for a total daily volume to calculate the return rate of 12 stock. The purpose of this analysis is elimate the risk of bad performance stock invest choice and increase investment return rate by choosing good performance stock.

Results:

Description of stock category: 
AY: Atlantica sustainable Infrastructure
CSIQ: Canadian Solar Inc.
DQ: Daqo New Energy
ENPH: Enphase Energy Inc.
FSLR:First Solar Inc.
HASI:Hannon Armstrong Sustainable Inc.
JKS:Jinko Solar holding company Limited.
RUN:Sunrun Inc.
SEDG:Solaredge Technologies Inc.
SPWR:SunPower Corporation
TERP:TerraForm Power
VSLR:Vivint Solar.

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
        tickerVolumes(ticketIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
            
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
        'If  Then
          If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
          tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
          End If
            

            '3d Increase the tickerIndex.
           If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
           tickerIndex = tickerIndex + 1
           End If
            
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        
    Next i
![2017 performance](https://user-images.githubusercontent.com/93842672/141137623-e3df836a-fec5-4426-8bfa-4f5c183c3e69.png)
-ANALYSIS of 2017 stock performace
 
 From a big picture, the total stock performance of 2017 is remarkble except TERP:TerraForm Power have negative 7.2% . Among all the growth of 2017 new energy stock, DQ ENPH FSLR TERP have obtained over 100% growth. DQ hold the first place 199.4%. SPWR have largest total daily volume 782,187,000. DQ have smallest total daily volume 35,796,200.


-ANALYSIS of 2018 stock performace

Overview of 2018 new energy stock marekt，most of stock give negative return except ENPH(81%) AND RUN(84%) give positve return.Also,ENPH have the largest total daily volume.  Among all the negative return, DQ(-62.6%) and JKS(-60.5%) give the largest negative return rate. DQ have largest decrease return rate.


Summary:

<img width="237" alt="2017 running time" src="https://user-images.githubusercontent.com/93842672/141146624-63efdd50-1d2e-4bfe-a5a1-e447cb1ae63d.png">

<img width="235" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/93842672/141385589-4d66f3b9-57dd-4e90-8459-c3b30a6b43f2.PNG">
The advantages of refactoring code is optimize the design and structure of orginal code,but still perform the same function. Also another advantage of refactoring code is let this code can fit in more scenario.
When this refactor code apply to the original VBA script,one pros is having less processing time. It indicate it is more efficient, the program go faster.When we process extreme large database similar to this.It will cause less time to get exact same result.

