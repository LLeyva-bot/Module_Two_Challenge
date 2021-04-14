# Module_Two_Challenge_Green_Stocks_Analysis

## Objective

#### Perform an analysis of DAQO New Energy Corp and 11 other green energy stocks using 2017 and 2018 data.  The yearly data contains the daily total volume and daily starting and ending price of each stock.  Use VBA to create a macro to automate the calculation of the _Total Daily Volume_ and yearly _Return_ of each stock.  The macro will be used to analyze the 2017, 2018 and future yearly data. Refactor the code to make the VBA script run more efficiently.

## Results

#### In 2017, the 'DQ'(DAQO New Energy Corp) stock had the largest yearly _Return_ of 199.4%.  All but one green energy stock had a positive yearly return.  'TERP' was the one stock with a negative yearly _Return_ of -7.2% as seen in the “All Stocks (2017)” table.
![2017%20Stock%20Analysis%20Table](https://github.com/LLeyva-bot/Module_Two_Challenge/blob/main/Resources/2017%20Stock%20Analysis%20Table.PNG)
#### In 2018, the _Total Daily Volume_ increased for 8 out of the 12 analyzed stocks but only two kept a positive yearly return.  'ENPH' more than doubled the _Total Daily Volume_ to 607,473,500.  'RUN' increased its _Total Daily Volume_ to 502,757,100 and had the highest yearly _Return_ at 84%.  'DQ' drastically dropped its yearly _Return_ to -62.6% as seen in the “All Stocks (2018)” table.
![2018%20Stock%20Analysis%20Table](https://github.com/LLeyva-bot/Module_Two_Challenge/blob/main/Resources/2018%20Stock%20Analysis%20Table.PNG)

## VBA Script

#### The completed analysis can be found in ![VBA_Challenge.xlsm](https://github.com/LLeyva-bot/Module_Two_Challenge/blob/main/VBA_Challenge.xlsm). The first VBA script created is labeled "AllStocksAnalysis" and contains a nested for loop to retrieve the _Total Daily Volume_ and yearly _Return_ for each stock, one by one.  The performance time was similar for both data sets using "AllStocksAnalysis".  The 2018 data performance time was: 
![VBA_Challenge_2018.png](https://github.com/LLeyva-bot/Module_Two_Challenge/blob/main/Resources/VBA_Challenge_2018.png)
####  The first VBA script was refactored by removing the nested for loop and adding a tickerIndex to allow us to retrieve the _Total Daily Volume_ and yearly _Return_ for all stocks at one time.  The performance time improved using "AllStocksAnalysisRefactored". The 2018 data performance time was:
![BA_Challenge_2018(Refactored).png](https://github.com/LLeyva-bot/Module_Two_Challenge/blob/main/Resources/VBA_Challenge_2018(Refactored).png)
#### The code adjustments made are listed below.
1a.) Creates a ticker Index

    Dim tickerIndex As Integer
    tickerIndex = 0

1b.) Creates three output arrays

    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
2a.) Creates a for loop to initialize the tickerVolumes to zero.

    For j = 0 To 11
    
        tickerVolumes(j) = 0
        
    Next j
        
2b.) Loops over all the rows in the spreadsheet.

    For i = 2 To RowCount
    
3a.) Increase volume for current ticker

        If Cells(i, 1).Value = tickers(tickerIndex) Then
        
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        End If
        
3b.) Checks if the current row is the first row with the selected tickerIndex.
        
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
               
        End If
        
3c.) Checks if the current row is the last row with the selected ticker.  If the next row’s ticker doesn’t match, increase the tickerIndex.
        
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
3d.) Increase the tickerIndex.
 
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        
            tickerIndex = tickerIndex + 1
            
        End If
    
    Next i
    
4.) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1
        
    Next i
    
 ## Conclusion
    
 ####  The DAQO New Energy Corp stock performed well in 2017 but the yearly _Return_ dropped drastically in 2018.  Using only the 2017 and 2018 data set, the 'DQ' stock appears too volatile to recommend for investment. The 'RUN' and 'ENPH' stocks are the most recommened for investment.  Both increased _Total Daily Volume_ and continued with a positve yearly _Return_ between 2017 and 2018. The 'RUN' stock is a more preferable for investment due the yearly _Return_ increase between 2017 and 2018 from 5.5% to 84%.  The 'ENPH' stock yearly _Return_ went from 129.5% to 81.9%.
 #### Refactoring code is an important skill to improve ones programming knowledge.  It allows one to better understand their original code and improve it by making the code more universal to future data sets and/or more efficent in regards to performance time.  As shown above, the refactored VBA script improved the performance time by about .6 seconds. A disadavantage of refactoring code is it's time consuming and the time it takes to adjust may not be worth the benefit.  Although, the refactored VBA script performs at a faster speed, the performance difference is barely noticeable. 
