# Stock-analysis for VBA

## Overview: VBA Stock Analys Project

### Steve has created a nice VBA code to help analyse stock performance. Since the code was tested on few stocks, we were tasked with
refactoring the code in order to improve its efficiency, by taking fewer steps, using less memory and improving the logic of the code. 

## Results: Refactoring the VBA Code and Measuring Performance

### The refactor code showed improved efficiency, which was evident by the reduced length of time it took to complete. 

add pics of time

1. We began by creating a 'tickerIndex' variable and setting it to zero. We used this 'tickerIndex' variable to access the tickers array
and the three output arrays.

'''
    For i = 0 To 11
       tickerIndex = tickers(i)
'''

2. We created three output arrays: 'tickerVolumes', 'tickerStartingPrices', and 'tickerEndingPrices' to hold the output in our file. 

'''
    Dim tickerVolumes As Long
    Dim tickerStartingPrices As Double
    Dim tickerEndingPrices As Double
'''

3. We then utilized 'For' loops to initialize the 'tickerVolumes' to zero and loop over all of the raw data in the spreadsheet.

'''
Worksheets(yearValue).Activate
       tickerVolumes = 0
           
        
    ''2b) Loop over all the rows in the spreadsheet.
        For j = 2 To RowCount
'''

4. And then in the 'For' loop, we collected our 'tickerVolumes', 'tickerStartingPrices' and 'tickerEndingPrices' via our 'If-Then' statement to store 
for our output data. 

'''
If Cells(j, 1).Value = tickerIndex Then
            tickerVolumes = tickerVolumes + Cells(j, 8).Value
            End If
                      
        If Cells(j - 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then
            tickerStartingPrices = Cells(j, 6).Value
            End If
         
        If Cells(j + 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then
            tickerEndingPrices = Cells(j, 6).Value
            End If
'''

5. And finally, we looped through the arrays to output the Ticker, Total Daily Volume and Return.

'''
Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickerIndex
        Cells(4 + i, 2).Value = tickerVolumes
        Cells(4 + i, 3).Value = (tickerEndingPrices / tickerStartingPrices) - 1
'''

In the end the results of the data from the original code and the refactored code remained the same, demonstrating that we did not lose any 
data integrity in this process. 

## Summary

In conclusion, refactoring code can provide an improvement over the original product which can provide some advantages yet also some
disadvantages. The advantages include a cleaner script that is easier to read by others, logical errors can be more easily recognized and patterns
can emerge which help identify what can be done in the future. It also removes redundancies and create some re-usable code for future projects. 

The disadvantages of refactoring can range from being time intensive (particularly on a tight schedule), it could introduce unwanted bugs which
reduce performance and it may not yield any significant advantages to the end user. 

The advantages of this refactored code is that it loops over the Raw Data only once to collect all of the data for the output via the 
tickerIndex, which reduces the completion time. The disadvantage is that the additional code is required to ensure this occurs.  

###
