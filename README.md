# Stock Analysis with VBA
## Overview of Project
Steve initially wanted to create an analysis of the performance of all stocks in the data based on the year inputted. Although we were successful in creating this analysis, Steve wanted to refactor the script to be more efficient for when he expands the dataset. In this case, we measure the efficiency in how much time it takes for the script to run.
## Results - Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.
Generally, all stocks performed better in 2017 than 2018. In 2017, almost all stocks had a positive return, except for TERP, whose value went down 7.2%. <br/>
![2017 perfomance](https://github.com/cailynjmiller/stock-analysis/blob/main/Resources/2017%20Results.png) <br/><br/>
In 2018, almost all stocks had a negative return, except for ENPH and RUN, whose values both went up over 80%. <br/>
![2018 performance](https://github.com/cailynjmiller/stock-analysis/blob/main/Resources/2018%20Results.png) <br/><br/>
The original script uses a nested loop to loop through all of tickers and the dataset to sum up the daily volumes. Since there are 12 tickers and 3013 rows in each dataset, the data is looped through 36,156(12*3013) rows of data.<br/>
```
'3. Prepare for the analysis of tickers
'3.a Initialize variables for the starting price and ending price
    
    Dim startingPrice As Single
    Dim endingPrice As Single
    
'3.b Activate the data worksheet

    Worksheets(yearValue).Activate

'3.c Find the number of rows to loop over

    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

'4. Loop through all tickers
    
    For i = 0 To 11
    
        ticker = tickers(i)
        TotalVolume = 0

'5. Loop through the rows in the data
'5.a Find the volume for the current ticker
    
        Worksheets(yearValue).Activate
    
        For j = 2 To RowCount
            
            If Cells(j, 1).Value = ticker Then
                TotalVolume = TotalVolume + Cells(j, 8).Value
            End If

'5.b Find the starting price for the current ticker

            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                startingPrice = Cells(j, 6).Value
            End If
'5.c Find the ending price for the current ticker

            If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
                endingPrice = Cells(j, 6).Value
            End If
        
        Next j
    

'6. Output the data for the current ticker

        Worksheets("All Stocks Analysis").Activate
        
            Cells(4 + i, 1).Value = ticker
            Cells(4 + i, 2).Value = TotalVolume
            Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

Next i 
``` 
<br/><br/>

For both 2017 and 2018, it took approximately 0.5 seconds to loop through all of the data. <br/><br/>
2017 run-time with original script <br/>
![2017 run time](https://github.com/cailynjmiller/stock-analysis/blob/main/Resources/2017_Not_Refactored.png)<br/><br/>
2018 run-time with original script <br/>
![2018 run time](https://github.com/cailynjmiller/stock-analysis/blob/main/Resources/2018_Not_Refactored.png)<br/><br/>
In the new script, I created a variable called tickerIndex which I used in place of "i" when looping through the dataset. I initialized tickerIndex to be 0 so it first goes through the data where the first column is equal to the value of tickers(0), and at the end of the loop adds 1 to the ticker index (tickerIndex = tickerIndex +1) to then loop through the data for tickers(1) and so on. Since the tickers are in the same order as they are in the dataset, it only loops through 3013 rows of data once instead of having to go through it 12 times.<br/>
``` 
    '1a) Create a ticker Index
    Dim tickerIndex As Integer
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
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
         If Cells(i, 1) = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        'End If
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
            If Cells(i, 1) = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            End If
            '3d Increase the tickerIndex.
             If Cells(i, 1) = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            
        'End If
        End If
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4, 1).Value = tickers(i)
        Cells(4, 2).Value = tickerVolumes(i)
        Cells(4, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
            
    Next i 
 ```
<br/><br/>
This method was more efficient than the one used in the first script. For both 2017 and 2018, it only took approximately 0.1 seconds to perform the analysis as opposed to the 0.5 in the first script. This will be very useful when Steve expands the current dataset to include more stocks.<br/><br/>
2017 run-time with refactored script <br/>
![2017 run time rf](https://github.com/cailynjmiller/stock-analysis/blob/main/Resources/2017_Recfactored.png)<br/><br/>
2018 run-time with refactored script <br/>
![2018 run time rf](https://github.com/cailynjmiller/stock-analysis/blob/main/Resources/2018_Refactored.png)<br/><br/>
## Summary
### What are the advantage and disadvantages of refactoring the code?
One of the advantages to refactoring code is that it can make the code more efficient. At times, code can get long and repetiitive, which makes it take more time and power to run. By refactoring the code, you can create a more efficient way to get the same output.
Another advantage to refactoring the code is that it makes to code cleaner and easier to read. This helps when it comes to de-bugging and understanding how the code works.
A disadvantage to refactoring the code is the potential of breaking it. If you have not saved the previous version of the code, it can be difficult to revert it back to a working version again.
Another disadvantage to refactoring the code is that it can take a lot more effort to restructure for what can be a minor improvement.
### How do these pros and cons apply to refactoring the original VBA script?
By refactoring the VBA stock analysis script, it rain 5x faster, so it was successful in making the script more efficient. This will be especially useful when Steve adds more data to the dataset.
Refactoring this script took me a lot of time and required my full understanding of what the code was doing. Although it ran 5x fast than the original, the original still took less than a second to run. For the amount of time it took to run vs. the amount of time saved running the script, it may have not been worth it assuming the data stays the same.
