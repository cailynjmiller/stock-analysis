# Stock Analysis with VBA
## Overview of Project
Steve initially wanted to create an analysis of the performance of all stocks in the data based on the year inputted. Although we were successful in creating this analysis, Steve wanted to refactor the script to be more efficient for when he expands the dataset to include the entire stock market over the past few years.
## Results - Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.
Generally, all stocks performed better in 2017 than 2018. In 2017, almost all stocks had a positive return, except for TERP, whose value went down 7.2%. <br/>
![2017 perfomance](https://github.com/cailynjmiller/stock-analysis/blob/main/Resources/2017%20Results.png) <br/><br/>
In 2018, almost all stocks had a negative return, except for ENPH and RUN, whose values both went up over 80%. <br/>
![2018 performance](https://github.com/cailynjmiller/stock-analysis/blob/main/Resources/2018%20Results.png) <br/><br/>
The original script uses a nested loop to loop through all of tickers and the dataset to sum up the daily volumes. Since there are 12 tickers and 3013 rows in each dataset, the data is looped through 36,156(12*3013) rows of data.<br/>
    [copy of code] <br/><br/>
For both 2017 and 2018, it took approximately 0.5 seconds to loop through all of the data. <br/><br/>
2017 run-time with original script <br/>
![2017 run time](https://github.com/cailynjmiller/stock-analysis/blob/main/Resources/2017_Not_Refactored.png)<br/><br/>
2018 run-time with original script <br/>
![2018 run time](https://github.com/cailynjmiller/stock-analysis/blob/main/Resources/2018_Not_Refactored.png)<br/><br/>
In the new script, I created a variable called tickerIndex which I used in place of "i" when looping through the dataset. I initialized tickerIndex to be 0 so it first goes through the data where the first column is equal to the value of tickers(0), and at the end of the loop adds 1 to the ticker index (tickerIndex = tickerIndex +1) to then loop through the data for tickers(1) and so on. Since the tickers are in the same order as they are in the dataset, it only loops through 3013 rows of data once instead of having to go through it 12 times.
    [copy of code] <br/><br/>
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
