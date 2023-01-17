# Stock Analysis
Performing analysis on stocks to uncover trends

# VBA of Wallstreet 

## Overview of Project
The goal of this project was to analyze stock data from 2017 and 2018 to help Steve find a stock for his parents to invest in. To find a stock worth investing in, we calculated two measures for 12 individual stocks. First, we calculated the Total Daily Volume of the 12 stocks that we analyzed. Then, we found the return of each stock for the years 2017 and 2018. 

## Results 

### Ticker Index 
Once our data output sheet was formatted and our array of tickers was defined, we created a Ticker Index to make the analysis work for all tickers in our dataset. For this, we needed to initialize a for loop that would loop through all of our tickers one by one. Once tickers were established as our variable i, our For loop recognized that it would need to analyze each ticker. Once this was complete, we had to define the variables that we would be looking for through our datset: Total Volume, Starting Price and Ending Price. Our last step before looping through the rows was to set the Total Volume to 0, so that the volume count would start back at 0 for each ticker.
![This is an image](https://github.com/nicole-tough/stock-analysis/blob/main/TickerIndex.PNG)

### For Loop and If Statements
To loop through each row in the data set, we had to create a nested loop that would look through each row of data for each ticker. We calculated the Total Volume of each stock by setting our loop to look at the ticker in column 1, and then counting the data in column 8 if the ticker matched the one we were looking for. 

Finding the Starting Price and Ending Price of each stock required conditionals. To find the Starting Price, we needed to set an 'If' statement that found if the ticker matched the one we were looking for AND that the row before it did not match the ticker we were looking for. The find the Ending Price, our 'If' statement found it the ticker matched one we were looking for AND that the row after it did not match the ticker we were looking for.
![This is an image](https://github.com/nicole-tough/stock-analysis/blob/main/ForLoop.PNG)

### Outputs
To finish our For Loop, we needed to format a place in our workbook for the outputs to go to. We formatted the Total Volume of each stock to populate in one column of our results sheet. We also wanted the return from each stock, so we calculated that by using the following formula: Ending Price / Starting Price - 1 and putting that output into another column of our results sheet. 
![This is an image](https://github.com/nicole-tough/stock-analysis/blob/main/Outputs.PNG)

### Formatting 
Once the spreadsheet was coded wtih VBA to analyze each stock, we wanted the results of our analysis to be easy to read. In addition to some other minor formatting, we formatted the cells to be green if the return was positive and red if the return was negative.
![This is an image](https://github.com/nicole-tough/stock-analysis/blob/main/Formatting.PNG)

### Timer
The last step in this analysis was to add a timer to the code to make sure it would run quickly. To set a timer, we initialized a timer at the beginning of the subroutine and ended the timer at the end of it. We called a Message Box at the end to show the time it took for the system to perform the subroutine. 
![This is an image](https://github.com/nicole-tough/stock-analysis/blob/main/StartTimer.PNG)
![This is an image](https://github.com/nicole-tough/stock-analysis/blob/main/EndTimer.PNG)

### Stock Analysis Results
Our results showed that the only stocks that had a positive return in both 2017 and 2018 were stocks with tickers ENPH and RUN. The full results can be seen below.
![This is an image](https://github.com/nicole-tough/stock-analysis/blob/main/StockAnalysisResults_2017.PNG)
![This is an image](https://github.com/nicole-tough/stock-analysis/blob/main/StockAnalysisResults_2018.PNG)

## Summary

### Advantages and Disadvantages of Refactoring
One advantage of refactoring is that you don't have to create the code from scratch. If the code is already close to accomplishing the goal of the coder, then there is a framework already laid out for you. However, that leads to a disvantage which is that if the code is unfamiliar to you, it may take some time for you to understand what the original coder was doing. Another disadvantage of refactoring is that the coder before you may have made a mistake, and it may be hard for you to find that mistake within the previously written code. 

### Original VS Refactored VBA Script
The main advantage of the refactored VBA script is that it runs faster than the old VBA script. Part of the reason that the original script ran slower is because it had an additional subroutine that calculated the Total Volume and Return of just one stock and output those results into a separate spreadsheet. This would be an advantage if someone wanted to analyze that one stock in particular, but in the case of speed an efficiency it is a disadvantage. 
