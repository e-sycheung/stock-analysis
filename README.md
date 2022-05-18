# Stock Analysis

## Project Overview
Create a script in VBA to iterated through a stock data to analyze for to assist the client and his parents to make better investments. Refactor the program to ensure that it can timely analyze larger data sets. 

### Process:
1. Create variables for ticker elements. 
2. Assign variables are the proper data types. 
3. Create a for loop and if/then statement to iterate over all the rows to calculate each ticker’s “Total Daily Volume” and “Return”
4. Format the table to emphasize positive and negative returns. 

## Resources
- Data Source: green_stocks.csv
- Software: VBA Microsoft Excel 365

## Summary
### Results:

Initially the timing output before a refactoring process was the following: 

<img src=" " style=" width: 385px; height: 250px">


Although the program gave the desired output, the client wished to apply it to a larger data set and raised concerns about the execution time for different systems.

<img src=" " style=" width: 385px; height: 250px">


For the refactoring process, separate “for loops” were created for each task. The nested “for loops” in the initial program created an elongated cascading step-process to be complete before iterating over the next row of data.

Refactored Steps:
- Initializing tickerVolume variable to 0.
- Indexing the tickers. 
- Output calculated values.

### Analysis:

Refactor or Not to Refactor?

The final script accomplished the goals of the project. Although refactoring created efficiency, not every situation benefits from this process. In large integrated system, refactoring and improving one area can create adverse effect in another and break the system. Sometime the original script is needed so that the many cogwheels of the entire machine can run properly and effectively.
For the stock analysis on the new data set, the new data should be evaluated to ensure that format could still be processed by the script changes. The original and refactored script can be tested and edited to reveal the best approach for the desired output. 



