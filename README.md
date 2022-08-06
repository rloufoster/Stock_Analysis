# Analysis of Green Stocks Using VBA and Excel


## Objective

This project is an analysis of green energy stocks using VBA to measure performance for the years 2017 and 2018. The analysis looks at 12 green energy stocks, and its purpose is to measure trading volumes and total returns in order to assist the client in making more informed choices based on past stock performances. Although the previous code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute. To solve these issues, VBA code that was written for an earlier analysis has been refactored to produce a final product that runs more efficiently by taking fewer steps and using less memory than the initial code. The result is more flexible product that allows the client to use the analysis interactively by providing them with buttons and conditionally color-coded visuals.
 

### The Data

Initially, the client presented to me two excel worksheets with green stocks data from the years 2017 and 2018. Each worksheet contained 3012 rows of data. The stock information contained a ticker value, the date the stock was issued, the opening, closing and adjusted closing price, the highest and lowest price, and the volume of the stock. The goal of the initial analysis was to to retrieve the ticker, the total daily volume, and the return on each stock. The resulting VBA Macro was then refactored to build the code for this project.

### Method

The Module 2 VBA Macro Script was refactored so that the code was only looped through one time and all the information was collected in that one loop.

* Created tickerIndex set = to zero
* Created arrays for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices
* Used tickerIndex to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays.
* The script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes,              tickerStartingPrices, and tickerEndingPrices.
* Created working code for formatting the cells in the spreadsheet
* Created comments to explain the code
* Screenshot outputs for both stock analyses in the VBA_Challenge.xlsm along with the stock analyses for the original code.  Also included the run time pop-up messages for all four outputs.  See Results section.
  
## Results

The outputs for the 2017 and 2018 refactored analysis matched that of the initial analysis, but the Macros ran much faster and therefore, more efficiently. 

**See the 2017 and 2018 initial analysis elapsed time results below.**

![Elapsed_Time 2017_Original](https://github.com/rloufoster/Stock_Analysis/blob/main/Resources/Elapsed_Time_2017Original.png?raw=true)



![Elapsed_Time 2018 Original](https://github.com/rloufoster/Stock_Analysis/blob/main/Resources/Elapsed_Time_2018Original.png?raw=true)



**See the 1017 and 2018 refactored analysis elapsed time results below.**

![Elapsed Time 2017 Refactored](https://github.com/rloufoster/Stock_Analysis/blob/main/Resources/ELapsed_Time_2017Refactored.png?raw=true)


![Elapsed Time 2018 Refactored](https://github.com/rloufoster/Stock_Analysis/blob/main/Resources/Elapsed_Time_2018Refactored.png?raw=true)



## Summary

### The advantages of code refactoring are:

**Extensible Code:** Code refactoring makes the code more extensible for adding on many other functions.  It also helps in increasing the flexibility of the code and by this the capability of code increases.
**Maintainability:** After refactoring, the code is fresher, easier to understand or read, less complex and easier to maintain.  It also allows for it to be easily comprehended by the next programmer.

### The disadvantages of code refactoring are:

**Time Consuming:** You may have no idea how much time it may take to complete the process. It may also land you into a situation where you have not idea where to go.
**Chance of Mistakes:** In things go badly, you will introduce bugs and you will waste much more time solving the problem.

### The Advantages of Refactoring this Stock Analysis 

The biggest benefit that occurred as a result of the refactoring this particular code was the decrease in elaspsed runtime.  Since this was a short Macro, the advantage cleaner more efficient code outweighed the risk of the chance of introducing bugs or wasting a lot of time.

