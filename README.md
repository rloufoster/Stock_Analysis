# Stock_Analysis
Analyzing green stocks using VBA

## Objective

This project is an analysis of green energy stocks using VBA to measure performance for the years 2017 and 2018. The analysis looks at 12 green energy stocks, and its purpose is to measure trading volumes and total returns in order to assist the client in making more informed choices based on past stock performances. Although the previous code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute. To solve these issues, VBA code that was written for an earlier analysis has been refactored to produce a final product that runs more efficiently by taking fewer steps and using less memory than the initial code. The result is more flexible product that allows the client to use the analysis interactively by providing them with buttons and conditionally color-coded visuals.
 

### The Data

Initially, the client presented to me two excel worksheets with green stocks data from the years 2017 and 2018. Each worksheet contained 3012 rows of data. The stock information contained a ticker value, the date the stock was issued, the opening, closing and adjusted closing price, the highest and lowest price, and the volume of the stock. The goal of the initial analysis was to to retrieve the ticker, the total daily volume, and the return on each stock. The resulting VBA Macro was then refactored to build the code for this project.

### Method




## Results

The outputs for the 2017 and 2018 refactored analysis matched that of the initial analysis, but the Macros ran much faster and therefore, more efficiently. 

See the 2017 and 2018 initial analysis elapsed time results below.

![Elapsed_Time 2017_Original]()



![Elapsed_Time 2018 Original]()



See the 1017 and 2018 refactored analysis elapsed time results below.

![Elapsed Time 2017 Refactored]()



![Elapsed Time 2018 Refactored]()




## Summary

#### Pros and Cons of Refactoring Code

Refactoring helps make our code cleaner and more organized. A few advantages of a cleaner code include design and software improvement, debugging, and faster programming. It may also benefit other users who view our projects because it becomes easier to read, as it is more concise and straightforward. However, we do not always have the luxury to refactor our code due to disadvantages. These disadvantages may range from having applications that are too large to not having the proper test cases for the existing codes, which may ultimately pose some risk if we try to refactor our code.

#### The Advantages of Refactoring Stock Analysis 

The biggest benefit that occurred as a result of the refactoring was an decrease in macro run time. The original analysis took approximately one second to run, whereas our new analysis only took about a four of the time (approximately 0.25 seconds) to run. Attached below are the screenshots that indicate the run time for our new analysis.

Add png's of elapsed times here!!!