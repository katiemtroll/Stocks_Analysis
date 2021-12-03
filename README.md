# Stock Analysis

## Overview of Project

This workbook was created to help a stock trader evaluate the best stocks to invest in. 

### Purpose

Steve will utilize this Excel workbook to determine which of 12 different stocks would be the best option for his parents to invest in. Steve’s parents initially wanted to invest in DAQO (DQ) stock. 

## Results

### DQ Analysis
Steve’s parents initially wanted to invest in DAQO (DQ) stock. An analysis was conducted on DQ’s 2018 stock to determine if the annual return indicated that it would be a good investment. The analysis showed that in 2018, DQ had a negative return of -62.6%. With the low return, Steve decided he wanted an analysis of 12 stocks over two years to help his parents select a better option.

### All Stocks Analysis
The first iteration of analyzing the 12 stocks for the years 2017 and 2018 provided promising results for Steve to choose a stock to invest in (see analysis results below). 

![This is an image]()

However, as Steve began to think of future uses for this workbook, he determined that total run time of the code needed to be as low as possible to maximize time. It was decided that the code should be refactored to make the analyses more efficient.

### Refactored All Stocks Analysis
The initial iteration of analyses ran in 0.2773 seconds and 0.2734 seconds in 2017 and 2018, respectively.

**insert run times

####Refactoring
The code was refactored in a couple of places to increase efficiencies in the run time. 
1. A variable called `TickerIndex` was created and set to zero before looping through the data. This variable was used to access the correct index within each of the arrays in the code.

**insert tickerindex screenshot

2. Arrays were created for `tickers`, `TickerVolumes`, `TickerStartingPrices`, and `TickerEndingPrices`. The output arrays provided a simpler, cleaner solution for analysis output. `TickerIndex` was used to access the stock ticker index for each of these arrays. The code looped through the stock data utilizing the arrays and output the results to the “All Stocks Analysis” tab.

**insert array for loops 

## Summary & Conclusion

After the code was refactored, it can be noted that the run time is faster at 0.078125 seconds and 0.078125 seconds for 2017 and 2018, respectively. 




In general, key advantages of refactoring the code are increased quality and readability, and the code should run more efficiently. The risk of refactoring code is that you could break the code if you refactor incorrectly, or if you don’t have a strong understanding of the syntax of the language. In these instances, it may be better not to refactor the code, or to make sure that the original code is saved and changes are made little by little.

It is difficult to say if the time required to refactor this VBA code was justified by the changes in output. The change in run times was minimal between the initial and refactored codes, and the initial code may have likely been acceptable as-is. However, because we refactored the code, it should be much easier to return to the code to make future changes because of the improved structure and quality of the code.
