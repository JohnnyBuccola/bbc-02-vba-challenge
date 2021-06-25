# Challenge 2 - Stock Analysis

## Purpose
The purpose of this exercise was analyze a set of stock ticker information for a handful of stocks in the green energy segment and provide some basic analytics for the years 2017 and 2018.

## Results
A Visual Basic script was used to read through all rows of data and aggregate information for each of the 12 stock tickers. On a separate sheet, the information was output and formatted.  The script was run for each year, and a screenshot was taken. Both years' aggregate results (2017 and 2018) can be seen below.

![](Resources/VBA_Challenge_2017.png)
![](Resources/VBA_Challenge_2018.png)

By looking at the data, it's clear that all stocks generally performed better in 2017 than in 2018, with the exception of one stock - `TERP` - which stayed nearly flat.  The only two stocks which saw a gain in 2018 - an overall poor year for the sector - were `ENPH` and `RUN`, which showed 2017/2018 returns of 130%/82% and 5.5%/84%, respectively.  If only one stock were to be selected from this data set for investment (in 2019), `ENPH` would be the clear frontrunner, as it's shown the most consistent returns.

Additionally, the code was refactored once to improve the processing speed.  The original code looped through the entire data set twelve times (once for each ticker), ignoring any lines that were not relevant to the "current ticker".

By refactoring so that three arrays were initialized to store data for every ticker
```    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
```
along with a separate `tickerIndex` iterator to store which array element in the above arrays would be used for the current row, the code only had to loop through the data set _once_. 

The result of the refactor was a 78% decrease in speed, from .64s to .14s. (the difference between years run was negligible).

## Summary

Code refactor may be necessary for multiple reasons.  Most frequently, it's because the program does not perform or scale well enough, or has been written in a way that is difficult to debug or manage (sometimes known as "Tech Debt"). The advantage of refactoring is that it can fix these problems. The disadvantage is that the outcome is not always apparent or valuable to the end user (especially if the refactor is not done for performance reasons).  A refactor may also require extensive re-testing even though no new features or benefits are added.  

In the VBA script for this challenge, refactoring was extremely beneficial, as the improvement was highly significant for the low amount of code change that was required. Testing the refactored code was easy as well, because the inputs and outputs were very simple.
