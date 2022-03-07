# Stock Analysis
## Overview of Project
The purpose of this analysis is to create a tool for a client that would like to understand the stock market with respect to return rates in the years 2017 and 2018. The goal is to create a user friendly tool that enables the user to quickly identify visualize stock selections that yield positive returns and negative returns.

## Results
Unfortunately, the refactored code did not run as intended. However, the code prior to the refactoring did run as expected. Thus, the following analysis is predicated on the results prior to the refactoring with the expection of the statements regarding run times. 

2018 Realized a downturn in the market compared to the high yields in 2017 with respect to the 12 tickers that were analyzed. In 2017 the only company to experience negative earnings was TERP. This is illustrated here: https://github.com/mshedlosky/stock_analysis2/blob/main/2017_Run_Time_Analysis.PNG. Conversely, in 2018 ENPH and RUN were the only two companies that experienced positive gains, and therefore are the only two companies that reaped positive returns across both years. The following figure depicts the performance of 2018: https://github.com/mshedlosky/stock_analysis2/blob/main/2018_Run_Time_Analysis.PNG.

The code that was utilized in the intial analysis ran much slower than the refactored code. While the output did not generate the same results, and there remains to be improvement needed to accomplish this, there is a distinct difference in run times between the two methods. The refactored code ran much quicker. The refactored code took 0.10 seconds to run the 2017 data set compared to 1.10 seconds to run the original data set. Here are images of both run times: https://github.com/mshedlosky/stock_analysis2/blob/main/VBA_Challenge_2017.png.PNG. https://github.com/mshedlosky/stock_analysis2/blob/main/2017_Run_Time_Analysis.PNG. A similar trend was seen when the 2018 data was run. The refactored code took 0.15 seconds while the original code took 1.12 seconds to run. 

In addition to more efficient run times, the refactored code does look much different. Illustrations of the original code can be seen here: https://github.com/mshedlosky/stock_analysis2/blob/main/Original_Code.PNG and https://github.com/mshedlosky/stock_analysis2/blob/main/Original_Code.2.PNG.

The refactored code can be seen as more succinct with AND statements utilized, additional variables, and arrays. The refactored code can be seen here: https://github.com/mshedlosky/stock_analysis2/blob/main/Refactored_Code.PNG and https://github.com/mshedlosky/stock_analysis2/blob/main/Original_Code.2.PNG.

## Summary

### Advantages and Disadvantages of Refactoring Code
With consideration that code is the catalyst for automation, identifying methods to enhance efficiencies and sustainability in perpetuity is critical. The advantages to refactoring code is that performance is enhanced in the way of taking up less memory and time, and efficiencies are improved by enabling updates in a more seamless manner. Ultimately code that is written today is intended to be extended upon tomorrow. Thus, continuous refactoring is essential. While continuous improvement is important, it can also be costly. The disadvantages to refactoring code is that refactoring requires time, resources, and financial commitment. Further, refactoring has the potential to create new bugs.
### Advantages and Disadvantages of the Original and Refactored VBA




