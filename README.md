# Stock Analysis: Refactor VBA Code and Measure Performance

## Overview of Project

### Purpose
The purpose of this project is to refactor the previous "All Stocks Analysis" VBA code so that Steve is able to analyze the entire stock market over the last few years. Additionally, performance will be measured to ensure the run time is improved due to the refactored code. This will allow Steve to analyze the entire data set accurately and efficiently to provide his parents with investment recommendations. 

## Results

### Analysis of Refactored Code and Performance

##### Refactored VBA Code

The intent of this project is to refactor the _All Stocks Analysis_ code in VBA so that the code will loop through all the data one time making it more efficient to run the analysis for many stocks. The first step in this process is to create a _tickerIndex_ variable and set it to equal zero before iterating over all the rows, as well as creating the three output arrays; _tickerVolumes, tickerStartingPrices and tickerEndingPrices_:
> ![Screen Shot 2021-02-07 at 11 26 59 AM](https://user-images.githubusercontent.com/77405273/107157286-1963fb80-6938-11eb-9848-0aadfbe893c6.png)

In the refactored code, the tickerIndex is used to access the stock ticker index for the tickers array and all three output arrays. This adds efficiency by capturing this information for each ticker before moving onto the next one:
> ![Screen Shot 2021-02-07 at 11 32 52 AM](https://user-images.githubusercontent.com/77405273/107157308-44e6e600-6938-11eb-99b9-5b1bf0249820.png)

The refactor code is written to loop through the stock data while reading and storing the values for _tickerVolumes, tickerStartingPrices and tickerEndingPrices_ from each row:
> ![Screen Shot 2021-02-07 at 11 33 37 AM](https://user-images.githubusercontent.com/77405273/107157320-5cbe6a00-6938-11eb-811f-41b3115522ff.png)

Formatting for the cells is included in the refactored code, as well as detailed comments to indicate the purpose of the code:
> ![Screen Shot 2021-02-07 at 11 34 22 AM](https://user-images.githubusercontent.com/77405273/107157341-7790de80-6938-11eb-8bfc-239ba086bae4.png)

To insure that the refactored code output for the 2017 and 2018 stock analysis was accurate, the results were compared to the previous outputs and did result in the same output data:
> ![Screen Shot 2021-02-07 at 11 36 09 AM](https://user-images.githubusercontent.com/77405273/107157392-c179c480-6938-11eb-8672-a5e9373b2f62.png)
> ![Screen Shot 2021-02-07 at 11 36 29 AM](https://user-images.githubusercontent.com/77405273/107157400-cb9bc300-6938-11eb-9314-b014bdc58e35.png)

##### Performance Analysis

The refactored code reduced the run time of the script, resulting in the results being compiled faster. A timer was included to capture the start and end time of the executed code. _startTime_ and _endTime_ variables were were declared as Single data types and the _startTime_ variable was set as equal to the _Timer_, which results in the clock starting after receiving the _yearValue_ input. 
> ![Screen Shot 2021-02-07 at 12 08 03 PM](https://user-images.githubusercontent.com/77405273/107158180-26371e00-693d-11eb-86ca-c40e88b0fcbc.png)

The refactored code was observed to run faster than the original Stock Analysis code. For 2017 the original run time was 0.3359375 seconds and the refactored code run time was **0.09375 seconds**.
###### Stock Analysis Timer 2017
> ![VBA_Challenge_2017](https://user-images.githubusercontent.com/77405273/107158193-364efd80-693d-11eb-902f-a12d5fa8a66b.png)

For 2018 the original run time was 0.328125 seconds and the refactored code run time was **0.09375 seconds**.
###### Stock Analysis Timer 2018
> ![VBA_Challenge_2018](https://user-images.githubusercontent.com/77405273/107158196-3949ee00-693d-11eb-8b12-1249d33fc44a.png)

## Summary

- **Advantages and Disadvantages to Refactoring Code**
  Refactoring code is a means to restructure existing code to improve its quality or efficiency, without changing the external behavior or functionality. Some advantages of refactoring code are enhancing the code to improve performance, rewriting bad code that may contain bugs and restructuring code that may have duplicates or be difficult for another person to understand. Some disadvantages of refactoring code are that it may be more time consuming that writing new code, it could introduce bugs into the code and could result in delaying a previously set deadline. 

- **Advantages and Disadvantages of the Original and Refactored VBA Script**
The advantages of the original code were that it was written clearly and functioned as expected, delivering accurate results on the 12 stocks included. Disadvantages of the original code were that it was not written as efficiently as it could have been, resulting in longer run times, and might not have worked if used to analyze a much higher number of stocks at once. 
The advantages of the Refactored VBA script are that it utilizes the _tickerIndex_ to capture the information for each output value as it works through each ticker. This resulted in faster run times and the ability to use this script to run an analysis on many stocks at once. Disadvantages of the refactored code are that it took additional time to write and through the process, errors were introduced and had to be corrected. For example, I initially included the three output arrays incorrectly, resulting in code that did not work. 
