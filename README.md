# Stock Analysis: Refactor VBA Code and Measure Performance

## Overview of Project

### Purpose
The purpose of this project is to refactor the previous "All Stocks Analysis" VBA code so that Steve is able to analyze the entire stock market over the last few years. Additionally, performance will be measured to ensure the run time is improved due to the refactored code. This will allow Steve to analyze the entire data set accurately and effecietnly to provide his parents with investment recommendations. 

## Results

### Analysis of Reafactored Code and Performance

#####Refactored VBA Code

The intent of this project was to refactor the _All Stocks Analysis_ code in VBA so that the code would loop through all the data one time making it more effeciant to run the analysis on many stocks. The first step in this process was to creat a _tickerIndex_ varialble and set it to equal zero before iterating over all the rows, as well as creating the three output arrays; _tickerVolumes, tickerStartingPrices and tickerEndingPrices_:
> ![Screen Shot 2021-02-07 at 11 26 59 AM](https://user-images.githubusercontent.com/77405273/107157286-1963fb80-6938-11eb-9848-0aadfbe893c6.png)

In the refactored code, the tickerIndex was use to access the stock ticker index for the tickers array and all three output arrays. This adds effeciency by capturing this information for each ticker before moving onto the next one:
> ![Screen Shot 2021-02-07 at 11 32 52 AM](https://user-images.githubusercontent.com/77405273/107157308-44e6e600-6938-11eb-99b9-5b1bf0249820.png)

The refactor code is written to loop through the stock data while reading and storing the values for _tickerVolumes, tickerStartingPrices and tickerEndingPrices_ from each row:
> ![Screen Shot 2021-02-07 at 11 33 37 AM](https://user-images.githubusercontent.com/77405273/107157320-5cbe6a00-6938-11eb-811f-41b3115522ff.png)

Formatting for the cells is included in the refactored code, as well as detailed comments to indicate the purpose of the code:
> ![Screen Shot 2021-02-07 at 11 34 22 AM](https://user-images.githubusercontent.com/77405273/107157341-7790de80-6938-11eb-8bfc-239ba086bae4.png)

To insure that the refactored code output for the 2017 and 2018 stock analysis was accurate, the rsults were compared to the previous outputs and did result in the same putput data:
> **2017** ![Screen Shot 2021-02-07 at 11 36 09 AM](https://user-images.githubusercontent.com/77405273/107157392-c179c480-6938-11eb-8672-a5e9373b2f62.png)
> **2018** ![Screen Shot 2021-02-07 at 11 36 29 AM](https://user-images.githubusercontent.com/77405273/107157400-cb9bc300-6938-11eb-9314-b014bdc58e35.png)

#####Performance Analysis


#### Stock Analysis Timer 2017
image

#### Stock Analysis Timer 2018
image

## Summary

- **Advantages and Disadvantages to Refactoring Code**

- **Advantages and Disadvantages of the Origianl and Refactored VBA Script**
