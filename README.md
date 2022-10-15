# Stock Analysis 

## Overview

### The Purpose of the Project

- The purpose of the project is to analyze stock market performance over the last few years for Steve to help his parents make sound investment decisions.
- Another objective is to refactor VBA code and measure performance of the refactored code through runtimes.

## Results 

### VBA Code and Workbook

- The VBA Code to analyze stock performance and runtime for the year entered by the user can be found in the link [VBA_challenge.vbs](https://github.com/manasidek/stock_analysis/blob/main/VBA_challenge.vbs)

- The workbook containing the above VBA code and showing the stock performance for the selected year is in the .xlsm file [VBA_challenge.xlsm](https://github.com/manasidek/stock_analysis/blob/main/VBA_Challenge.xlsm)

### Stock performance and execution time for 2017
 ![Stock Performance 2017](https://github.com/manasidek/stock_analysis/blob/main/Resources/All%20Stocks%202017.png)
 
 ![Execution Time 2017](https://github.com/manasidek/stock_analysis/blob/main/Resources/VBA_Challenge_2017.png)

### Stock performance and execution time for 2018
  ![Stock Performance 2018](https://github.com/manasidek/stock_analysis/blob/main/Resources/All%20Stocks%202018.png)
  
  ![Execution Time 2018](https://github.com/manasidek/stock_analysis/blob/main/Resources/VBA_Challenge_2018.png)

## Summary

### Advantages of Refactoring Code
- Refactoring can make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. 

### Disadvantages of Refactoring Code
- If refactoring is not done properly, then it may introduce new bugs and errors into the code.
- Refactoring can be time consuming if the code is long and complicated.

### Pros of the refactored VBA script
- The refactored code will run perfectly well if other years are added to the dataset for the same stocks, without affecting the performance significantly.

### Cons of the refactored VBA script
- If more stocks are introduced in the dataset, it will require a change in the code to include those additional stocks in the 'tickers' array. If the number of new stocks is too high, then the code will need to be modified to dynamically read those new stock tickers from the dataset.
