# Stock Analysis 

## Overview

### The Purpose of the Project

- The purpose of the project is to analyze stock market performance over the last few years for Steve to help his parents make sound investment decisions.
- Another objective is to refactor VBA code and measure performance of the refactored code through runtimes.

## Results 

### VBA Code and Workbook

- The VBA Code and workbook **before refactoring** to analyze stock performance and runtime for the year entered by the user can be found in the links [VBA_challenge - Before Refactoring.vbs]()
[VBA_challenge - Before Refactoring.xlsm]()

- The VBA Code and workbook **after refactoring** to analyze stock performance and runtime for the year entered by the user can be found in the links [VBA_challenge.vbs]()
[VBA_challenge.xlsm]()


### Stock performance and execution time before refactoring for 2018
 ![Stock Performance 2018]()
 
 ![Execution Time 2018]()

### Stock performance and execution time after refactoring for 2018
  ![Stock Performance 2018]()
  
  ![Execution Time 2018]()

## Summary

### Advantages of Refactoring Code
- Refactoring can make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. 

### Disadvantages of Refactoring Code
- If refactoring is not done properly, then it may introduce new bugs and errors into the code.
- Refactoring can be time consuming if the code is long and complicated.

### Pros of the refactored VBA script
- The refactored code showed an improvement in runtime (from 0.8s to 0.1s) for the year 2018 while providing the same results as the original code.
- The refactored code will run perfectly well if other years are added to the dataset for the same stocks, without affecting the performance significantly.

### Cons of the refactored VBA script
- If more stocks are introduced in the dataset, it will require a change in the code to include those additional stocks in the 'tickers' array. If the number of new stocks is too high, then the code will need to be modified to dynamically read those new stock tickers from the dataset.
