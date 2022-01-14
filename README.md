# Green Stock-analysis
# Project Overview

## Background
Steve asked me to analysis a green stock for his parents for investment decision. In order to do, I used the Visual Basic Application (VBA) in Excel to find out each stocks annual volume and return. Steve able to analyze each stock at just click of a button and able to know what the best option for his parents would be.
### Purpose
The purpose of this project was to make an efficient way to look at multiple stocks and analyze them by using VBA coding. To ensure that I did right, I compared the new execution time with the original workbook.
## Results 
### Refactoring the Code
In order to make my code more efficient, I switched the nesting order to for loops. I created three output arrays: ticker Volumes, tickerStartingPrices, and tickerEndingPrices. The ticker array was used to establish the ticker symbol of a stock. I matched the other arrays with tickers array by using the tickerIndex variable.By doing this way, the analysis would be completed much faster than using the nested for loops for earlier. 
### Run time and 2017 vs 2018 Stock Performance
There is a huge difference in the 2017 vs 2018 performance og Green Stocks. And based on the run-times, the refactored code runs 0.4 seconds faster than the original code. Here below see the run times of both year.
![image](https://user-images.githubusercontent.com/96403349/149468692-817fc264-dd94-434d-9977-ff7bb5255a0f.png)
![image](https://user-images.githubusercontent.com/96403349/149469457-4c40cdbf-5680-4c18-b2f6-61384acaa1a3.png)
## Summary
### Advantages and Disadvantages of refactoring code
The major advantage of refactorging code is making the code more efficient. The disadvantage is errors may be occur and destory an original working code. It is always smart to save original code before refactoring code just incase you end up not being able to refactor the code.
### Advantages and Disadvantages of refactored VBA Script
The mojar advantage of refractoring code in VBA script is that we can always use original code as we want to and can put new code with our old code using different modules. The disadvantage is that if we have not understanding of the syntax, we'll struggle to refactor our code as the syntax matters specialy when trying to make code more efficient. 


