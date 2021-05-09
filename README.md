# RefactoringStockAnalysis
Regina Negrycz
genglist@yahoo.com
Module 1 Challenge
Submitted 8 May 2021
Due 9 May 2021
VBA_Challenge.xlsm
VBA_Challenge_2017.png
VBA_Challenge_2018.png
README.md

## Overview of Project
I had analyzed twelve stocks for Steve in the past.  Steve would like to do more research for his parents.  He asked that I expand my analysis to review stocks over the past two years.  To do this, I will need to make my code more efficient.

### Purpose
The purpose of this challenge was to enhance existing code to run more efficiently and faster.  The refactored code should take fewer steps, use less memory and/or improve the logic of the code making it easier to be read by others.   

## The Data
The information provided for each stock includes the opening value, high value, low value, closing value, adjusted closing value and the volume for each day that stock was sold on the market.  The goal was to loop through each ticker's volumes, starting prices and ending prices to determine the total daily value and the return. 

## Results

#Analysis
I created nested for loops to loop through all the rows of the spreadsheet analyzing the starting and ending rows for each ticker.  The output arrays of tickerVolumes, tickerStartingPrices and tickerEndingPrices using the variable tickerIndex allowed me to output the ticker, total daily volume and return columns on the spreadsheet.

##Thoughts on Refactoring
Refactoring cleans up the code, making it more organized and allowing it to run faster.  It is easier to read and more concise.   While doing so, the programmer is altering code that originally worked and could make syntax errors that would deem the code unusable.







