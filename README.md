# Stock Analysis

## Overview of the Project
The purpose of this project is to help Steve conduct an analysis of all the stocks in this workbook with the click of a button. It is important to specifically focus building a code that can handle a much larger data set and on refactoring the code to make sure it is running efficiently. 

## Results
### Stock Preformance
Based on the stocks in this workbook, the market preformed worse in 2018 than they did in 2017. 2017 seemed to be a year with high growth, except for stock TERP. On the other hand, 2018 seemed to be a year with high losses, except for stocks ENPH and RUN. The stock ENPH is the only stock that had extremely high growth pattern both years, increasing 129.5% in 2017 and 81.9% in 2018. The stock TERP is the only stock that had a constant loss pattern in both years, losing 7.2% in 2017 and 5% in 2018. For more details on the stock performance, see the two tables below: 
![Stock_Performance_2017](https://github.com/andreabassetti/stock-analysis/blob/main/png/Stock_Performance_2017.png)
![Stock_Performance_2018](https://github.com/andreabassetti/stock-analysis/blob/main/png/Stock_Performance_2018.png)
### Code Refactoring
In order to ensure that the code could run with a large dataset and efficiently, the code we built in the module needed to be refactored. To check that the execution time was faster i executed a timer function. I started the timer with `startTime = Timer` after the `InputBox` code line so that it would not count the time it took Steve to chose the yearValue. I ended the timer with `endTime = Timer` after the whole code ran, but before the `MsgBox` code line. Below is the series of screenshots from the execution times, in both cases you can see that the times were slightly reduced when the code was refactored.

|    | First Code |  Refactored Code  |
| ----------- | ----------- | -----------  |
| 2017 | ![VBA_Module_2017](https://github.com/andreabassetti/stock-analysis/blob/main/png/VBA_Module_2017.png) |  ![VBA_Challenge_2017](https://github.com/andreabassetti/stock-analysis/blob/main/png/VBA_Challenge_2017.png) |
| 2018 | ![VBA_Module_2018](https://github.com/andreabassetti/stock-analysis/blob/main/png/VBA_Module_2018.png) |  ![VBA_Challenge_2018](https://github.com/andreabassetti/stock-analysis/blob/main/png/VBA_Challenge_2018.png) |


## Summary


Overview of Project: Explain the purpose of this analysis.
Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.
Summary: In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?
