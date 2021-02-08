# Stock-Analysis

## Project Overview
My client, Steve, needed help with his analysis of what stocks were worth investing in.  Although our client is well versed in the use of Excel, it was determined that using VBA and automating the anaylsis process would better serve his purposes.

## Resourses
 - Data Source: [Green_stocks.xlsm](https://github.com/stephenanayashilliard/Stock-Analysis/blob/master/green_stocks.xlsm)
 - Excel
 - VBA (Visual Basic)

## Results
When I began too code for my client's project, I initially started out writing a simple if/then statement so that my client could run an analysis specifically analysing DAQO stocks based on the year, their total daily volume and the stock's annual return.

#### DQ Analysis
   
![DQ_Analysis](https://github.com/stephenanayashilliard/Stock-Analysis/blob/master/Resources/DQ%20Analyis.png)

The analysis revealed that "DAQO" had been performing poorly over the last year. The client requested the ability to analyse all stocks over multiple years. To accomplish this  the program refered to as "AllstockAnalysis" was created.

#### Sub AllStocksAnalyis

![All_Stocks_Analysis](https://github.com/stephenanayashilliard/Stock-Analysis/blob/master/Resources/AllstocksAnalyis%201.png)

To further aid the client's ability to analyse the data easily, further coding was done to allow for formatting the data.

#### Formatting the Data

![formatting](https://github.com/stephenanayashilliard/Stock-Analysis/blob/master/Resources/formatting.png)

#### Timer

Because the client is often working with financial clients in face to face meetings, the client requested the ability to be able to see how fast the program was producing the desired data. The following subroutine was added. 

![Timer](https://github.com/stephenanayashilliard/Stock-Analysis/blob/master/Resources/AllstocksAnalyisandtimer.png)

The program measuring the results was created with the following outcomes.

##### Report Time for 2017
![Original Report Time 2017](https://github.com/stephenanayashilliard/Stock-Analysis/blob/master/Resources/Greenstock%202017.png)

##### Report Time for 2018
![Original Report Time 2018](https://github.com/stephenanayashilliard/Stock-Analysis/blob/master/Resources/Greenstock%202018.png)

As you can see the run times for 2017 and 2018 were .484375 seconds and .5859375 seconds respectively.  The client then asked if it would be possible to have the formulas report their results even faster.  This is especially important as more stock's information will be added in the future.  To accomplice this,  the code needed to be refactored by switching the nesting order of the loops and using arrays.  

#### The Refactored Code

![Refactored_Code](https://github.com/stephenanayashilliard/Stock-Analysis/blob/master/Resources/Refactored%20code.png)  

##### The Refactored Run time for 2017
![Refactered Run Time 2017](https://github.com/stephenanayashilliard/Stock-Analysis/blob/master/Resources/VBA_Challenge_2017.png)

##### The Refactored Run time for 2018
![Refacterd Run Time 2018](https://github.com/stephenanayashilliard/Stock-Analysis/blob/master/Resources/VBA_Challenge_2018.png)

The new runtimes were .2109375 seconds for 2017 and .0859375  seconds for 2018.

## Summary
The final project did provide the automated analysis the client requested and the refactored code did decrease the amount of time it took to run each analysis.  In hindsight, the amount of time required to refactor the code was disproportionly greater than what was actually shaved in seconds off of the time needed to run the original report.  

### Advantages and Disadvantages of Refactoring Code
Refactoring code has several potential advantages including the ease at which one can read the code as well as the ability to make faster changes within the code if needed. There are also disadvantages when refactoring such as running the risk of introducing bugs into the original coding.  This is especially dangerous if one is under a deadline and does not have the critical time to test the refractured programming. In the case of this project, refactoring the code introduced several bugs that had to be repaired.

### Comparison of the Advantages and Disadvantages between the Original and Refactored VBA Script.
In the case of my "green_stocks" and the refractured programing, although the refractured program is much easier to read from a programing stance, the running time saved was not worth the extra time put into refracturing as well as the risk of destablelizing stable code.  In conclusion, refracting makes sense when the amount of data that needs to be analysed is vast;  you need to have the ability to make quick changes to the program in the future; and/or the code needs to be more easily understood.  
