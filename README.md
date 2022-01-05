# stock-analysis
# Refracting the code for Green stocks Data
## Overview of the project:
### Back ground
#### Steve's parents seek  his help to plan their investement in Green energy stocks. Steves parents are passionate about Green energy and believe alternative energy production industry to have high demand in the future.Steve's parents have decided to invest their money in the DAQO (DQ) New Energy Corp, a company that makes component for solar panelsSteve wants to help his parents decision of diversifying their funds by looking deep in to  DAQO's stocks in the past years and also look in to additional Green energy stocks 
#### To do this, he created an excel file with all the necessary data like daily volume of stocks,, opening and closing pice and the highest and lowest price of the dayHe had the data for 2 years - 2017 and 2018. Our goal is to help steve with analysis and use VBA code to  create some macros that steve could use to automate analysis for this data and additional data or any other stocks in the future.


### Initial Analysis
#### Initial analysis was done  using VBA code to anlyze the Total stock volume for DAQO and we also calculated the percentage return. The purpose of this analysis is to refact the written VBA code to make it more efficient and clean


## Results

### Goals of Analysis
First goal was to know from the user the year they want to run the analysis on.
This was done by creating an input box. 

code:yearValue = InputBox("What year would you like to run the analysis on?")
Second goal was to create the start time to determine the run time of the code. So the timer was turned on and assigned to start time variable

**code: startTime = Timer**

Similary timer was turned on at the end of the code abd assigned to a end time variable

**code: endTime = Timer**

The analysis was recorded in the new work sheet All Stocks Analysis. For the analysis, we want to know the total volume of stocks for each ticker (company). We also want to calculate the total percentage return for each companies. As we have to compile these data for each companies/tickers, all the tickers were inttialized in an array as string 

code:

**Dim tickers(12) As String**
    
    **tickers(0) = "AY"**

    **tickers(1) = "CSIQ"**

    **tickers(2) = "DQ"**

    **tickers(3) = "ENPH"**

    **tickers(4) = "FSLR"**

    **tickers(5) = "HASI"**

    **tickers(6) = "JKS"**

    **tickers(7) = "RUN"**

    **tickers(8) = "SEDG"**

    **tickers(9) = "SPWR"**

    **tickers(10) = "TERP"**

    **tickers(11) = "VSLR"**

### Calculation of total volume
The goal is to calculate total vloume for each company/ticker. So we have to loop throught the ticker array by assigning a  variable to the index of the array.

**For j = 0 To 11**

**ticker = tickers(j)**

Total volume was calculated by looping through the rows of each ticker in the data work sheet of the year user entered in the input box

code:
 **Worksheets(yearValue).Activate**

so first step is to activate that work sheet . Next we have to lopp through each row from the second row to the end of the row using a for statement 

**For i = 2 To RowCount**

Our goal during the looping is to collect Total daily volume, starting price and the ending price. Starting and ending proce is needed to calculate the percentage return using the formula given below

**endingPrice / startingPrice - 1**

After collecting the values, some formatting was done to make the data look clear

**'Create bottom border for  A3 to C3 cells**

**Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous**    
**Range("B4:B15").NumberFormat = "#,##0"**    
**Range("C4:C15").NumberFormat = "0.0%"**

Additional formatting was done to add highlight color for eprcentage returns that will help has visualize the results more clearly. Postive return percentages were color coded with green abd rest all was color coded red


### 2017 Stock Analysis results
- Over all 2017 looks like a great year for most of these companies in the energy industry.
- All companies had a positive percentage return except TERP
- Values highlighted in Green shows positive percentage return
- The total daily volume of stocks varied for all these copmpanies
- Analysis showed that DAQO had low Total daily stock volume  in 2017 that other companies
- But DAQO had the highest perecntage return (199.4%) compared to other companies in 2017

### 2018 Stock Analysis results
- 2018 analysis looks gloomy for most of these companies in the energy industry.
- All companies had a negative percentage return except couple of them (ENPH and RUN)
- The total daily volume of stocks was higher compared to 2017 for few of them
- Analysis showed that DAQO had very high Total daily stock volume  in 2018 than 2017
- But DAQO had negative  perecntage return (- 62.6% )  in 2018
- Many other that did well in 2017 had negative return in 2018
- Analysis shows that many energy industries were hit hard in 2018

###  Execution time of the refactored script

https://github.com/rajimuth/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png
![VBA_Challenge_2017.png](https://github.com/rajimuth/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)


- The refracted code ran 0.75 seconds and the original script ran for 0.78 seconds
- So refraction of the code definitely gave an efficient run time

## Summary
## Advanatages and Disadvantages of Refactoring a code
**Advantages**
**Refactoring improves the Design of Existing Code - Martin Fowler.** 
- Improve the clarity of the code or in other words clean code
- Remove Bugs in the original code
- Improve run time
- Remove repeating codes or excessive parameters
- Remove code duplications
- helps in adding more funtionality to the code

References:

https://www.ionos.ca/digitalguide/websites/web-development/what-is-refactoring/

https://www.quora.com/What-are-the-pros-and-cons-of-refactoring

**Disadvantages**
- There is a chance to introduce new bugs in the process
- Requires high coordination effort when a large team is involved in the process
- Limited opportunties to add new function as it could interefre with the architechture/strutcture of the original code
- Refactoring without proper goals and clarity of the project and lack of coordination can lead to added issues/errors
