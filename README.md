# Stock Analysis using VBA

## 1. Overview of Project
This challenge employs the use of Excel VBA to asses the total volumes and returns of certain stocks in the years 2017 and 2018. Our stakeholders are Steve's partents who would like to use a user-friendly medium to do this task. The analysis will aslo detail how the program was improved through refactoring. 

## 2. Purpose
  1) To provide Steve's parrents with Total Daily Volumes and Anual returns for the years 2017 and 2018
  2) Developing Excel into a user-friendly medium by including a Run stock Analysis button and a Clear data button
  3) To arrive at the most time efficient by exporign the ideiology of Refactoring

## 3. Result of Data Analysis

### 3.1. Year 2017

All the targeted stocks faired well in the year 2017 except for the ticker "TERP". The maximum being "DQ" with nearly 200% return proved to be a good year for those who may have invested in these stocks. The image included in this seciton details the output of individual tickers for out stakeholeder's perusal.

_Image of analysis for the Year 2017 along with time taken to execute code_

![](https://github.com/AllenAx91/stock-analysis/blob/main/RESOURCES/VBA_Challenge_2017.png)

### 3.2. Year 2018

With the exception of "ENPH" and "RUN", all other stocks failed in the year 2018. "TERP" which brought good returns in 2017 went down to -5.0%. Seteve's partents were interested in "DQ" which too did not do well this year. 

_Image of analysis for the Year 2018 along with time taken to execute code_

![](https://github.com/AllenAx91/stock-analysis/blob/main/RESOURCES/VBA_Challenge_2018.png)

## 4. Run time optimization

Two sets of VBA programms were written to explore which method is more time efficient. The first one was with a nested for loop which resulted in running lines equal to 12 * the number of row. On the other hand, the refactored code used Arrays to omit the need of iterating through Tickers. 

### 4.1. Code with nested For loop

As mentioned in the previous section, the nested for loop involved running a lot of lines. If there were a lot more Stock data available in teh worskeets 2017 and 2018, the nested for loops would have taken a longer time to give results. 

_Screen shot of Run times for the Years 2017 and 2018 using **nested for loop**_

![](https://github.com/AllenAx91/stock-analysis/blob/main/RESOURCES/NestedFor_2017.png) ![](https://github.com/AllenAx91/stock-analysis/blob/main/RESOURCES/NestedFor_2018.png) 

### 4.2. Refactored code with Arrays

The number of lines run in the code was cut down since there was only one for loop used. This time, the code ran only once through all the rows in the worksheet while simultaneously storing the values of Starting Price and Ending price. 

_Screen shot of Run times for the Years 2017 and 2018 using **arrays**_
![](https://github.com/AllenAx91/stock-analysis/blob/main/RESOURCES/Refactored_2017.png) ![](https://github.com/AllenAx91/stock-analysis/blob/main/RESOURCES/Refactored_2018.png)

### 4.3. Conclusion 

Based on the run times using the Refactored code and the old code, it is clear that the program can be run approximately 3 times faster by avoiding the nest for loop. 




