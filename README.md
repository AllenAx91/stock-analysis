# Stock Analysis using VBA

## 1. Overview of Project
This challenge employs Excel VBA to assess the total volumes and returns of stocks in 2017 and 2018. Our stakeholders are Steve's parents, who would like to use a user-friendly medium to do this task. The analysis will also detail the learning from improving the code through refactoring.

## 2. Purpose
  1) To provide Steve's parents with Total Daily Volumes and Anual returns for the years 2017 and 2018
  2) Developing the excel into a user-friendly medium by including a Run stock Analysis button and a Clear data button
  3) To arrive at the most time-efficient code by exploring the ideology of Refactoring
  4) To make the code reusable, if in case the stakeholders need to do additional analysis

## 3. Result of Data Analysis

### 3.1. Year 2017

All the targeted stocks faired well in 2017, except for the ticker TERP. The highest was DQ, with nearly 200% return. It proved to be a good year for those who have invested in these stocks.

_Image of analysis for the Year 2017 along with time taken to execute code_

![](https://github.com/AllenAx91/stock-analysis/blob/main/RESOURCES/VBA_Challenge_2017.png)

### 3.2. Year 2018

Except for ENPH and RUN, all other stocks failed in 2018. TERP had brought good returns in 2017 but went down to -5.0%. Steve's parents were interested in DQ. However, it was evident that it did not do well this year.

_Image of analysis for the Year 2018 along with time taken to execute code_

![](https://github.com/AllenAx91/stock-analysis/blob/main/RESOURCES/VBA_Challenge_2018.png)

## 4. Run time optimization

The simplest way to write a program for this challenge was to use nested for loops. A revised logic can then be used to refactor the code after the successful implementation of the first code. However, one must consider the advantages and disadvantages of refactoring the code. 

* Advantages: 
  1) The refactored code can be used for larger data
  2) Faster results can be derived
  3) Highly sort after if the priority is storage efficiency
* Disadvantages:
  1) Very time-consuming. If the scope of the project is small and has strict time constraints, the aim should be to write a simpler code
  2) Refactoring makes the code more complex and is harder to de-bug. For example, a simple overlooked spelling error can set the programmer behind by many hours
  3) Not viable when there is a risk of bugs escaping even the testers or quality assurance team
  4) Not economically feasible if there is not enough budget allocated for the overall project
  5) Other may find it hard to comprehend the code if there is not enough documentation explaining the code

_Go to this link for more on [Refactoring](https://anarsolutions.com/code-refactoring-concept-analysis/#:~:text=Maintainability%3A%20After%20refactoring%2C%20the%20code,no%20idea%20where%20to%20go.)._

After careful consideration of the above pros and cons, two VBA programs were written to explore which method is more time-efficient. The first one was with a nested for loop which resulted in running lines equal to 12 * the number of rows. On the other hand, the refactored code used Arrays to gather all information in one shot.

### 4.1. Code with nested 'For loop'

As mentioned in the previous section, the nested for loop involved running more lines of the code. If there were more stock data available in the worksheets, the nested for loops would have taken longer to give results. However, this code was the quickest logic to arrive at. 

_Screen shot of Run times for the Years 2017 and 2018 using **nested for loop**_

![](https://github.com/AllenAx91/stock-analysis/blob/main/RESOURCES/NestedFor_2017.png) ![](https://github.com/AllenAx91/stock-analysis/blob/main/RESOURCES/NestedFor_2018.png) 

### 4.2. Refactored code with Arrays

The number of lines executed in the code was less since only one 'for loop' was used. This time, the code ran only once through all the rows in the worksheet while simultaneously storing the values of Starting Price and Ending price for each ticker. We had to spend more time modifying the old code to a complex one.

_Screen shot of Run times for the Years 2017 and 2018 using **arrays**_
![](https://github.com/AllenAx91/stock-analysis/blob/main/RESOURCES/Refactored_2017.png) ![](https://github.com/AllenAx91/stock-analysis/blob/main/RESOURCES/Refactored_2018.png)

## 5. Pros & Cons of the original and refactored script

The refactored program executed the code, approximately three times faster by avoiding the nested for loop. It is a considerable improvement, especially when our stakeholders expressed interest in trying this code on other data sets. The refactored code has given the stakeholders the option of reusing this code.

_(See section 4 for more advantages and disadvantages)_

To further break down the pros and cons for future considerations: 

### 5.1. Code with nested 'For loop'

1) _Advantages_: 
* If our stakeholders needed our help urgently, this code would have served their purpose
* Very easy to understand

2) _Disadvantages_: 
* This programme has only a 'one-time" application
* Code needs editing for it to be transferrable

### 5.2. Refactored code with Arrays

1) _Advantages:_
* Faster run time - approximately three times faster 
* Very storage efficient

2) _Disadvantages:_
* Time-consuming to code for a first-time programmer 
* Hard to De-bugg


