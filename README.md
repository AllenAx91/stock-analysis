# Stock Analysis using VBA

## 1. Overview of Project
This challenge employs the use of Excel VBA to asses the total volumes and returns of certain stocks in the years 2017 and 2018. Our stakeholders are Steve's partents who would like to use a user-friendly medium to do this task. The analysis will aslo detail how the program was improved through refactoring. 

## 2. Purpose
  1) To provide Steve's parrents with:
    I. Data Analysis
      a. Total Daily Volumes
      b. Anual returns for the years 2017 and 2018
    II. Developing Excel into a user-friendly medium by Including a:
      a. Run stock Analysis button
      b. Clear data button
  3) To arrive the most time efficient by exporign the ideiology of Refactoring

## 3. Result of Data Analysis

### 3.1 Year 2017

All the targeted stocks faired well in the year 2017 except for the ticker "TERP". The maximum being "DQ" with nearly 200% return proved to be a good year for those who may have invested in these stocks. 

_Image of analysis along with time taken to execute code _

[]{https://github.com/AllenAx91/stock-analysis/blob/main/RESOURCES/VBA_Challenge_2017.png}


Two sets of VBA programms were written to explore which method is time efficient. 

### 3.1. Challenges and Difficulties Encountered

To initiate our analysis, the data file had to undergo multiple levels of refining to arrive at the interpretable state it is in now. Some of the issues encountered are listed below:

 - Unix Date format - _Modified to short date format_
 - Category and Subcategory were merged - _Categories were split from subcategories using "Text to columns option"_
 - _Filtering options were enabled on the header_

On confirming that the data was readable, we performed some statistical analyses and then proceeded with a targeted analysis strategy as detailed in the remainder of this report.

### 3.2. Analysis of Outcomes Based on Launch Date

#### 3.2.1. Analysis through Data Visualization
For this analysis, a line chart was generated using excel to portray the influence of the launch date on the outcomes of theater campaigns since 2015. The X-Axis and Y-Axis represent the "**Count of Outcomes**" and "**Months of the year**" respectively.
The representations of the coloured lines are as listed below:
 * Blue: Successful Campaigns
 * Orange: Failed Campaigns 
 * Grey: Cancelled Campaigns



#### 3.2.2. Data interpretation

Although this chart provides a lot of information, the two points listed below stand out. 
 * _May has the most number of successfull campaigns_
 * _The graph also relveals that the count of failed outcomes in May are almost the same for the months of Jun, Jul, Aug and Oct_

### 3.3. Analysis of Outcomes Based on Goals

#### 3.3.1. Analysis through Data Visualization

To dwell deeper into the data and comprehend why certain plays were successful, another line chart was generated. This time, the X-Axis represents **ranges of goals** (roughly in 5000 increments) and the Y-Axis represents the **Outcome's percentage**. The three outcomes have are represented as lines as detailed below:
 * Blue: Successful Campaigns
 * Orange: Failed Campaigns 
 * Grey: Cancelled Campaigns



#### 3.3.2. Data interpretation

Clearly, as the goals proceed to be extremely ambitious, the percentage of successful campaigns declines from **76%** in the range <1000 goals to **50%**in the range (15000 to 19999). 

## 4. Conclusion 

The month of May is probably the most successful because of the general outlook of the publicity in this season. It is during the spring that people are more than happy to indulge in leisure activities. This information will prove to be extremely useful for Louise's scheduling. One must also pay attention to the fact that the total number of campaigns in May is the highest of the year. This signifies that there is more to just launching a campaign in May. The competition is higher as more campaigns happen in the vicinity. Perhaps, a collaboration with other upcoming artists to hold joint campaigns can be considered. Also, It would be pragmatic for Louise to marginally mark down her initial campaign goal as the percentage of successful campaigns is significantly higher below the goal range of 10000 to 14999.
 
## 5. Limitations

The data is only a sample size as it was retrieved from a single website. Lousie will have to do a similar analysis based on data from other sources to validate the results explained in Section 4 "Conclusion" of this report. 

## 6. Recommendations 

In addition to the analysis performed, Louise has access to the excel file named ["Kickstarter_Challenge.xlsx"](https://github.com/AllenAx91/kickstarter-analysis/blob/main/Kickstarter_Challenge.xlsx) that has been set up to help Louise derive more conclusions. Here are some of the recommended analyses that could be performed are: 
   1) Box charts to determine and rule out outliers
   2) Barcharts to display which subcategory proved to be more successful
   3) Filters can be updated to include country and then Line charts can be brought up to see how location influences outcome
