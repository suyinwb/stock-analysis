# Stock Analysis

## Background
The initial purpose of our analysis is to help Steve's parents analyse:
* how actively DQ was traded in 2018 as they believe that if a stock is traded often, then the price will accurately reflect the value of the stock. We sum up all of the daily volume for DQ to get the yearly volume and a rough idea of how often it gets traded. We calculate the yearly return for DQ. The yearly return is the percentage increase or decrease in price from the beginning of the year to the end of the year. In other words, if you invested in DQ at the beginning of the year and never sold, the yearly return is how much your investment grew or shrunk by the end of the year.
From this, we found out that Daqo dropped over 63% in 2018.
* Due to the above, we replicated our script to analyse all the stocks on the dataset to give Steve's parents a full view of all the stocks' performance.
* In order to make the spreadsheet more user-friendly, we've included buttons to make it easier for Steve to use.
* We've also added flexiblity for Steve to input the year he is interested to analyse with aninput field.
* We've also added a functionality to calculate how long it takes to execute the output and elapsed time in a message box.

From this work, Steve is able to use the tool with ease and churn his dataset for analysis.

## Overview of Project

Steve loves the workbook you prepared for him. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although your code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.

The same dataset from his first stock analysis will be used to measure the time difference.

### Purpose
The current dataset in the spreadsheet is small and already it takes a few seconds to run each year. Therefore, in order to expand to include many years and more stocks, we need to reduce the runtime through code refactoring.

## Analysis and Challenges

## Methodology: Analytics Paradigm

#### 1. Decomposing the Ask
To get the codes working faster and more elegantly with less traversal of the dataset.

#### 2. Identify the Datasource
Same dataset is used.

#### 3. Define Strategy & Metrics
Look at the current code and visualise the calls and routines to refine and reduce data calls.
Store data into arrays.

#### 4. Data Retrieval Plan
Use stock analysis dataset in Excel

#### 5. Assemble & Clean the Data
Excel VBA scripting

#### 6. Analyse for Trends
compare timer from old codes with new codes


#### 7. Acknowledging Limitations
* Unable test out the new codes with a much larger dataset for a dry-run.
* Using VBA, the scripts will run in real-time.

#### 8. Making the Call:
The "Proper" Conclusion is indicated below on [Results](#results)

## Analysis

**2018 timer**

>Old Runtime for 2018

![Old Runtime for 2018](resources/2018_timer.png)

>New Runtime for 2018

![New Runtime for 2018](resources/VBA_Challenge_2018.png)



With the new code, our runtime is 0.238 seconds compared to 4.96 seconds.

**2018 timer**

>Old Runtime for 2017

![Old Runtime for 2017](resources/2017_timer.png)

>New Runtime for 2017

![New Runtime for 2017](resources/VBA_Challenge_2017.png)



## Challenges

### Challenges and Difficulties Encountered


## Results
From [Analysis](#analysis), we can conclude that with the new code, our runtime is 0.261 seconds compared to 5.3 seconds.

Therefore the speed increase for the data are:

**2018:** 20.8 times faster

**2017:** 20.3 times faster

## Appendix
```
Sub ClearWorksheetRefactored()

    Cells.Clear

End Sub
```
