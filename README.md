# dq-stocks-vba-analytics

## Overview
* Client is in need of data wrangling services for financial / stock price calculations
* Client has provided functioning XLSM macro for their inital calculations
* Client requests a refactoring of their XLSM macro for efficiency gains
* XLSM file contianing source macro and refactored macro is found here: [Refactored XLSM Macro](https://github.com/nabilram/dq-stocks-vba-analytics/blob/main/resources/VBA_challenge.xlsm) 
    * Source Macro is Sub "AllStocksAnalysis"
    * Output Macro is Sub "AllStocksAnalysisRefactored"

## Methods and Results

### Methods
* The use of nested for loops through arrays and assigned indices was utilized
    * An Array ("tickers") for stock ticker names (*strings*) was created
    * An Index ("iTicker") was nested within the Array (*integer*), to keep track/assign values to these tickers
* Conditional formatting along with functional buttons were added for enduser's ease of use
    * Green highlight for positive, calculated Return values. Red highlights for negative Return values
    * Button to initialize analytics added; button to clear data for possible macro re-runs added

### Before
* Source macro had run times of _**0.844**_ seconds for 2017 and _**0.836**_ seconds for 2018. See below:
![Source_Macro_Runtime_2017](https://github.com/nabilram/dq-stocks-vba-analytics/blob/main/resources/Before_Refactor_2017.PNG)
![Source_Macro_Runtime_2018](https://github.com/nabilram/dq-stocks-vba-analytics/blob/main/resources/Before_Refactor_2018.PNG)

### After
* Refactored macro had run times of _**0.160**_ seconds for 2017 and and _**0.164**_ seconds for 2018. See below:
![Output_Macro_Runtime_2017](https://github.com/nabilram/dq-stocks-vba-analytics/blob/main/resources/VBA_Challenge_2017.PNG)
![Output_Macro_Runtime_2018](https://github.com/nabilram/dq-stocks-vba-analytics/blob/main/resources/VBA_Challenge_2018.PNG) 

### Effeciency Gain
* For 2017 and 2018, over ~80% in efficiency was gained for macro runtimes through refactoring the code
* Efficiency Gain $ calculated as:
    * 1 - (Refactored RunTime / Source Macro RunTime)

## Summary
* Refactoring Advantages - Output Macro
    * As seen above, **an ~80% efficency gain was achieved** with the same result data
    * Memory burden on one's machine is less -- as it needs to do less work/calculations

* Not Refactoring Disadvantages - Source Macro
    * 80% of run time is wasted
        * While this may seem small in seconds, if big data sets are used then this will result in minutes or hours of run time wasted
    * Redundancy in similar code design
        * A set of non-nested, loop statements in linear order would have been functional with the same results
        * This, however, does not take advantage coding best pracices
            * In this case, that missed best practice is: nesting similarly designed code - with indices, counters, and data levers

* Refactoring Challenges - Domain Expertise
    * A level of domain expertise on the nature, scope, and meaning of the data itself is of great importance
    * The deeper the domain expertise available the easier and more efficient the refactoring output is
    * In this challenge, more time was spent in understanding the finance data (and calculations needed) than the actual technicality of nested for loops