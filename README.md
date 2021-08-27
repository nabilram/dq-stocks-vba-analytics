# dq-stocks-vba-analytics

## Overview
* Client is in need of data wrangling services for financial / stock price calculations
* Client has provided functioning XLSM macro for their inital calculations
* Client requests a refactoring of his XLSM macro for efficiency gains
* XLSM file contianing source macro and refactored macro is found here: [Refactored XLSM Macro](https://github.com/nabilram/dq-stocks-vba-analytics/blob/main/resources/green_stocks.xlsx) 

## Results
### Before

* Source macro had run time of *0.844* seconds for 2017 and *0.836* seconds for 2018. See below. 
![Source_Macro_Runtime_2017](https://github.com/nabilram/dq-stocks-vba-analytics/blob/main/resources/Before_Refactor_2017.PNG)
![Source_Macro_Runtime_2018](https://github.com/nabilram/dq-stocks-vba-analytics/blob/main/resources/Before_Refactor_2018.PNG)

### After
* Refactored Macro had run time of 0.160 seconds for 2017 and and 0.164 seconds for 2018. See below. 
![Output_Macro_Runtime_2017](https://github.com/nabilram/dq-stocks-vba-analytics/blob/main/resources/VBA_Challenge_2017.PNG)
![Output_Macro_Runtime_2018](https://github.com/nabilram/dq-stocks-vba-analytics/blob/main/resources/VBA_Challenge_2018.PNG)

### Effeciency Gain
* For  both 2017 and 2018, over ~80% in efficiency gains for macro runtime was achieved through refacoring the code. 
* Efficiency Gain calculated as:
    * 1 - (Refactored Run Time / Source Macro Run time)

## Summary

