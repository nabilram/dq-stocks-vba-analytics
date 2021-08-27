# dq-stocks-vba-analytics

##Overview
* Client is in need of data wrangling services for financial / stock price calculations
* Client has provided functioning XLSM macro for their inital calculations
* Client requests a refactoring of his XLSM macro for efficiency gains
* XLSM file contianing source macro and refactored macro is found here: [Refactored XLSM Macro](C:\Users\nabil\Documents\GitHub\dq-stocks-vba-analytics\resources\green_sticks.xlsx) 

##Results
### Before
* Source macro had run time of 0.844 seconds for 2017 and and 0.836 seconds for 2018. See below. 
![Source_Macro_Runtime_2017](C:\Users\nabil\Documents\GitHub\dq-stocks-vba-analytics\resources\Before_Refactor_2017.PNG)
![Source_Macro_Runtime_2018](C:\Users\nabil\Documents\GitHub\dq-stocks-vba-analytics\resources\Before_Refactor_2018.PNG)

### After
* Refactored Macro had run time of 0.160 seconds for 2017 and and 0.164 seconds for 2018. See below. 
![Output_Macro_Runtime_2017](C:\Users\nabil\Documents\GitHub\dq-stocks-vba-analytics\resources\VBA_Challenge_2017.PNG)
![Output_Macro_Runtime_2018](C:\Users\nabil\Documents\GitHub\dq-stocks-vba-analytics\resources\VBA_Challenge_2018.PNG)

### Effeciency Gain
* For  both 2017 and 2018, an ~80% in efficiency gain in runtime was was achieved through refacoring the code. 
* Efficiency gain calculated as:
    * 1 - (Refactored Run Time / Source Macro Run time)
    
##Summary

