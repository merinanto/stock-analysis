# stock-analysis
Overview of Project
Purpose of the project is to help Steve’s parents invest wisely on the energy companies.
Opening price, selling price and total volume details for two years is available in the data file.
Based on the data, returns for the share will be calculated which provides an insight for investment in the company. Code which is written for one stock DQ will be re written for all stocks and all the years.

#Results:

Code was refactored with option to enter which year must be considered for calculation and returns for all the share/company are calculated.
https://github.com/merinanto/stock-analysis/VBA_Challenge.xlsm
When the scripts ran for first time it’s taking almost 1.8 seconds. But when script was ran again and again execution time is below 1 seconds.
https://github.com/merinanto/stock-analysis/blob/main/resources/VBA_Challenge_2017.png
https://github.com/merinanto/stock-analysis/blob/main/resources/VBA_Challenge_2018.png
#Summary 

Due to refactoring of returns for each share/company can be compared for different years and how company progressed in 2 years.

As array is being used performance is much better than compared to without arrays.

One of disadvantage is at a time one data for one year is available at excel file .So  for comparison study data for each year must be copied manually to a different file . Its easy for analysis if data for both year available in same sheet.

Another disadvantage is if any year unavailable in the sheet is entered as input the exception is not handled.








 
