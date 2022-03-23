# stock-analysis with Excel VBA
Refactoring code to loop through data and VBA scripting 
## Overview of Project 
### Background 
For this project we are helping Steve and his parents look through a collection of stock, to better inform them of which stocks to invest in. They are currently interested in green energy production and have their mind set on DAQO New Energy Group, however Steve wants to look through several thousand stocks in addition to this one and has created an Excel file to do so. My role is to use VBA, an extension on Excel to automate the process. 
### Purpose 
The purpose of this project is to refactor code in order to loop through the data and collect specific information. We are looking to make the process more efficient by taking the original code and making it easier to read and execute. 
## Results
### Analysis
Using the code provided, the foundation to organize our data was established. We formatted the corresponding sheet, initialized arrays, prepared for the analysis of tickers, looped through the tickers and rows in order to output the data for the tickers we selected. By also comparing the stocks from the original and refactored data set, it was evident that the original macro had ran for about .30 seconds longer than the refactored script. Based on the outputs of the selected tickers, one can determine that ENPH and RUN both had positive return values in both 2017 and 2018. If Steve is looking to diversify his parents stocks, those may be two viable options for him. The code and screenshots are provided below. 
[Code provided] (https://2u-data-curriculum-team.s3.amazonaws.com/dataviz-online/module_2/challenge_starter_code.vbs)
![VBA Challenge 2017](https://user-images.githubusercontent.com/100797549/159597947-5259dfec-dc96-4b76-aa10-3676e1f91dc9.png)
![VBA Challenge 2018](https://user-images.githubusercontent.com/100797549/159598062-0f960d86-d65d-4835-adbe-46952ddc2d2c.png)
## Summary 
### Refactoring Code
Advantages of refactoring code help the user present concise and organized data. It requires less steps, memory and time to accomplish the task. This would also be beneficial to anyone else who is looking at the code you created, in order to clearly understand and pinpoint certain commands or actions. Some disadvantages of refactoring often include times when the refactorer is someone else rather than the original coder. Bugs could be harder to avoid, applications may be too large or incompatible with the systems in place. 
### Refactoring VBA Script 
The advantages of refactoring this VBA script was the amount of time we shaved to create the analysis. The facility to read, debug and run the script was the greatest benefit. The only disadvantage I faced with refactoring the VBA script was adjusting my Worksheet files to coincide with those in the script. I had included no spaces in the title words, for example I had “AllStocksAnalysis” rather than “All Stocks Analysis” and VBA was rather particular about that so I made the necessary adjustments. 
