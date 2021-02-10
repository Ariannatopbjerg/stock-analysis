# Stock Analysis using Microsoft Excel VBA code
## Background
Steve wanted to know more about the stock DAQO to see if his parents should invest in this stock. I developed Microsoft Excel VBA code so Steve can analyze his dataset with a click of a button. This code allowed Steve to look at a few years (2017,2018) of information showing what the stock was doing. According to the data, DAQO was not successful. Knowing this, Steve now wants to look at the entire stock market over the last few years. 
## Purpose 
Since my code might not work as efficiently and/or fast enough for a greater number of stocks, the goal is to refactor my VBA script so that when Steve looks at his dataset or other longer datasets, the code will be much faster to process the data.  
## Methods
### The Data
Inside `VBA_Challenge.xlsm`, you will find two spreadsheets named “2017” and “2018.” These two sheets contain information about 12 stocks: when the stock was issued, the opening, closing and adjusted closing price, the highest and lowest price, and the volume of the stock. The third sheet; “All Stocks Analysis,” shows all the 12 stocks and their total daily volume and return rate for a specific year. The year will depend on which year you choose to type in when using the VBA script. 
### Steps to refactoring data 
To refactor the code I previously had, I copied certain chunks of the code to create the input box, chart headers, ticker array, to activate appropriate worksheets, and to clear contents. In the VBA script that was given for the refactoring, steps were given to set up the proper structure. In the VBA script within `VBA_Challenge.xlsm` , you will see the steps and code for the updated code.
### Challenge
Since I am new to VBA code, it was difficult to construct the code so it would work properly. I am used to R and python syntax, and at times wanted to write my code in those forms. With trial and error, I was able to figure out the code I needed for the refactoring.
## Summary
### Results
After refactoring the code, the VBA script ran about three times faster when looking at 2017 data and 2018 data than the original. 
