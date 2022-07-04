# An Analysis of Stock Market Dataset Using Excel VBA

## Overview of Project
Steve, who we worked together with us in the previous modules want to do a little more research for his parents. He wants to expand the dataset to include the entire stock market over the last few years. Therefore, we want to create a logistic structured code to help him calculate the information he needs from the stock market dataset in an efficient way. We first setted the tickerindex to zero before looping over the rows and created arrays for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices. The tickerindex is used to access the stock ticker index for the arrays we created in the previous step. The script loops through the data set and storing values from the tickerStartingPrices and tickerEndingPrices. Last, we changed the color of the formatting cells, by making positive returns green and negative returns red. 
### Purpose
In this project, we are using VBA code in Excel to loop through the VBA Challenge Stock Dataset in order to collect useful information from this dataset. We want to make the code we created efficient for others to use in analyzing the data set. 
## Results
### Comparing 2017 and 2018 Stocks Analysis
By comparing the 2017 and the 2018 stocks data, there is a difference in the total daily volume between the two years that resulted in less than a $100,000,000 in increased volume. It was not enough to generate a positive 2018 return percentage. The tickers ENPH and RUN had would have been considered good investments due to the positive returns in 2018.
### Comparing the Original Times and Refactored Times 
The run times for the original code took around .4 seconds, while the run times for the refactored code took around .07 seconds for the 2017 analysis and 0.06 seconds for the 2018 analysis. Therefore, refactoring the code did make the run times decrease, which increases the efficiency, optimizing the code.
![This is an image](https://github.com/sherryli1116/ExcelVBA_Stock_Analysis/blob/main/resources/VBA_Challenge_2017.png)
![This is an image](https://github.com/sherryli1116/ExcelVBA_Stock_Analysis/blob/main/resources/VBA_Challenge_2018.png)
## Summary
### Detail Analysis of Our Result
#### Refactoring Code Advantages and Disadvantages
Advantages: We can generate the data we need in a short amount of time using the code we refactored. The stock data were being marked in red and green in a shorter amount of time.
Disadvantages: When refactoring the codes, the logistic structure is easily affected by a small typo or indent mistake, we need to double check to make sure there is no small mistakes affecting the code we produced. It is also a bit time consuming to generate these codes, we might need to reconsider this method if we are in a tight project time frame.
#### Original and Refactored VBA Script Advantages and Disadvantages:
Advantages: When creating the nested conditional and for loops in VBA, logic errors are easily to detect. The debugging function highlighted the exactly place we need to modify. Refactored VBA script led to better quality of codes which can create a shortcut for us to generate outputs. It also shows clearly the explanations on each step to run the code.
Disadvantages: However, we might need time to retest the code to make sure everything works fine. If there is one mistake, the code will not run and it might take us longer time to do the work, longer than the original method. 
### How do these pros and cons apply to refactoring the original VBA script?
To refactor the code, we need to do testing by the end of each new addition or loops to check for any small mistakes might occur. Also, the code we produced might not be suitable for another set of data or be too complicated for others to apply. Therefore, we need to make sure the code we generated has a clean format, well organized and updated.
