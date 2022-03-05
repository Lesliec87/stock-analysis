# stock-analysis

## Overview of Project: 
In this project we are assisting Steve, who has recently graduated from his finance degree and has taken on his parents as his first clients. His parents are eager belivers in green energy becoming more promonent in the future and are started investing in green energy stocks. 

However, they had not done much research on this and recently invested all their money into a company called DAQO New Energy Corp. Steve is looking into DAQO stock and also wanting to analyse a handful of other green energy stocks to determine, if his parents need to diversify their funds.

We have been working with the excel file Steve created containing the stocks data and he is happy with the workbook he can at the click of a button, analyze an entire datase. Steve wants to do a bit more research where the database includes the entire stock market over the last few years. 

### Purpose:
The purpose of this project is to refactor the code so we can run it succesfully when analysing thousands of stocks instead of only a dozen stocks and we will determine if refactoring the code it now runs faster. 

## Results:

We were able to refactor the code succesfully for 3012 stocks from the past two years (2017 and 2018). 

For both years we have included the stocks of 12 different organizations where we collected their Total Daily Volumes and Returns. 

All Stocks(2017): 

![Stocks2017](https://github.com/Lesliec87/stock-analysis/blob/main/Resources%202/Stocks(2017).png)

All Stocks(2018): 

![Stocks2018](https://github.com/Lesliec87/stock-analysis/blob/main/Resources%202/Stocks(2018).png)

- We can see that for 2017 the stocks mostly had positive returns except for TERP stocks that had a loss in returns. 
- We can see that for 2018 the stocks had mostly negative returns except for ENPH stocks and RUN stocks.
- As for DAQO they had a 199.4% positive return in 2017 and a 62.6% negative return in 2018. 
- Ultimately, Steve can advise his parents to diversify into ENPH stocks and RUN stocks since they have consistently had positive returns for the past 2 years. 

After running successfully the code we were able to confirm that they now run in less time, for 2018 the time was around 0.133 seconds and for 2017 the time was around o.125 seconds. Review the time performed for both years below: 

![time2017](https://github.com/Lesliec87/stock-analysis/blob/main/Resources%202/VBA_Challenge_2017.png)

![time2018](https://github.com/Lesliec87/stock-analysis/blob/main/Resources%202/VBA_Challenge_2018.png)

To have a more efficient way for the code to run this analysis we added a "tickerIndex" to access the correct index across the four different arrays which consisted of the 12 tickers array and the three output arrays: tickerVolumes, tickerStartingPrices and tickerEndingPrices. This was used to create the "For loops" that ran through each row and pulled the information to get the results for (# All Stocks(2017)) and (#(All Stocks(2018)).



Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.


## Summary: 

In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?
