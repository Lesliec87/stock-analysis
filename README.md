# Stock-Analysis

## Overview of Project: 
In this project we are assisting Steve, who has recently graduated from his finance degree and has taken on his parents as his first clients. His parents are eager belivers in green energy becoming more promonent in the future and have started investing in green energy stocks. 

However, they had not done much research on this and recently invested all their money into a company called DAQO New Energy Corp. Steve is looking into DAQO stock and also wanting to analyse a handful of other green energy stocks to determine, if his parents need to diversify their funds.

We have been working with the excel file Steve created containing the stocks data and he is happy with the workbook he can at the click of a button, analyze an entire datase. Steve wants to do a bit more research where the database includes the entire stock market over the last few years. 

### Purpose:
The purpose of this project is to refactor the code so we can run it succesfully when analysing thousands of stocks instead of only a dozen stocks and we will determine if refactoring the code it now runs faster. 

## Results:

1. We were able to refactor the code succesfully for 3012 stocks from the past two years (2017 and 2018). For both years we have included the stocks of 12 different organizations where we collected their Total Daily Volumes and Returns. 

### All Stocks 2017: 

![Stocks2017](https://github.com/Lesliec87/stock-analysis/blob/main/Resources%202/Stocks(2017).png)

### All Stocks 2018: 

![Stocks2018](https://github.com/Lesliec87/stock-analysis/blob/main/Resources%202/Stocks(2018).png)

- We can see that for 2017 the stocks mostly had positive returns except for TERP stocks that had a loss in returns. 
- We can see that for 2018 the stocks had mostly negative returns except for ENPH stocks and RUN stocks.
- As for DAQO they had a 199.4% positive return in 2017 and a 62.6% negative return in 2018. 
- Ultimately, Steve can advise his parents to diversify into ENPH stocks and RUN stocks since they have consistently had positive returns for the past 2 years. 

2. After running successfully the code we were able to confirm that they now run in less time, for 2018 the time was around 0.133 seconds and for 2017 the time was around o.125 seconds. Review the time performed for both years below: 

![time2017](https://github.com/Lesliec87/stock-analysis/blob/main/Resources%202/VBA_Challenge_2017.png)

![time2018](https://github.com/Lesliec87/stock-analysis/blob/main/Resources%202/VBA_Challenge_2018.png)

3. To have a more efficient way for the code to run for a more extended database in this analysis we added a "tickerIndex" to access the correct index across the four different arrays which consisted of the 12 tickers array and the three output arrays: tickerVolumes, tickerStartingPrices and tickerEndingPrices. This was used to create the "For loops" that ran through each row and pulled the information to get the results for [All Stocks 2017](#All-Stocks-2017) and [All Stocks 2018](#All-Stocks-2018).

**Below you can see the code that as it was finally ran:** 

![code1](https://github.com/Lesliec87/stock-analysis/blob/main/Resources%202/Code%201.png)

![code2](https://github.com/Lesliec87/stock-analysis/blob/main/Resources%202/Code%202.png)

![code3](https://github.com/Lesliec87/stock-analysis/blob/main/Resources%202/Code%203.png)


## Summary: 

In this analysis we were able to assist Steve with getting sufficient data on what the best green stocks that his parents could potentially invest in to diversify their investment and if their current investment in DAQO has been a profitable investiment for the future. 

**What are the advantages or disadvantages of refactoring code?**

-The advantages of refactoring the code were that we could expand the amount of data we could run, faster time to run the code than the original code and Steve  will now be able to apply this to his future clients he might have interested in investing. 
-However, the disadvantages about refactoring the code is that it is a bit more time consuming and more debugging to go through since you have to make changes to original code. 

**How do these pros and cons apply to refactoring the original VBA script?**

- The pros were determined by the change in time it took in running for the original code the times for both years were from 0.55 seconds to 0.75 seconds and in the refactored code the times lowered to run from 0.125 seconds to 0.135 seconds. 
- The cons were that the time to debug the refactored code was definitely a lot longer since the code was a bit more complex to figure out and you need to take that time into consideration for a projects deadline.

