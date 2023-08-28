# VBA-challenge
yearly stock data analysis using VBA
## Summary
This script helps with following analysis of stock market data for various tickers over multiple years:
*  a) Calculates the yearly change in stock price for each ticker symbol, by the following approach: 
	 * 1. for each ticker we search for records dated with earliest and latest date of the year 
    * 2. from the opening price at the start of the year we subtract the closing price at the end of the year.
 * b) Calculates the percentage change in stock price for the same period, as mentioned in a), by the dividing yearly change by opening price at the start of the year
Values for a) and b) are highlighted in such a way that positive changes have green background, negative changes have red background and 0 has yellow background
 * с) Calculates the total stock volume over the year for each ticker symbol, by summing up all volumes for a ticker and a year
 * d) Identifies the stocks with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume" by looking for max, min and max values respectively within percentage change in stock price per ticker, and volume per ticker
The script can be executed across multiple Excel sheets at once by looping through each worksheet
* e) The bonus part allows us having the merged summary of data in separate sheet along with being grouped by ticker X year
##  Running specifics:
In order to run the script Multiple_year_stock_data.xlsx should be opened with enabled macro selection and newly created macro with the script added in Stock_YearStatistics_Option1.bas
In order to run extended version use script in Stock_YearStatistics_Extended.bas
## Final results:
By running scripts for Multiple_year_stock_data.xlsx as an input there will be results, like shown in Screenshots_results folder
Highlighting of positive/negative year change in price helps us with quick understanding of stock performance
Comparing year-over-year tickers' total volume helps us with understanding the most popular ones and what are the market trends.
For example, it’s worth to put an attention to DJH or LOM that have one of the highest volume sum for the last 3 years and have constant positive grow in yearly price change
Similarly, it is obvious that RKS is the worst player as it constantly (for all 3 analyzed years) has the highest decrease in yearly price change.
