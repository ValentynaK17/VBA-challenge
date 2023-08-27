# VBA-challenge
yearly stock data analysis using VBA
Summary:
This script helps with following analysis of stock market data for various tickers over multiple years:
a) Calculates the yearly change in stock price for each ticker symbol, by the following approach: 
	 1. for each ticker we search for records dated earliest and latest per year 
  2. from the opening price at the earliest date of the year we subtract the closing price at the latest date of the year
 b) Calculates the percentage change in stock price for the same period, as mentioned in a), by the dividing yearly change by opening price at the start of the year
Values for a) and b) are highlighted in such a way that positive changes have green background, negative changes have red background and 0 has yellow background
 с) Calculates the total stock volume over the year for each ticker symbol, by summing up all volumes for a ticker and a year
 D) Identifies the stocks with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume" by looking for max, min and max values respectively within percentage change in stock price per ticker, and volume per ticker
The script can be executed across multiple Excel sheets at once by looping through each worksheet
Running specifics:
In order to run the script excel should be opened with enabled macro selection
Key Findings:
RKS is definitely not the ticker to invest as it constanly (for all 3 analysed years) has the highest decrease in yearly price change. Highlighting of positive/nagative year change in price helps us with quick understanding of stock performance
Comparing year-over-year tickers' total volume helps us with understanding the most popular ones and what are the market trends
