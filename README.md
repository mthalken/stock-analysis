# **Challenge 2: Stock Analysis**
## Review of Stock Analysis by Year using VBA 
### The purpose of this analysis was to teach us how to use VBA in Excel to process data. We learned how to use If, For Loops, Formatting, Arrays, and basics of VBA to process Stocks on 2017 and 2018 and compare the difference.
## Analysis and Challenges
### **Analysis of Stocks Based on Year** 
To perform the comparison of the stocks and 12 specific tickers we created a new sub "AllStocksAnalysisRefarctored" and activated the new worksheet "All Stocks Analysis" to format the sheet by creating a header row and identifying the array of the tickers we are going to use to run the search. (Link to worksheet)

After preping the worksheet for our results we went back to our year worksheet by allowing user input for the specific year (png of year question). We defined out output arrays and index to hold the data we are searching for. We then created a For Loop to run the rows in the spreadsheet and add the ticker ID Volume to our ticker Index. We used an If statement to find the starting price and ending price and store within the specified index. (png of code)

We then wrote a For Loop to go through the arrays to output the Ticker, Total Daily Volume, and Return on the "All Stocks Analysis" worksheet. (png of code)

We formatted the worksheet "All Stocks Analysis" and used conditional formatting to show if the return was negative(RED) or positive green. (png of worksheet png)

As part what the client asked for was a timer to show the runtime of each search. To do this we created a start time and endtime timer and a message box at the end to display the amount of time the code took to run in seconds. (png of start and end time code)


### **Challenges and Difficulties Encountered:** 
The challenges that I encountered was learning to write the specific VBA code. I understood the base verbage but was missing the specific numbers of outputs on the arrays in 1b. I kept getting a "Compile error: Expected array" code. I was able to utilize the Slack AskBCS Learning Assistant to walk me through the understanding of defining the arrays. 
## Results
### **Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.**
This concludes that May is the best month to start a theater campaign and December is not. 
## Summary
- What are the advantages or disadvantages of refactoring code?
- How do these pros and cons apply to refactoring the original VBA script?
