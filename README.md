# **Challenge 2: Stock Analysis**
## Review of Stock Analysis by Year using VBA 
### The purpose of this analysis was to teach us how to use VBA in Excel to process data. We learned how to use If, For Loops, Formatting, Arrays, and basics of VBA to process Stocks on 2017 and 2018 and compare the difference.
## Analysis and Challenges
### **Analysis of Stocks Based on Year** 
To perform the comparison of the stocks and 12 specific tickers we created a new sub "AllStocksAnalysisRefarctored" and activated the new worksheet "All Stocks Analysis" to format the sheet by creating a header row and identifying the array of the tickers we are going to use to run the search. See the spreadsheet [here.](https://github.com/mthalken/stock-analysis/blob/main/VBA_Challenge.xlsm)

After preping the worksheet for our results, we went back to our year worksheet by allowing user input for the specific [year](https://github.com/mthalken/stock-analysis/blob/main/User_Input_Box.PNG). We defined out output arrays and index to hold the data we are searching for. We then created a For Loop to run the rows in the spreadsheet and add the ticker ID Volume to our ticker Index. We used an If statement to find the starting price and ending price and store within the specified index. 
![png](https://github.com/mthalken/stock-analysis/blob/main/For_Loops_%26_If_Statements.PNG)

We then wrote a For Loop to go through the arrays to output the Ticker, Total Daily Volume, and Return on the "All Stocks Analysis" worksheet. 
![png](https://github.com/mthalken/stock-analysis/blob/main/Header_Row.png)

We formatted the worksheet "All Stocks Analysis" and used conditional formatting to show if the return was negative(RED) or positive green. 
![png](https://github.com/mthalken/stock-analysis/blob/main/formatting.PNG)

As part what the client asked for was a timer to show the runtime of each search. To do this we created a start time and endtime timer and a message box at the end to display the amount of time the code took to run in seconds. 
![png](https://github.com/mthalken/stock-analysis/blob/main/Start_Time_Code.PNG)
![png](https://github.com/mthalken/stock-analysis/blob/main/End_Time_Code.PNG)
![png](https://github.com/mthalken/stock-analysis/blob/main/VBA_Challenge_2017.png)
![png](https://github.com/mthalken/stock-analysis/blob/main/VBA_Challenge_2018.png)

The challenge asked us to refactor the analysis. Through this I was able to make it run faster by creating another array for ticker volumes and use that inserted thoughout the for loop. 

### **Challenges and Difficulties Encountered:** 
The challenges that I encountered was learning to write the specific VBA code. I understood the base verbage but was missing the specific numbers of outputs on the arrays in 1b. I kept getting a "Compile error: Expected array" code. I was able to utilize the Slack AskBCS Learning Assistant to walk me through the understanding of defining the arrays. 
## Results
Through this analysis of 2017 and 2018 stock prices we can see that all tickers that we ran except one had a positive return in 2017 but only 2 had positive returns in 2018. This concludes that if we were to make a recommendation based on the two years of data that we have the best options for investing would be ENPH and RUN. To make a more studied recommedndation it would be benifical to see more current stock volumns as well as just a larger field of years to support the recommendation. 
![png](https://github.com/mthalken/stock-analysis/blob/main/2017_All_Stock_Analysis.PNG)
![png](https://github.com/mthalken/stock-analysis/blob/main/2018_All_Stock_Analysis.PNG)

## Summary
- What are the advantages or disadvantages of refactoring code?
    The advatages are that refactoring cleans up the code you have written to make it easier to understand, run faster, and helps find bugs. The only disadvantage is that it takes time, so being able to refactor code is very benificial in the long run. 
- How do these pros and cons apply to refactoring the original VBA script?
    In this challenge we saw a decrease in time run from refactoring the VBA script. We were also able to dive deeper into the understanding of VBA to prepare for future VBA scripts we will run. 
