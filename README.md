# Analyzing Daily Stock Market Data (2018-2020)

## **Motivation/Reason for Analysis:**

The motivation for this analysis was to assess expertise with Excel VBA and to understand my ability to use the software to evaluate and analyze a large data set.

## **About the Data Set:**

The data set selected for this analysis contains 2018-2020 stock market data including daily opening and closing prices, high and low values, and the total trading volumes for the stocks listed on the stock market (exhange not specified).

## **Approach:**

I wrote an Excel VBA script that loops through the worksheets in the workbook (the different years) and outputs to a summary table a list of the stocks listed on the exchange using their ticker symbols and calculates their annual change in price, percentage change, and the total volume traded. I also added another table where users can view the stocks with the greatest percent change (both increase and decrease) and greatest total volume traded for that year. These two summary tables allow users to both compare stocks over the time horizon analyzed as well as determine which stocks performed the best and conversely the worst in a specific year. 

## **Takeaways/Lessons Learned:**

* When dealing with a large data set like the one used in this analysis, it is easiest to break down the data set into a smaller one and use this new smaller data set to build our code. Once the code works on this smaller data set and outputs the data/information we want, we can apply our code to the larger data set. This assignment already broke the larger stock market data set down for us--by giving us a smaller data set with just A-F stocks for a given year to loop through-- but this is certianly a lesson learned that I will leverage in future assignments and in my work. 

* Once we have the smaller data set, it then helps to decompose the question, problem, task, etc. into smaller questions, problems, task so that building our code appears less daunting than when approaching the problem/task as one large question. For example in this analysis, I started working in the smaller data set with just stocks A-F for a given year and I did not proceed to calculating the yearly change until my code was able to retrieve all stock ticker symbols. Similarly, I did not proceed to calculating the percent change until my code successfully counted the yearly change, and so on. Additionally, I did not add the worksheet loop functionality until my code successfully completed all tasks on the first worksheet. Finally I did not apply my code to the larger data set until all tasks were completed on every worksheet in the smaller data set. 
    * Ultimately, creating a smaller data set from the larger one and decomposing the question/problem into individal questions problems, etc. makes the task at hand more approachable and allows us to think one line of code at a time.

* From the completed analysis of the stock market data set, it is clear the performance of a stock in one year does not influence its performance in the next year. Similarly, the total volume traded in a year can vary, likely depending on the stocks performance.


## **Further Anaysis:**
Moving forward, it would be interesting to visually display the data using Excel graphs. For example, it would be interesting to view the trends of a stock over our time horizon using a line chart. It would also be interesting to determine if there was a relationship between yearly change and total volume using a scatter plot. If I had to guess, I would say we would find a somewhat positive correlation. :) 


## **Code File:**
* [Code File](02_VBA_StockMarketScript.bas) - This file contains my code with comments for the assignment. 
