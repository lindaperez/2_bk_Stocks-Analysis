# AllStocksAnalysis
## Overview of Project

### Purpose
The purpose is to reflect on some techniques to improve the development of macros using Excel and VBA. To do this it was used a dataset in excel named AllStockAnalysis where it was performed several steps to answer how actively was a set of stocks and what was the real value of every stock in the years 2017 and 2018.

##Results:
### Stock performance 2017 and 2018

In 2017, The stock DQ was not very actively traded, with a minimum volume of 35,796,200. The most actively traded and valuable was FSLR with 684,181,400 of total daily volume. The  yearly return of the ticker DQ was the most higher with 199.4%

![This is an image](https://github.com/lindaperez/stocks_analysis/blob/main/Resources/VBA_Challenge_2017_Prev.png)

Picture from AllStocksAnalysis non Refactored 2017

In 2018, ENPH was the stock most actively traded with 607,473,500 and a very high return of 81.9%. The other most important stock was RUN with a higher volume of 502,757,100 with the higher return of the year 84%. The ticker DQ was down from the average in terms of daily trading and had a loss of 62.6%

![This is an image](https://github.com/lindaperez/stocks_analysis/blob/main/Resources/VBA_Challenge_2018_Prev.png)

Picture from AllStocksAnalysis non Refactored 2018

### Execution Times 2017 and 2018

The execution times in 2017 were 0.3125 and 0.1484, before the refactor and after the refactor respectively, the improvement with the refactor of the performance was around 53% approximately.

In the first development, the approach was to traverse the whole sheet and find volumes and returns one time for every stock. The complexity for this process was  10*3013, O(n*m),  where n is the number of different stocks and m is the total number of rows.

After the refactor, the approach was to traverse the whole sheet finding volumes and returns only one time for all the stocks.  The complexity of this process is only 3013 O(m), where m is the number of rows in the sheet.

In this sense, if it is given 1000 stocks the complexity with the first approach would be 10000*3013 O(m*n). But the complexity for the second approach would be the same as 3013 O(m).  


Picture from AllStocksAnalysis 2017
![This is an image](https://github.com/lindaperez/stocks_analysis/blob/main/Resources/VBA_Challenge_2017.png)

The execution times in 2018 were 0.3438 and 0.1484 before the refactor and after the refactor respectively, the improvement of the performance was around 59% approximately.


Picture from AllStocksAnalysis 2018
![This is an image](https://github.com/lindaperez/stocks_analysis/blob/main/Resources/VBA_Challenge_2018.png)


## Summary

## What are the advantages or disadvantages of refactoring code?

### Advantages

- If it is a large dataset it will run in a shorter time.
- It could be more readable, consequently easy to maintain.
- It could be better organized, consequently easy to maintain.
- It could be an opportunity to do a Double-check of the functionality.
- Bugs can be solved.

### Disadvantages

- I take time to understand again the problem and think about the way to solve it efficiently.
- It is necessary to have a good understanding of the programming language, data structures,  and testing.
- It could take more time to do it than the original solution.
- Risk doesn't keep the functionality.

## How do these pros and cons apply to refactoring the original VBA script?

To refactor the code It was used a dictionary to access the offset of every stock in O(1) time.
The sheets were traversed in one time storing the data efficiently, just in a one-pass. To extract data from the sheet it was used the property Value2 that gets the values from the sheet in a raw way, without formatting. For the formatting of the sheets, it was used the property With to change the element in one pass.

### Pros
Now, the solution is more readable and faster, the code is better organized, and became self-explanatory. The functionality changed during the refactoring but it was fixed comparing the results and following the initial logic. It is encapsulated by chunks of functionality which makes it better to maintain.

### Cons
It took time to understand the business, the data structures available and limitations of VBA with Excel, and techniques to improve the functionality. During the refactoring, the functionality changed but some tests were done to fix it and make it return the results expected.
