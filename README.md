# Stocks Analysis


## Overview of Project

Is it a good idea to invest in a company (DQ) that makes silicon wafers for Solar Panels?



### Objective

 The purpose is to discover, automate and analyze how actively was a set of stocks and what was the real value of every stock in 2017 and 2018 to determine if investing in DQ a good decision.



### Data source 

An excel file with the stock data. [Green Stock Dataset](https://github.com/lindaperez/stocks_analysis/blob/main/green_stocks.xlsx?raw=true)

![image](https://user-images.githubusercontent.com/1729991/180491463-29246473-0e58-468c-aace-36290fe0c04f.png)

  Facts:
  - Yearly volume: How often a stock gets traded. 
  - Yearly return: Is the percentage increase or decrease in price from the beggining of the year to the end of the year. 
    - Example:  if you invested in DQ at the beginning of the year and never sold, the yearly return is how much your investment grew or shrunk by the end of the year.



[Non refactored solution](https://github.com/lindaperez/stocks_analysis/blob/main/green_stocks.xlsm?raw=true)


[Solution](https://github.com/lindaperez/stocks_analysis/blob/main/VBA_Challenge.xlsm?raw=true)



## Results:


### Stock performance 2017 and 2018

In 2017, The stock DQ was not very actively traded, with a minimum volume of 35,796,200. The most actively traded and valuable was FSLR with 684,181,400 of total daily volume. The  yearly return of the ticker DQ was the most higher with 199.4%


![Picture Analysis non Refactored: 2017](https://user-images.githubusercontent.com/1729991/180496113-2bb7300e-c7ae-4d6d-aaf2-24362364f2aa.png)<img width="280" alt="Screen Shot 2022-07-22 at 10 55 14 AM" src="https://user-images.githubusercontent.com/1729991/180496621-c9b0bb0a-e2fe-4c80-9156-20ce9fd184cd.png">




In 2018, ENPH was the stock most actively traded with 607,473,500 and a very high return of 81.9%. The other most important stock was RUN with a higher volume of 502,757,100 with the higher return of the year 84%. The ticker DQ was down from the average in terms of daily trading and had a loss of 62.6%


![Picture Analysis non Refactored: 2018](https://user-images.githubusercontent.com/1729991/180496299-3cc8dd3d-8519-474a-be9d-d9a642e0205f.png)
<img width="280" alt="Screen Shot 2022-07-22 at 10 53 46 AM" src="https://user-images.githubusercontent.com/1729991/180496424-019104ea-4d1d-4c35-80fa-54252440526b.png">
  



### Approaches 

In the first development, the approach was to traverse the whole sheet and find volumes and returns one time for every stock. The complexity for this process was  O(n*m),  where n is the number of different stocks and m is the total number of rows a total of 10*3013.

After the refactor, the approach was to traverse the whole sheet finding volumes and returns only one time for all the stocks.  The complexity of this process is only  O(m), where m is the number of rows in the sheet, a total of 3013.

In this sense, if it is given 1000 stocks the complexity with the first approach would be 10000*3013 O(m*n). But the complexity for the second approach would be only the number of rows, O(m).  

### Analysis of execution times (2017,2018)


The execution times in 2018 were 0.3125 and 0.1797 before the refactor and after the refactor respectively, the improvement of the performance was around 42% approximately.


- Picture Analysis Refactored: 2017
<img width="575" alt="Screen Shot 2022-07-22 at 11 19 17 AM" src="https://user-images.githubusercontent.com/1729991/180500159-0c918810-39c5-46fb-8284-a99ae6ec0ab9.png">



The execution times in 2018 were 0.3438 and 0.2227 before the refactor and after the refactor respectively, the improvement of the performance was around 35% approximately.


- Picture Analysis Refactored: 2018
<img width="573" alt="Screen Shot 2022-07-22 at 11 16 50 AM" src="https://user-images.githubusercontent.com/1729991/180499803-1ef6a14a-56a9-444f-8e13-0900c5bd18e8.png">





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
