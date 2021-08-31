# Stock Market Analysis (2017 – 2018)
## Overview
The purpose of this project is to provide a full analysis of the stock market between the 2017-2018 period based on 12 different tickers. 
To accomplish this goal, I will provide a code designed in VBA as a mechanism to analyze a large amount of data set. This solution will provide flexibility for our client to analyze thousands of stocks in the future. 

## Results
### Stock Ticker ranking and suggestions
After comparing the individual stocks results to the average between the 2017-2018 period, I can rank the ticker stocks to provide investment advice.

<p align="center"><img src="https://user-images.githubusercontent.com/88695570/131441155-2054b725-8e36-4b9c-9a7c-304b2c3acacc.png">

  **Best options for investment**

  1. ENPG: This ticker shows a positive return for both years. If we analyze the average for the 2-year period, we can obtain a positive return of 105.75%. This calculation shows   that this stock is the most profitable, robust, and sustainable option over the time, even though the yearly volume is not one of the highest. 
  
  2. RUN: Despite of having the lowest return value in 2017, it obtained the highest percentage on the dataset for the 2018 period.  The 2018 period affects all stocks negatively,   but RUN stock remained as one of the best options. Same as the ticker ENPG, this is a positive pattern for long-term investments, especially when the RUN ticker is the fifth       highest volume on the data set.
 
  **Other possible options with good return**
  
  To label these ticker as “Second order”, I considered the 2017-2018 average and a possible investment with a projection of 2 years as a reference:
  
  3. SDEG
  4. DQ
  5. FSLR
  6. VSLR

  **Options with low retunr**
  
  These tickers obtained a positive result. However, their average return is smaller compared to the two previous groups.

  7. CSIQ
  8. HASI
  9. AY

  **Stock ticker not recommended**

  Based on the individual results by year and the average for the 2-year period, I would recommend the client to avoid these stocks. These three options have negative returns, but   the one that keeps the same negative results for each period is TERP. I would suggest to not invest in these stocks
  
  10. JKS
  11. TERP  
  12. SPWR

### Logic perfomance
  **Code desgined during the offline Modules vs Code desgined during the Challenge**
  
  These are the code’s speed results: 
  
  |   Year    |   Code - Module Section  | Second Header |
  |   :---:   |           :---:          |    :---:      |
  |   2017    |           0.593          |    0.531      |   
  |   2018    |           0.578          |    0.523      |  
  
  The code structure for this week’s challenge turns out to be slightly more efficient in comparison to the code developed during the module section. I will elaborate a hypothesis   of the possible reasons for this result on the Summary section.  

## Summary  
### Refactoring Code - Advantages and disadvantages
  **Advantages**
  - It helps to find the most efficient code for the program.
  - It reduces execution time. This can help to analyze bigger dataset.
  - It challenges the developer to improve his own coding techniques.

  **Disadvantages**
  - It is possible to obtain an opposite result and increment the performance time. 
  - It required more ability and experience. This could affect the project`s time of release for the client. 

### Refactoring Code - Advantages and disadvantages
For this case, refactoring the code made the VBA script run faster. After comparing both codes, the difference lays on where the nested for loops are executed. 

  - Code designed for the Modules: The main loop is based on the number of tickers and the secondary loop (internal) is running through the row of the raw data. Once the secondary   loop is completed, the results are written on the cells. Then, the logic continues with the main for loop. It seems like this process increases the time because the main loop     sequence is interrupted to perfmon an action directly on the worksheet.
 
  - Code designed for the Challenge: In this case, the main loop runs over the raw data and the internal secondary loop goes through the 11 stocks. The most important detail for     this solution is how the results are stored in VBA arrays. The results are written at the end of all the sequence over the excel cells.
  This method could reduce the execution time because the loops are not interrupted during their internal sequence. 
