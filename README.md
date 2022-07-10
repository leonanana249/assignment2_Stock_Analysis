# Assignment2_Stock_Analysis
## Overview of Project 
The purpose of this project is to refactor codes to loop through stock tickers to generate total volumes and returns in 2017 and 2018 with higher efficiency. 
## Results 
Most of tickers performed better in 2017 than 2018 since the returns and daily volumes were higher in 2017. Except for TERP, it had higher daily volumes and slightly lower negative return in 2018. The run time for the 2017 calculation is slighter longer than that in 2018 loop. The reason could be the higher volumes in 2017 required longer time to calculate. Moreover, the refactored script runs shorter than the original one because the tickerIndex, tickerVolumes, tickerStartingPrices and tickerEndingPrices are all defined as arrays. The arrays will push the logic to calculate using the number values instead of going through the ticker strings. Therefore, the loops can run more efficiently. 

> ![image](https://user-images.githubusercontent.com/107721712/178129155-387ebb2e-3cbc-43ea-bebd-805d5b95e7c9.png)

![image](https://user-images.githubusercontent.com/107721712/178129221-ab916339-6838-44f7-ac71-12791573ec8a.png)

![image](https://user-images.githubusercontent.com/107721712/178129320-757dc21b-dbcf-4c3f-9896-72401dc7e204.png)
![image](https://user-images.githubusercontent.com/107721712/178129349-3a8518f3-6ee8-4ec4-9981-345f61561e5c.png)
![image](https://user-images.githubusercontent.com/107721712/178129555-850d400b-f64e-4ae1-84e4-d51fe16d035e.png)
## Summary
### Advantages and Disadvantages of Refactoring Codes 
Refactoring codes can help the users to apply scripts to a higher volume of data with efficiency. However, it might need more time to interpret the nested logics. 
### Advantages and Disadvantages of Refactoring to Original VBA Scripts 
Refactoring to original VBA scripts means users don’t need to build a script from scratch. People can save time only on efficiency improvement. However, refactoring may take long time and only save couple seconds of running time. Therefore, the users should really think the benefits of refactoring if the scripts will be applied to much heavier dataflows. 
