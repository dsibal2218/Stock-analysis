# Stock-analysis
In this challenge, we are editing, or refactoring, the VBA codes for the analysis we did in class during VBA module. We are going to refactor so that we are looping through all the data one time in order to collect the same information that we did in the VBA module. Then, we will determine whether refactoring the code successfully made the VBA script run faster. Finally, we will present a written analysis that explains the findings.

## Analysis requirements and work detail:
1. Create a tickerIndex and set to zero
2. ![image](https://user-images.githubusercontent.com/98235755/157303919-aa680be2-2b05-462e-8dac-84f021bc218e.png)

3. Arrays are created for tickers,tickerVolumes, tickerStartingPrices and tickerEndingPrices. Use the tickerIndex to access the stock ticker for all the variables mentioned.
![image](https://user-images.githubusercontent.com/98235755/157303983-7c61bda4-3389-4ea5-9359-c264a0405417.png)
![image](https://user-images.githubusercontent.com/98235755/157304060-00f3f05e-a9df-4712-9ddc-54bc9e062315.png)

4. Loop through stock data, reading and storing all of the values from each row: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices. Appropriate comment to explain the purpose of the codes are added
![image](https://user-images.githubusercontent.com/98235755/157304124-3a042f50-3ec0-47e7-b14c-8e2f999d35a8.png)

5. Format the cells in the spreadsheet
7. Compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored scripts.

## Results and Analysis

### Stock Analysis

The results of the analysis indicates that all but 1 (TERP) stock did well and had positive returns in 2017. In 2018, however, most of them had negative returns. Additionally, DQ had the highest return in 2017 but also had the highest negative return in 2018 among the 12 tickers in the dataset.
![image](https://user-images.githubusercontent.com/98235755/157330261-09a982e6-2cb5-4908-81da-518456c34fea.png)

![image](https://user-images.githubusercontent.com/98235755/157330290-e941a54c-152c-446d-a510-96bb166e36d0.png)

### Advantages and Disadvantages of Refactoring Codes
#### Advantages:
1. It improved the run time for the script - ie it ran more efficient. The run time for the original code (first two screenshots) vs refactored code (last 2 screenshots) are shown below.

![image](https://user-images.githubusercontent.com/98235755/157331806-5c106c74-6c91-4523-a7b9-6acb439031a2.png)
![image](https://user-images.githubusercontent.com/98235755/157331864-8a276e6a-c1bd-4b1e-966d-6d490174998d.png)
![image](https://user-images.githubusercontent.com/98235755/157332417-353b2dc0-9419-43b2-9d7e-61af8fb32ac2.png)
![image](https://user-images.githubusercontent.com/98235755/157332444-2da95b1d-187b-488e-8591-f39712b783db.png)
2. Another advantage is that the codes are easier to understand and debugging was easier. I had no nested for loops so it was easier to figure out where the issue was while debugging. Lastly, I found that refactoring made the codes less confusing and alot cleaner/clearer.

### Disadvantages:
1. It was tedious to do
2. In general, for somebody who doesn't have much experience and does not have much understanding of the original codes, refactoring could become risky and could actually break the program. I definitely took alot of time uderstanding it and took alot of trial and error. 
