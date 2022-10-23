# VBA_Challenge Optimization
## Overview of Project 
Our goal was to analyze 2017 & 2018 stock ticker data and compile the stock data using VBA Scripting..  
### Purpose
The purpose of this task was to optimize a previously created VBA Script AllStocksAnalysis and compare the run time to a newly created VBA script AllStocksAnalysisRefactored with a goal to reduce the amount of time it took for the script to run. 
## Analysis and Challenges
### Analysis
The AllStocksAnalysis VBA script calculated the tickers information utilizing a more basic structure looping through each ticker individually and calculating its data.  
 

The AllStocksAnalysisRefactored VBA script initiated 3 arrays with each set to 0.  Then it looped over each ticker calculating its information.  This process condensed the number of lines in the script and allowed it to run more quickly. 
 
 

We confirmed the AllStocksAnalysisRefactored VBA script ran quicker based on the below screenshots.
Below are screenshots of the Original VBA Scripts
2017							2018
  

Screenshots of the New VBA Script utilizing arrays & indexes.
2017							2018
  

We validated the end results were the same but the run time was not.
AllStocksAnalysis				   AllStocksAnalysisRefactored
   


### Challenges
The main challenge in this project was learning about how to imbed the array for the 3 objects in the correct spot.  It took me a while to understand that in our arrays we needed to initialize the array after the object.
 
Then initializing the tickers(tickerIndex) to keep track of the current ticker.
 




## Summary
Advantages of Refactoring Code
•	Programs will run quicker and more efficiently.  With long pieces of code, it will take a while for the program to run possibly even days.  With refactored code it can optimize these programs and allow them to run and take up the least amount of time, as there might be outside stakeholders waiting for the program to finish. The quicker the programs run the sooner stakeholders can get their data and information.
•	With quicker run times for code, it will consume less energy.  With code that runs faster it will consume less energy, if you’re running your own business and must pay the electricity bill each month and have these large programs running constantly you can cut down on your bill by refactoring and optimizing code.
•	Learning new methods of coding.  With refactoring code, you might be introduced into a new way to code and learn some new techniques and tools that will benefit you in the future.
Disadvantages of Refactoring Code
•	Redoing code that already works.  The main disadvantage of refactoring code is redoing a project that is already working and complete.  It can take a lot of time and energy to review, analyze, and optimize it with the hope that your efforts will result in more optimized code.
•	Trying to refactor code might have a reverse effect and make the project run longer and less efficient.  With refactoring code, the main goal is to have it run quickly and more efficiently, the opposite can occur.

Advantages of the refactoring AllStocksAnalysis and creating AllStocksAnalysisRefactored 
•	Learned about how to include arrays within objects to optimize the program.  Including arrays within the object allows us only run through the code once with a single for loop to gather the data.  as noted previously noted in the analysis section in the AllStocksAnalysis we initiated two for loops to loop through the tickers and again to calculate each tickers information.  
 

Disadvantages of refactoring the code
•	In the original script AllStocksAnalysis we had functioning code that ran quickly with accurate results.  We had to add additional time to learn about arrays and indexes to optimize the code with only a benefit of saving about a half a second between run times.  The benefit based on run time was minimal.  




