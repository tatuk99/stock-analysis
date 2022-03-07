# Stocks Analysis

## Overview of Project

This project was created to analyze the outcomes of refactoring code, in order to speed up the run time of the analysis in VBA. Steve, the person who is asking us to create this code, originally asked us to analyze stock results of 12 stocks. However, he wants to expand the selection, and analyze a much greater volume of stocks in the future. Therefore, we must refactor this code in order to make it faster, and work for a larger set of data. 
In this analysis, we are looking at the run time of the original, not-refactored code, and comparing it to the run time of the refactored code. 

## Results

### Run time for the original VBA code
If we run the original code, we get the following run times for 2017 and 2018, for the 12 stocks that are part of the analysis in this project:

![Run time for 2017 (original code)](Resources/VBA_Original_Time_2017.png)
![Run time for 2018 (original code)](Resources/VBA_Original_Time_2018.png)
	
### Run time for the refactored code
The run time of the refactored code, is outlined below:

![Run time for 2017](Resources/VBA_Challenge_2017.png)
![Run time for 2018](Resources/VBA_Challenge_2018.png)

Evidently, the refactored code is running faster than the original code, which is a great sign, since it will mean that a larger set of data will run with much more efficiency than the original code used in the analysis.

## Summary

- What are the advantages or disadvantages of refactoring code?
	
The clear advantages of refactoring code are that it becomes much more concise and easy to understand (for the creator of both the original and refactored code), along with speeding up the process for the computer. When comparing the challenge VBA code with the original code, it has less steps and more direct lines, which make more sense to someone who is refactoring it, since the functionality stays the same. Additionally, this shortened version of the code makes it easier to use for larger datasets, since we are using "i" as the variable to account for all tickers in this analysis, having the computer go through everything line by line, finding the correct one to analyze. The consiceness of the newer code also speeds up the process of computing the results we want the computer to run, making the code run faster. 
Some disadvantages include having somebody else look at the refactored code, and understand what is going on. The simplification leads to certain elements being missing, therefore can be somewhat confusing for somebody who has not seen the original code before. The simplification can also lead to mistakes which weren't there before, causing a longer amount of time spent on debugging errors caused by the simplification. 

- How do these pros and cons apply to refactoring the original VBA script?

My VBA script now runs faster than it did before, and the code is shorter than the original. The computer does not need to analyze previous things to the extent that it did with the original code. As far as the negative aspect of the refactoring, I ran into times where I make mistakes with naming the variables, missing an "s" at the end of certain names. Even when coming back to the code again after a few days, the issue of figuring out where I had left off and why I was writing the lines the way I was was sometimes a challenge, because certain steps were no longer there. I had to go back to the original code to figure out what I was looking at. 
Now it runs faster - the main pro.
