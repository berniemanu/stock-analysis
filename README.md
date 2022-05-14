# VBA Challenge

## Overview of Project:
Refactoring or editing VBA codes learnt in module 2 to loop through all the data at the same time to collect necessary information. Refactoring is a critical next step once basic code is written. Refactoring makes the code efficient, readable, reusable with reduced run-time. This challenge is a great opportunity to practice this necessary skillset.

## Results:
VBA codes help analyse all years data for the stock universe without repeating the analysis over and over again in excel. Let's dive deeper into the 2017 and 2018 stock performance to help Steve arrive at better investment recommendations.

### 2017 Stocks Performance and Code Execution Time

#### 2017 Stocks Performance
* 2017 was a good year for the whole set of stocks under review except TERP.
* Steve's parents favorite stock DQ and SEDG nearly tripled their stock price.
* ENPH and FSLR more than doubled in value. Most of the stocks generated double digit returns.

#### 2017 Code execution Original vs. Refactored
Refactored codes took 9 times longer than the Original code to execute.

#### *Refactored code - 2017 Run time*

![VBA_Challenge_2017.png](https://github.com/berniemanu/stock-analysis/blob/main/VBA_Challenge_2017.png)


#### *Original code - 2017 Run time*

![Original_2017.png](https://github.com/berniemanu/stock-analysis/blob/main/Original_2017.png)


### 2018 Stocks Performance and Code Execution Time
* 2018 was a bad year for most of the stocks except ENPH and RUN which nearly doubled in value.
* DQ was the worst performer of all, wiping out nearly 2/3rd of its value despite a trading 3X more than 2017 volumes.
* Significant trading volume increase coupled with significant price decline clearly potrays low investor confidence in DQ stock in 2018.

#### 2018 Code execution Original vs. Refactored
Refactored codes took 9 times longer than the Original code to execute.

#### *Refactored code - 2018 Run time*

![VBA_Challenge_2018.png](https://github.com/berniemanu/stock-analysis/blob/main/VBA_Challenge_2018.png)


#### *Original code - 2018 Run time*

![Original_2018.png](https://github.com/berniemanu/stock-analysis/blob/main/Original_2018.png)


## Summary:

### *Advantages and Disadvantages of Refactoring:*
* Refactoring gives us an opportunity to refine and tweak our first draft of codes to make it more efficient and readable.
* However refactoring can become an endless process as every code has some potential to be continuously improved. 


### *Pros and Cons of refactoring in this VBA challenge:*
* Refactoring helped me understand the overall logic of the entire module coding as I had to fill in the missing blanks.
* However I believe the usage of tickerIndex and arrays for volume and prices increased the run time from <1 second to >6 seconds.
* Overall refactoring is a helpful process to dig deeper into the logic of the first draft codes and will increase code efficiency when done without constraints.
