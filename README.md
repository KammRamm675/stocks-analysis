# The Wolf of Davis Street

## Overview of Project
    
### Purpose

-   The purpose of this analysis is to provide three pieces of information concerning run-time differences between original and refractured script as well help Steve with understanding the performances of chosen stock between the years 2017 and 2018.  

## Analysis and Coding 

### Analysis of New VBA Script Compared to Previously Used Script. Does it Run Faster?

-   During the process of developing both pieces of VBA script it did become evident in the differences of amount of time VBA needed to process the stock information.  To recap, please reference https://github.com/KammRamm675/stocks-analysis/blob/main/VBA_Challenge.xlsm for the actual Excel Sheets that are used in the analysis.  The VBA script being used is the refactored version which is connected to the "Run Analysis For All Stocks" button on the "All Stocks Analysis" sheet.  After clicking said button, a pop-up screen is shown with the amount time it takes (run-time) for each year (2017 and 2018) to run an analysis of stock information on the pages of the same year.  2017 gives a run-time of 0.6484375 seconds (please refer to https://github.com/KammRamm675/stocks-analysis/blob/main/VBA_Challenge_2017.png).  2018 gives us a run-time of 0.6445313 seconds (please refer to https://github.com/KammRamm675/stocks-analysis/blob/main/VBA_Challenge_2018.png). NOTE: the times proved are the results of the refactored VBA script. Roughly one-tenths of a second slower.

-   The original version used before being refactored gave us a result time of 0.546875 and 0.5429688 seconds for the years 2017 and 2018, respectively.  To get these results, the macros on the "Run Analysis for All Stocks" button had to be changed back to the first version before clicking the button.  CODING NOTE: Using the code "yearValue" in string like "Worksheets(yearValue).Activate" allows analysis to run on a selected year with the button provided.  replacing "yearValue" with a specific year within range will result in only values for the year specified. 

## Results

### What Should Steve Advise His Parents To Invest In? 

-   There are HUGE differences between 2017 and 2018 concerning the twelve (12) stocks in question.  Please refer to the same afore mentioned Excel Sheet "All Stocks Analysis."  To clarify we are looking at stocks for tickers AY, CSIQ, DQ, ENPH, FSLR, HASI, JKS, RUN, SEDG, SPWR, TERP, and VSLR.  After running the analysis for year 2017 e see we are mostly in the green for all stocks.  DQ experienced almost a 200% return for investments made.  Which is a large return.  TERP experienced a loss of 7.21% making it one (1) of two (2) stocks that lost money from initial investment. 

-   The 2018 year practically pulled a bad turn around. Almost all stocks lost money with ENPH having an 81.92% return making (1) of (2) stocks that were in the green.  DQ lost the most experience a loss of 62.60%. 

-   Overall, there are stocks that came out on top.  If you take the values of returns for 2017 and added them to the returns for 2018 there is a gross overall return on investment.  (2017 percentage + 2018 percentage = growth overall return on investment).  Please, reference the PDF, "Gross Overall Return," https://github.com/KammRamm675/stocks-analysis/blob/main/Gross%20Overall%20Return.pdf.  Looking at this PDF we can tell Steve that a wise investment would be ENPH and RUN.  This is because both show continual growth.  NOTE:  ENPH did make the biggest overall return BUT the second and third largest stocks (DQ and SEDG) showed signs of shrinking returns and practically all other stocks lost more than half of their initial returns. This is why Steve is advised that RUN and ENPH are safe investment STRICTLY due to these stocks’ growth patterns.  CODING NOTE: Gross over all return was determined by the summation of percentages of years between 2017 and 2018 (see above explanation and equation). 

## Summary

### Advantages and Disadvantages of Refactoring Code in General.

-   To reiterate, refactoring is the process of taking code that already exists and restructuring it without changing the outcome.  

-   Advantages:  There’s a few advantages to refactoring.  Readability, Maintainability, and extensibility.  Readability improvements make it easier for any level of developer to read and understand the code.  Maintainability improvements provide developers a code that’s easier to look over and correct any problems that may arise.  Maintainability can be a result of improved readability. With both improved readability and maintainability, the extension at which the refactored code can add new abilities and uses can rise. 

-   Disadvantages:  As I have found many times in the refactoring process, a disadvantage of refactoring is the introduction of bugs.  If I didn't write a part correctly after downsizing a portion (for instance with nesting loops) I always came across "run-time error "6": overflow" and "run-time error "9": Subscript out of range." In a business sense and from what I experienced refactoring TAKES TIME and time = money. You would really have to make sure that refactoring something would have greater benefits in the long run. 

### Advantages and Disadvantages of the Original and Refactored VBA Script. 

-   Advantages:  As I mentioned in the above "disadvantages" portion, the original VBA script will save users money. It’s there.  Its tried and true. As the military says, "if it's not broke, why fix it?" It'll save time, money, and effort that all can be put towards something else.  That being said, an advantage of refactored VBA script would be casted out in the long run.  Sure, it'll cost time, money, and effort but will make it easier to maintain (creating more time freed up for other activities), more room in computer memory, and get results desired faster.  I would guess that refactoring helps find the bugs in the original code.  I don’t look at my run-time errors as mistakes on my part so much as I look at it as "there’s a problem in this line of code."

-   Disadvantages:  I feel the biggest disadvantage to having the original code is very dependent on the code itself.  If it's a code where it takes 500 lines to develop to get the results we were challenged with this week (getting the return for two different years on twelve (12) stocks) then it’s useless because it took me 115 lines. So over complexity is a disadvantage of original code.  I would also assume that refactoring might cause security concerns.  If it’s too easy to read then anyone with any level of coding can manipulate the VBA script as to oppose a complex code would weed out the 13 year old in a coffee shop on their phone. 
