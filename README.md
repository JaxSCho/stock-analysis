# VBA of Wall Street

## Overview of Project

### Purpose

Steve's parents are looking to invest in green energy stocks. Using his green energy stock dataset, we were able to help Steve with his performance analysis of a dozen green energy stocks - precisely the total daily volume and yearly return. Now Steve wants to expand the dataset to include the entire stock market over the last few years and would like an efficient VBA code to compile the results of this larger dataset. This project will present the green energy stock performance findings between 2017 and 2018 and if the refactored VBA script could execute the original VBA script faster.

## Results

The following results analyze the stock performance between 2017 and 2018 and the execution times of the original script and the refactored script used to analyze Steve's green energy stock dataset.

### Stock Performance

To compare how the green energy stocks performed between 2017 and 2018, we examined how actively the stocks were traded and how much an investment grew or shrunk by the end of the year if one never sold after an initial investment at the beginning of the year -- total daily volume and yearly return -- for both years. To get a rough idea of how often a particular stock was traded in a year, we summed up the total number of shares traded throughout the day (i.e., total daily volume). The yearly return is the percentage difference in price from the beginning to the end of the year. Figures 1 and 2 display the 2017 and 2018 total daily volumes and yearly returns, respectively, for a dozen green energy stocks. 

The 2017 total daily volume ranged from 35,796,200 shares of DQ stock to 782,187,000 shares of SPWR stock, with an average of 263,886,592 shares traded throughout the day. Conditional formatting allows us to quickly determine that only the TERP stock made a negative yearly return in 2017.

![VBA_Challenge_2017](https://github.com/JaxSCho/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

<b>Fig.1 - 2017 Stock Performance and Refactored Code Execution Time</b>

For 2018, the total daily volume ranged from 83,079,900 shares of AY stock to 607,473,500 shares of ENPH stock, with an average of 275,503,183 shares traded throughout the day. Once again, conditional formatting allows us to quickly determine that ENPH and RUN were the only stocks that saw positive returns in 2018.

![VBA_Challenge_2018](https://github.com/JaxSCho/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

<b>Fig.2 - 2018 Stock Performance and Refactored Code Execution Time</b>

Overall, the stocks analyzed in this dataset performed better in 2017 compared to 2018 (i.e., 92% of the stocks had positive returns in 2017 vs. 17% in 2018) despite the higher trading activity in 2018 (i.e., 4% increase in average total daily volume in 2018).

### Execution Times of Original vs. Refactored Scripts

Since Steve may want to perform his stock analysis on larger datasets in the future, we proceeded through a code refactoring exercise to improve the performance analysis execution time. To consider if the refactored script can compile the results faster than the original script, we used the `Timer` function in VBA to calculate how long the code takes to execute and output the elapsed time in a message box. 

The `startTime` is set after the code receives the year input from the user, and the `endTime` is set before the script concludes. By subtracting the `startTime` from the `endTime`, the message box states how long the code took to run --

`MsgBox ("This code ran in " & (endTime - startTime) & "seconds for the year " & (yearValue))`

The original VBA code took about 0.7 seconds to compile the results for either 2017 or 2018 stock datasets (0.6953 seconds and 0.7070 seconds for 2017 and 2018 datasets, respectively). The refactored VBA code took less than 0.2 seconds to execute the scripts for either the 2017 or 2018 stock datasets (0.1680 seconds and 0.1836 seconds for 2017 and 2018 datasets, respectively), as seen in Figures 1 and 2 above. Therefore, the refactored script improved the execution time of the stock performance analysis by half a second.

## Summary

The results of this analysis were created using Steve's [Green Energy Stock dataset](https://github.com/JaxSCho/stock-analysis/blob/main/VBA_Challenge.xlsm). This document provides the background data files and VBA scripts for this analysis and can be used for future project insights.

- What are the advantages and disadvantages of refactoring code?

When refactoring code, a primary task is to improve the code's quality and efficiency by removing redundancies and duplications, using less memory, or improving the logic and execution time of the code. Another primary goal of refactoring code is to reduce a project's overall technical debt. In other words, if code is not cleaned up, it can snowball, slowing down future improvements because your future self and other developers will have to spend extra time understanding and tracking code before they can change in later iterations. Clean code can also be reused and the basis for code elsewhere, thus, reducing time spent in initial future code development. Refactoring can also help developers better understand code and design decisions. As the code's functionality increases or shifts, there are benefits from seeing how others have worked inside and built up the code.

However, since we are not changing or adding to the code's external behavior and functionality, the refactoring process may seem superfluous and unimportant compared to higher priority tasks (i.e., the benefit is not apparent). Imprecise refactoring could also introduce new bugs and errors into the code. If larger teams work on refactoring, the coordination effort required could be surprisingly high. 
 
- How do these pros and cons apply to refactoring the original code?

For a novice coder, refactoring the original code provides an excellent mechanism to learn and engage their creativity and resourcefulness with coding while picking up valuable code quality improvement practices. The use of descriptive comments improves the comprehensibility of code for all programmers, which aids in maintaining and extending code. By using arrays, we optimized the functionality of the original code over the long term by improving execution time for future analysis of larger datasets. The refactored code can also be reused or applied to future coding projects, reinforcing the learned coding concepts and shortening the code building time. Since the original code and dataset were quite rudimentary with minimal performance impact, the benefits of code refactoring may be overshadowed by time spent debugging new bugs and errors, especially for a beginner. 
