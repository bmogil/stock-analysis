# Stock-Analysis
## Overview of Project
### Purpose
* The purpose of this challenge is to put our refactoring skills to work. As we learned, refactoring is a critical part of programming with any language because a programmer's first attempt at code won't always be the best way to accomplish a task. And, sometimes refactoring someone else's code will be our only objective when working with the existing code at a job.
* More specifically, the purpose of this assignment is to refactor the Module 2 solution code to loop through all the data once and accomplish the same end-goal, wherein, we garnered the same stock data we did throughout completing the asynchronous module.
---
## Results
### 2017 All Stocks Analysis Run Time Results
The end results displayed the refactored code, running the macro for the "2017" sheet, to be 2.515625 seconds faster than the original code, as shown by the images below:
![](Resources/2017_run_time_original.png)
![](Resources/2017_run_time_refactored.png)
---
### 2018 All Stocks Analysis Run Time Results
The end results displayed the refactored code, running the macro for the "2018" sheet, to be 2.1171875 seconds faster than the original code, as shown by the images below:
![](Resources/2018_run_time_original.png)
![](Resources/2018_run_time_refactored.png)
---
## Summary
### Advantages/Disadvantages of Refactoring Code
#### Advantages
1. It is easier to refactor code than start from scratch because once you determine that you have acquired a 'template' per se, of code that will execute your desired task(s), likely all you have to do is perform some minor tweaks to customize it to your specific task.
2. Further, it is easier to spot minor/major inefficiencies within a script of code when you are starting with code in front of you, opposed to writing a first draft yourself and trying to spot the same inefficiencies because of receny/self-bias.
### Disadvantages
1. If the code does not have comments it can be harder to understand what's going on within the code, and thereby harder to do anything with it.
2. The code could be organized differently compared to the best practices we are learning in our bootcamp, also making it harder to interpret, analyze, use, etc. 
---
### Advantages and Disadvantages of Original/Refactored VBA Script
## Original
1. An advantage of the original script is that it is easier to understand at first glance (in my opinion at least). To further elaborate, I think that since the code is not refactored, the individual tasks within the code are more spread out and in my opinion, it takes less thought-processing time to comprehend many isolated statements than multiple combined statements (especially as a beginner programmer).
2. A disadvanatge of the original script is clearly its lacking of efficiencies that are created through refactoring. Although the original script may be easier to comprehend, the refactored script is superior in and of itself because of the synergies realized upon refactoring. For example, the original script had a nested for loop, while the refactored script did not. The refactored script had one primary for loop used to acquire all of our desired information. Moreover, since a nested for loop is exponential in nature, it makes sense why the refactored code runs faster than the original. The following is a picture of the original code with the nested for loop:
![](Resources/original_code.png)
## Refactored
1. Refactoring has many advanatges, as the one discussed above. But, another advantage could be that it is more suitable for experienced programmers to analyze my refactored code than my original code because their minds are not thinking in the same isolated steps that beginner programmers like me like to have. Another advantage of the refactored code is that it combines the formatting macro and stock info. macro into one macro, thereby, only one button needs to clicked from the front-end to run the entire macro. Lastly, the refactored script does not use a nested for loop, but rather uses one primary for loop to obtain all the desired stock data. It does this by making use of two other for loops that act in sequence to help the primary for loop execute its desired functions successfully. Additionally, the refactored code makes use of a "tickerindex" variable that helps the different arrays within the primary for loop keep track of the current ticker being iterated. The following is a picture of the refactored code:
![](Resources/refactored_code.png)
