
# **Stock-Analysis**

  After providing Steve with a workbook to analyze stocks his parents were interested in investing in, he now would like to explore a wider range of stocks for his parents to invest in. 

## **Overview**

  The purpose of conducting this analysis is to determine how Excel's Visual Basic Application can be used more efficiently to analyze multiple sets of data.

## **Results**

https://github.com/Emc1518/stock-analysis/blob/68d3369767608bd15cdfa8ebafbe8fe280d60b4e/VBA_Challenge_Code

In refactoring the code I was able to remove the `AND` operators from the 'For` loop as well as remove variables that were already included in another place.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/81889167/116769678-d6339900-aa0b-11eb-9b1a-886a653a5e3d.png)


![VBA_Challenge_2018](https://user-images.githubusercontent.com/81889167/116769679-d6cc2f80-aa0b-11eb-867f-1b91857b367e.png)

The refactored code ran in 0.14 seconds for 2017 and 0.12 seconds for 2018, where the original code ran in 0.58 seconds for 2017 and 0.57 for 2018.  The refactored code ran roughly 0.3 or 0.4 seconds faster than the original code therefore the refactored code was more efficient. 


  The most important advantage of being able to refactor code is that it allows you to run the code in less time so it is much more efficient for projects with an enormous amount of data. However, I would say that a disadvantage of refactoring code is making errors while trying to find inefficiencies. If a code was already working and you made changes to it that did not work, now you have ruined the original code. 

  Refactoring the original was challenging because it seemed to be working just fine and then trying to trim off fractions of seconds from the run time was hard to do. I found that I had to have a better handle on the rules and concepts in VBA before I could even begin to refactor the original code. 
  
