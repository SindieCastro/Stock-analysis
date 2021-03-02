# Stock Analysis with Excel Using VBA

## Overview of Project

### Background

The client asked for assistance in creating a workbook to analyze green energy stock performance data for 2017 and 2018, specifically for DAQO New Energy Corporation (DQ). The workbook allowed the client to analyze the performance of DQ with a click of a button. The client was impressed with the workbook and asked for another workbook to expand his analysis.  

### Purpose

The purpose of this project is to create a workbook to expand the client’s analysis of a dozen green energy stocks using Visual Basics for Applications (VBA) in Excel.
## Results

### Workbook 

The workbook begins with a view of two buttons; one to run the analysis and the other to clear the sheet. *(See Figure 1: Buttons)*.

**Figure 1: Buttons**

![VBA_Challenge_buttons](https://user-images.githubusercontent.com/78306719/109703053-a591c900-7b5a-11eb-8bf0-60f4a5dc3197.PNG)

Once the Run Analysis button is clicked, a window pops up asking “What year would you like to run the analysis on? *(See Figure 2: Question Pop-up)*.

**Figure 2: Question Pop-up**

![VBA_Challenge_question](https://user-images.githubusercontent.com/78306719/109703246-da9e1b80-7b5a-11eb-9ece-0d0cea1b9b82.PNG)

After the year is typed, a table pops-up stating the amount of time it takes to run the code. *(See Figure 3: Time Code Ran for 2017 and Figure 4: Time Code Ran for 2018)*. In this instance, the 2017 data was analyzed in 0.1484375 seconds and the 2018 in .2109375 seconds.

**Figure 3: Time Code Ran for 2017**

![VBA_Challenge_2017](https://user-images.githubusercontent.com/78306719/109703448-1a650300-7b5b-11eb-92ff-93c2a8191a3a.PNG)

**Figure 4: Time Code Ran for 2018**

![VBA_Challenge_2018](https://user-images.githubusercontent.com/78306719/109703698-63b55280-7b5b-11eb-8fba-841eebe16e88.PNG)

### VBA Coding 

The VBA coding was created to analyze DQ but it was refactored to analyze a dozen stocks. Although the coding looks similar, there are minor changes to expand the analysis. For instance, the refactored code has a tickerIndex set to zero (tickerIndex = 0). This part of the code was not executed with the DQ analysis code *(See Figure 5: Compare Workbook Codes)*.

**Figure 5: Compare Workgroup Codes**

![VBA_Challenge_Copare codes](https://user-images.githubusercontent.com/78306719/109703848-91020080-7b5b-11eb-84f2-05d453e03924.PNG)

### Stock Performance for 2017 and 2018

The analysis shows the return investment percentage for most stocks was high in 2017 (shown as green) while the majority decreased in 2018 (shown as red). DQ performed well in 2017 at 199.4% but underperformed in 2018 t -62.6% *(See Figure 6: All Stocks Analysis for 2017 and Figure 7: All Stocks Analysis for 2018)*.

Figure 6: All Stock Analysis for 2017

![VBA_Challenge_2017 results](https://user-images.githubusercontent.com/78306719/109702780-4df35d80-7b5a-11eb-87aa-770f0c96b95c.PNG)

Figure 7: All Stock Analysis for 2018

![VBA Challenge 2018 results](https://user-images.githubusercontent.com/78306719/109702772-48961300-7b5a-11eb-918f-54188ab0f610.PNG)

## Summary

**What are the advantages or disadvantages of refactoring code?** 

One advantage of refactoring code is the ability to expand the data being analyzed. For this project, the first analysis focused on the performance of DQ; However, after refactoring, the analysis expanded to 12 stocks. A second advantage of refactoring is the improvement made to the code making it easier for the reader to follow. Refactoring also has disadvantages. One disadvantage is that refactoring is time consuming. For example, the need to expand the analysis to other stocks increased the time to complete this project. Another disadvantage is the need for a budget increase because the project money can run out (Source: [https://stackoverflow.com/questions/43983284/what-are-the-advantages-and-disadvantages-of-refactoring-code-smell-in-software)].

**How do these pros and cons apply to refactoring the original VBA script?**

The pros and cons apply to refactoring the VBA script in that it was time consuming. For this project, the length of time to refactor increased compared to writing the original code. When comparing the codes, it is easier to follow the second code because more comments to explain the purpose of the code were included. The budget was not an issue for this project. 

