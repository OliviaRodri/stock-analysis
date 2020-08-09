# stock-analysis
stock analysis # 

   
## Overview of Project 
### Purpose

The purpose of this analysis was to refracture a script for Steve and his parents who are his first time clients. Steve performed a stock analysis to find the return on a green company his parents invested in, DaQu New Energy Corp. Steve then wanted to do the same stock analysis on several green companies. He requested the script to be refactored to analyze several stocks. Also, the script was to run quicker and more efficient . In order to accomplish this task, the original code was refactored.
	
## Results
### Results
There is a clear contrast in stock performance between the years 2017 and 2018. From 2017 to 2018, there were only two tickers that resulted in a positive return, as shown below.
	
![Image_Here](Resources/2017analysis.PNG)

![Image_here](Resources/2018Analysis.PNG)




### Summary
Summary: In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?
Historical data suggests that if the monetary goal is kept below $5000, there is about a 75% chance of success. On the other hand, if the goal is set at $5000 or higher, but less than $15000, then there is about a 55% chance of success. This analysis indicates that Louise's planned goal of $12000 would have about a 50% likelihood of success. 

![LINE_GRAPH_HERE](Resources/Outcomes_vs_Goals.png)

### Challenges and Difficulties Encountered

My challenges and difficulties were in the nested functions. The formulas became large in the outcomes based on goals excel sheet. At first, I was hard coding some criterias in the formulas which was tedious and time consuming. I then created intermediate cells that I could reference in the formulas. This made it much easier to complete the analysis. 

## Results 
Two conclusions about the Outcomes based on Launch Date are:

- Based on the launch date historical data, the month of May is the best time for Louise to launch her crowdfunding campaign for her play, Fever. Plays launched in May have had the highest percentage of being successful and May also has had the highest number of campaigns launched. 

- The second conclusion is for Louise to avoid launching a crowdfunding campaign in the month of December, it has had the lowest percentage of success.


 What can you conclude about the Outcomes based on Goals?

- The conclusion for the outcomes based on goals, is that historical data shows that plays with a goal between $5000 and $15000 have had a %54 success rate. In Louise's case, her planned goal of $12000 would likely have about a 50% rate of success.

What are some limitations of this dataset?

- The donation amount per doner is unknown, therefore using the average donation can be masking skewed distributions. Knowing the donation amount per donor could provide insight into how to advertise the campaign.

- There was no data on how the kickstarter was promoted. For example, through social media platform, email, mail, etc. This data would reveal which promoting method has had the highest level of success.


What are some other possible tables and/or graphs that we could create?

- A pivot table that is like the outcome based on launch date, with the addition of a filter for the monetary range. 
- Another sheet like "the outcomes based on goals" that can show the same results by month, rather than for the whole dataset time frame.



