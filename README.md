### stock-analysis
VBA macros in excel procedurally guiding representations of percentages and result of 2017 and 2018 stock.
### *Project Overview*
The purpose of this analysis according to Module 2 was to create a VBA Macro that could trigger pop-ups and inputs, read and change cell values, and format cells. Additionally, I was to create a for loop and conditionals as well as nest them for direct flow. This was to exhibit application of coding skills such as syntax recollection, pattern recognition, learn how to debug potential problems faced via problem decompisition.  
### *Results*
Using a copy of the analysis, I put two charts side by side to represent the data as seen here:
![2017_2018_side_by_side](https://user-images.githubusercontent.com/96352415/150004528-2c40f23b-f103-4d8a-8c07-2a21a3c3b861.png)
An issue I ran into when reformatting information from the original datasets is that the positive Returns for some reason kept displaying positive returns as red. Despite this minor setback, we can see within the two graphs that 2017 was an overall more successful year than 2018 for stocks and ultimately the return in 2018 was harshly impacted.
Prior to converting the refactored code, I noticed that my run time averaged under a second. The code for that was

![original_code](https://user-images.githubusercontent.com/96352415/150009361-92dfdb08-59af-49d6-a551-f509e6d54460.png) 

whereas the refactored looks like this.

![refactored_1](https://user-images.githubusercontent.com/96352415/150009430-29727991-7a3a-4ca7-b83b-7d6fff6ad3f4.png)
![refactored_2](https://user-images.githubusercontent.com/96352415/150009463-2515bbb6-0bc6-4127-9158-25f66b9bf3d7.png)

The original code has nested loops which I feel contribute to a more streamlined process rather than repetitive data which may slow down processing time.

If you take a look here,
![stocks_analysis_2017](https://user-images.githubusercontent.com/96352415/149973265-78649335-b5fe-4da1-8117-9e9dc62ab0f1.png)
![2018_analysis](https://user-images.githubusercontent.com/96352415/149973318-a293838a-7367-4c1f-afa2-cec564854263.png)

we can see that it added an additional 1-3 seconds to process. 
### Summary
I certainly feel that refactoring code can get repetitive and create more problems than necessary. I had the most success adding on to one continuous portion of the coding as far as time taken, as well as being able to read and understand my process and make connections of all the active functions and how they were facilitated, vs making a separate module and doing more work to achieve the same thing. However, simplifying your code without nesting does help someone viewing it for the first time to see the individual functionality of each line, and why it's there.

I do feel that it could support the original script by adding layers that ultimately enhance the original functions, despite being repetitive. 
