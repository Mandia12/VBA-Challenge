# VBA---Challenge

Purpose:

This project calculates and presents valuable information from data sets that represent stock values over a year's time for the years 2018, 2019, and 2020. These values found are Yearly Change, Percent Change, and Total Stock Volume for each stock. After calculating those values, the project finds which stock had the Greatest Percent Increase, Greatest Percent Decrease, and Greatest Total Volume.

Explanation:

Most of the code was written by myself with the help of previous exercises done in class. However, I derived some code from stackoverflow.com like the For loop that cycles through all the worksheets in lines 4-6 and line 127. Also, line 48 to create the percentage format for column "K". 

First, I created a For Loop to cycle through all rows of data in column "A".

Then I created an If Statement to check if the next value was different than the current one. 

If next value is different then I know I'm on the last row of a specific ticker symbol. With the code in the If statement, I find, store, then place Yearly Change, Percent Change, and Total Stock Volume for each indiviual ticker symbol into a summary table. I then formatted the summary table. 

After the summary table is created, I created another For Loop to cycle through the summary table to compare the values for each ticker symbol.

From the summary table I extract and place the stocks with the Greatest Percent Increase, Greatest Percent Decrease, and Greatest Total Volume.

Finally, I contained my script with the For Loop the cycles through each of the worksheets. 

References:

https://stackoverflow.com/questions/43738802/how-to-apply-vba-code-to-all-worksheets-in-the-workbook

https://stackoverflow.com/questions/12801884/vba-format-columns


