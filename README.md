# VBA Challenge

## Overview

This project involves analyzing stock data over a year's time for the years 2018, 2019, and 2020. The objective is to calculate key metrics for each stock, such as Yearly Change, Percent Change, and Total Stock Volume. Additionally, the project identifies the stocks with the Greatest Percent Increase, Greatest Percent Decrease, and Greatest Total Volume.

## Dataset

The data used in this project includes stock values for the years 2018, 2019, and 2020. The dataset comprises multiple worksheets, each containing the stock data for a specific year.

## Features

* Yearly Change Calculation:
Computes the Yearly Change for each stock by comparing the opening price at the start of the year to the closing price at the end of the year.
* Percent Change Calculation:
Calculates the Percent Change for each stock to measure the relative change in stock value over the year.
* Total Stock Volume Calculation:
Sums up the total stock volume traded over the course of the year for each stock.
* Identification of Extreme Values:
Determines which stocks had the Greatest Percent Increase, Greatest Percent Decrease, and Greatest Total Volume across the entire dataset.

## Steps

1. **Data Processing:**
   * A for loop cycles through each row of data in the worksheets.
   * An if statement checks for the transition between different ticker symbols to calculate and store Yearly Change, Percent Change, and Total Stock Volume.
     
2. **Summary Table Creation:**
   * After processing the data, a summary table is created to store the calculated metrics for each stock.
   * The summary table is formatted for clarity.
     
3. **Value Extraction:**
   * A second for loop cycles through the summary table to identify and store the stocks with the most significant percent changes and total volumes.
     
4. **Multi-Sheet Analysis:**
   * The entire script is contained within a loop that cycles through each worksheet, allowing for the analysis of data across different years.
  

## Tools

* Visual Basic: For scripting and data processing.
* Excel: The source of the stock data worksheets.
  
## How to Use

1. Ensure that the stock data for 2018, 2019, and 2020 is available in separate worksheets.
2. Run the Python script to perform the analysis.
3. The output will include a summary table and the identified stocks with extreme values.
