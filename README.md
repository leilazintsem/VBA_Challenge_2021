# VBA_Challenge_2021

## VBA Homework – Stock Market

### Objectif

This Repository has been created for the submission of the VBA homework assignment for my Data Science and Visualization Bootcamp at Northwestern University.

VBA Scripting

Student: Leila Zintsem a Moute

March 2021

---

### About the project

The following repository displays my solution to the Homework assignment VBA_Challenge. This repository is structured in different folders:

* The first folder “Images” contains the screenshots of the project solutions.
  The pictures Titled “2014”, “2015”, “2016” are the screenshots solution of the Multiyear stock dataset. The other ones (A, B, C, D, E, F, P) are screenshots for the testing dataset “Alphabetical testing”
* The second folder VBA_script has my VBA script which I used for both the Alphabetical testing dataset and the Multi year stock dataset.
* The Third folder “sources” contains the two datasets used for this project.

Through this assignment, I hope you will be able to see the technical skills I have gained throughout this section of the Data Science and Visualization bootcamp.

---

### The Project Description

For this Project, I was asked to use VBA scripting to analyze real stock market data. I was given Two datasets:
the alphabetical testing dataset, which I used to test out my code and the Multiyear stock dataset where I applied the code tested on the alphabetical testing dataset. These datasets were not uplaoded to this repository because they are too large. Howerver, the following information were provided:

* Stock Ticker symbol
* Date
* Open price
* The stock high and low
* closing price
* Volume  of the stock

---

##### Project challenges

1. Create a script that will loop through all the stocks for one year and output the following information

* Ticker symbol
* Yearly change from opening price at the beginning of a given year to the is closing price at the end of that year
* The percent change from opening price at the beginning of a given
  year to the closing price at the end of that year
* The total stock volume of the stock

2. Conditional format changes using VBA script to highlight positive change in green and negative change in red.
3. Create a summary table that will show the return of the stock "Greatest % increase", "Greatest % decrease" and "Greatest total volume".
4. Adjust VBA scrip so that it can be run on every worksheet just by running the VBA Code once.

   ---

##### Project outcomes

---



###### VBA script Explanation

The first step in my VBA script was to loop throughout all the worksheets. Then for all the worksheets, I scripted a code to Insert the summary tables where my output will be displayed.
Secondly, I declared all my variables. You will see that Open price was assigned value before the For loop. This was done so that the value within cells (2,3) get grabbed before the loop starts, otherwise the loop would have grabbed the wrong Open price. Moving forward, I looped throughout all the tickers. The loop starts on row two (i=2), that is because the first ticker is listed on row two.
Inside the i for loop, I have set two conditions:


•	The first condition is set so that the program knows when to pause the loop to apply the conditions; Inside this  condition, I nestled another IF condition to conditional format the yearly change.

•	The second IF condition was added so that the program ignores all the bad data containing 0 value in them. This step is important because if this condition is not set, we will get an overflow error message. The program will try to divide yearly change by open price values containing the 0 and that will be an overflow error.

At this point on the for loop, we are on the last row of the first ticker and we want the program to grab a new open price when it goes on to next ticker, therefore we set a new open price, which is the open price of the next ticker. Moving forward, I wrote a script to add the total volumes for each ticker as well as the script to print ticker symbol, total volume, yearly change, and percent change. The for-loop end with an else condition where Volumes are added if the first “IF statement” is false.
The second part of the Script focus on the second table where the max/min Percent change and Max volume are displayed
For this part I declared my variables, then I used a for loop to loop through the tickers in the first table. Inside the loop I set the Conditions with the IF and else IF statements, so the max/min value are grabbed, format, and printed on the second table.
My code ends with a “next i” to move onto the next row, and “next ws” to move on to the next worksheet.


As you can see, I was able to write the VBA script and resolve the assignment challenges.


###### References:

Excel VBA: Color Index Codes List & RGB Colors. Retrieved from Automatedexcel.com
https://www.automateexcel.com/excel-formatting/color-reference-for-color-index/
Macro to Loop Through All Worksheets in a Workbook. Retrieved from Support.microsoft.com
https://support.microsoft.com/en-us/topic/macro-to-loop-through-all-worksheets-in-a-workbook-feef14e3-97cf-00e2-538b-5da40186e2b0
Help with avoiding division by zero error in VBA. Retrieved from MrExcel.com
https://www.mrexcel.com/board/threads/help-with-avoiding-division-by-zero-error-in-vba.783862/
