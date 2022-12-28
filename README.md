# VBAchallenge
Bootcamp Challenge 2: VBA Complete

In this challenge we used VBA to analyze stock market data. Each company was represented by a several letters in the "ticker" column along with daily highs, lows, opening, and closing values. Our job was to use VBA to determine the yearly change, percent change, and total stock volume for each company and report those values seperately along with the ticker label.

To do this I defined my vairables and used a loop to cycle through each row of data adding the total volume along the way. Once the program reached the end of one company's ticker it reported the change between the first opening value and the final closing value for that company, caluclated the percent changed, and reported the total volume next to the ticker label on the first line of the new chart. The yearly change column then applied green formatting for positive change values, red formatting for negative change values, and yellow for no change. Finally, he percent change column was formatted as a precntage with two decimal places. The program then reset and did the same thing for each company until reaching the final row of data. Finally, I assigned the appropriate column headers to my new data.

To create the final chart showing the greatest precent increase and decrease and greatest total volume I used another looping function. Instead of looking for the change in ticker values this loop retrieved the largest (or smallest) value by comparing each value to the previous largest (or smallest) value in the columns containing the summary data. Then reported the greatest percent increase and decrease and greatest total volume in each column and the corresponding ticker value. Again, assigning proper headings and number formatting as needed.

The new data is much easier to utlize, leaving only the yearly calculations rather than sifting through hundreds of lines of data for each company. It is also faster and more reliable than creating the same analysis the data by hand.

I have attached three screen shots, one for each year of data of the Multipule Year Stock Data spreadsheets, as well as a copy of my VBA script and the Alphabetical Data Listing as an working example.
