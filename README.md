# VBA-Challenge
## Module 2 Challenge - VBA

________________________________________

### Background

The joy of VBA is that it takes the tediousness out of repetitive tasks with the click of a button.  I have used VBA scripting to analyse generated stock market data in this challenge.
Two excel files were provided for this challenge:

<img src = 'https://github.com/Mago281/VBA-Challenge/assets/131424690/e4c520ca-6b5f-4cce-8fed-9f9b2218328a' width = '170' height = '65'>

________________________________________
### Steps undertaken


Created a script that looped through all the stocks for one year and output the following information:

1.	The ticker symbol in the XXX file loops through all the stocks for one year to show the following:

•	Columns 1 and 2 are populated with the values for the ticker: Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.

•	The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.

•	The total stock volume.


2.	Used conditional formatting in Column J (yearly change) to highlight positive changes in green and negative changes in red.


The result is as follows:
 

________________________________________

### Bonus

Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". The solution should match the following image:
1.	values are populated in columns P and Q.
2.	Number formatting is applied to relevant cells in column Q.
3.	Headings are added to columns O, P and Q.
4.	Formatting (bold text and autofit width) are applied to the results.
5.	Steps 1 to 8 are then repeated on all worksheets.

 

Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.

________________________________________

Other Considerations

•	Use the sheet alphabetical_testing.xlsx while developing your code. This dataset is smaller and will allow you to test faster. Your code should run on this file in under 3 to 5 minutes.

•	Make sure that the script acts the same on every sheet. The joy of VBA is that it takes the tediousness out of repetitive tasks with the click of a button.

•	Some assignments, like this one, contain a bonus. It is possible to achieve proficiency for this assignment without completing the bonus, but the bonus is an opportunity to further develop your skills and receive extra points for doing so.


