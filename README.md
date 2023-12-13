# VBA-Challenge
## Module 2 Challenge - VBA

________________________________________

### Background

The joy of VBA is that it takes the tediousness out of repetitive tasks with the click of a button.  I have used VBA scripting to analyse generated stock market data in this challenge.
Two excel files were provided for this challenge:

<img src = 'https://github.com/Mago281/VBA-Challenge/assets/131424690/e4c520ca-6b5f-4cce-8fed-9f9b2218328a' width = '170' height = '60'>

________________________________________

### Steps undertaken

**1.  Retrieval of Data & Column Creation**

    Created a script that looped through all the stocks for one year and read/stored the following values from each row:

    -  ticker symbol

    -  total stock volume

    -  Columns 1 and 2 are populated with the values for the ticker: Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.

    -  The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
  

**2.  Conditional Formatting**

  Used conditional formatting in Column J (yearly change) to highlight positive changes in green and negative changes in red.


**3.  Calculated Values**

  Ensured that all three of the following values were calculated correctly and displayed in the output:
  
    *  `Greatest % Increase`
    
    *  `Greatest % Decrease`
    
    *  `Greatest Total Volume`


**4.  Looping Across Worksheet**

  The VBA script can ran on all sheets successfully.
    

The result is as follows:
 

________________________________________

### Bonus

Added functionality to the script to return the stock with the `"Greatest % increase"`, `"Greatest % decrease"`, and `"Greatest total volume"`. 

The solution match the following image:

1.	Values were populated in columns P and Q.

2.	Number formatting was applied to relevant cells in column Q.

3.	Headings were added to columns O, P and Q.

4.	Formatting (bold text and autofit width) were applied to the results.

5.	Steps 1 to 8 were then repeated on all worksheets.

 

Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.

________________________________________

### Other Considerations

•	Used the sheet _`alphabetical_testing.xlsx`_ while developing the code.  This dataset was smaller and allowed faster testing. The code ran on this file in under 5 minutes.

•	Made sure that the script acted the same on every sheet.



