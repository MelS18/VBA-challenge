# VBA-challenge
'Module assignemnt text:
Create a script that loops through all the stocks for one year and outputs the following information:
 The ticker symbol
 Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
 The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
 The total stock volume of the stock. The result should match the following image:
![image](https://github.com/MelS18/VBA-challenge/assets/129136787/48ec1cb4-003d-488e-aa71-c79bdedac5ce)
' Add functionality to your script to return the stock with the:
"Greatest % increase", "Greatest % decrease", and "Greatest total volume".
-----------------------------------------------------------------------------------------------------------------------------------------

How the code works: 

1. Counts number of worksheets 

2. I have to define the variables  as: 
 Dim Greatest_Increase As Double
    Dim Greatest_Decrease As Double
    Dim greatest_name As String
    Dim decrease_Name As String
    Dim Greatest_total_volume As Double
    Dim Greatest_total_volume_name As String

3. Create a loop for all sheets

4. We  are considering Stock_details(0)as opening price, Stock_details(1)as closing price and Stock_details(2)as the volumen of stocks and define de Columns names

5. Ones the loop move around the sheet, I determinate yearly change of the stock, reset variables for next stock and increment row counter

6. Add to the volume of the stock

7. Determinate opening price for every stock

8. Color codes yearly change using the Excel color code
9. Determinate stock with greatest percent increase, greatest percent decrease, and greatest total volume using MAX and MIN funtions
10. Print the results 
