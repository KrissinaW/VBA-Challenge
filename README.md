# VBA-Challenge

In this project we were assgined use VBA scripting to analyze stock market data by creating a script that loops through all the stocks in a dataset for one year and outputs the following data: 

- Ticker Sybol
- Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
- The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
- The total stock volume of the stock. The result should match the following image:

We were also required to add functionality to the script that returns the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". 
It was requested that we make the appropriate adjustments to the VBA script to enable it to run on every worksheet in the workbook (that is, every year) at once.
To test our code initially we needed a smaller dataset so that the macro would take less time and to allow us to debug the code if there were any issues. We used a file named alphabetical_testing.xlsx while developing your code. 
we were required to validate data by ensuring that the script acted the same on every sheet. 

What you can find in this Respoitory: 
The script used to automate the formatting of the workbook.   
The script file is called Stock_Challenge

I was able to get the Test Script to work and producce the result as shown on the module example.  I used CHAT GPT to assist me in debugging the Greatest % Decrease since I was running into the issue of my code not returning any negitive numbers.  I was able to pinpoint to problem when I noticed with the help of Chat GPT that I was using a formula incorectly and it was returning the absolute value instead of the actual lowest percent in the range. 

In this file you will be able to locate the screen shots for years 2018, 2019, and 2020. 
