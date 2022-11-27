# VBA-challenge - The VBA of Wall Street
There are two VBA Scripts for analyzing the stock from 2018 – 2020

## Introduction of VBA Scripts of Wall Street
1. Stock_Performance
 The VBA Script will output the following information:
      . Yearly Change: calculate the difference between the opening price at the beginning and the closing price at the end of each year.
      . Percent Change: show the difference of yearly change in a percentage format. 
          The formula is Yearly Change / the opening price * 100%
      . Total Stock Volume: Calculate the total stock volume by each ticker symbol each year.
      . Conditional Formatting: For better visualization, the negative yearly change has been highlighted in red color, and the positive yearly changes were highlighted in green color.
      
2. Stock_Greatest 
The VBA Script shows the Max and Min values of the difference each year and indicates which ticker symbol achieves its Max or Min value each year. The script also demonstrates which ticker symbol award the maximum total stock volume each year.

## Steps to run VBA Scripts – Stock_Performance.vbs
1.	Open the original excel file ”Multiple_year_stock_data”
2.	Click “Developer” –> “Visual Basic” -> “File” in Microsoft Visual Basic for Applications -> “Import File”
3.	Import “Stock_Performance.vbs” to the Multiple_year_stock_data.xlsx file and run the module.
Notes: Due to the huge volume of data, this may take 5-8 minutes to complete the calculation by using the VBA Scripts

## Steps to run VBA Scripts – Stock_Performance.vbs
1.	Please run the VBA Scripts – Stock_Performance.vbs firstly
2.	Import “Stock_Greatest.vbs” to the Multiple_year_stock_data.xlsx file and run the module.
Notes: The VBA Script may take about 2 minutes to complete its calculation.
