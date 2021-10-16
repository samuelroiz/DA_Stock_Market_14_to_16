# DA_Stock_Market_14_to_16
Data Stock Market of 2014 through 2016

By excel VBA scripting, it analyzed the stock market data to help break down variables into small pieces to allow the reader to access clean data. Also, added variables to help understand which stock market is successful or unsuccessful. Created shades of coloring, maximum, mininum, and other codes to clean up the data.

## Back ground

Here is a picture of the outcome when it analyzes the stock market. As you can see, in the first column is the tickers. The tickers represent the stock market symbol. However, there are multiple of the same ticker. The next column over is the date. After the date column is the opening price of the stock. The three columns away from the opening price will be the closing price. The close price is when the closing price at the end of that year. One of the columns will be the yearly change from the opening price at the beginning of a given year to the closing price at the end of that year. The column will also have colors where red represents negative and green is positive. The upper right corner of the analyzed data will represent the highest volume, increase, and lowest decreased percentage of the following year.

![Stock Market Analyzed Example](https://github.com/samuelroiz/DA_Stock_Market_14_to_16/blob/main/Images/Stock_Market_Analyzed_Outcome_Example.png)

## Images of 2014, 2015, and 2016 outcome

### Stock Market Analyzed 2014 Outcome
![Stock Market Analyzed 2014 Outcome](https://github.com/samuelroiz/DA_Stock_Market_14_to_16/blob/main/Images/Stock_Market_Analyzed_2014.png)

![Stock Market Analyzed 2014 Outcome Continued](https://github.com/samuelroiz/DA_Stock_Market_14_to_16/blob/main/Images/Stock_Market_Analyzed_2014_2nd_Section.png)

### Stock Market Analyzed 2015 Outcome

![Stock Market Analyzed 2015 Outcome](https://github.com/samuelroiz/DA_Stock_Market_14_to_16/blob/main/Images/Stock_Market_Analyzed_2015.png)

![Stock Market Analyzed 2015 Outcome Continued](https://github.com/samuelroiz/DA_Stock_Market_14_to_16/blob/main/Images/Stock_Market_Analyzed_2015_2nd_Section.png)

### Stock Market Analyzed 2016 Outcome

![Stock Market Analyzed 2016 Outcome](https://github.com/samuelroiz/DA_Stock_Market_14_to_16/blob/main/Images/Stock_Market_Analyzed_2016.png)

![Stock Market Analyzed 2016 Outcome Continued](https://github.com/samuelroiz/DA_Stock_Market_14_to_16/blob/main/Images/Stock_Market_Analyzed_2016_2nd_Section.png)

## CODE in VBA Script

`Sub StockMarketVBA_Final()'
'Dim ws As Worksheet'
'For Each ws In Worksheets'

'ws.Cells(1, 9).Value = "Ticker"'
'ws.Cells(1, 10).Value = "Yearly Change"'
'ws.Cells(1, 11).Value = "Percent Change"'
'ws.Cells(1, 12).Value = "Total Stock Volume"'
'ws.Cells(2, 15).Value = "Greatest % Increase"'
'ws.Cells(3, 15).Value = "Greatest % Decrease"'
'ws.Cells(4, 15).Value = "Greatest Total Vol."'
'ws.Cells(1, 17).Value = "Value"'
'ws.Cells(1, 16).Value = "Ticker"`

Full Code found in --> https://github.com/samuelroiz/DA_Stock_Market_14_to_16/blob/main/VBA_SCRIPT_CODE.txt
