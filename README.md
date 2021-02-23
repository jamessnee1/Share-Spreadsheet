# Share Spreadsheet Generator
Pull stock/share data from Yahoo Finance (including dividend data) and build an Excel spreadsheet. 

Usage:
Add ticker symbols for companies by adding them to a csv file, including number of shares (I have added an example CSV):
```
ANZ.AX,100
MSFT,50
TSLA,10
```
Any ticker symbol on Yahoo Finance can be used.
Then, run the script by running the following:
```
python generateShareSpreadsheet.py
```
It may take a while to generate depending on the number of ticker symbols added to your data file.
