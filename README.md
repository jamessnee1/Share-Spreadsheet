# Share Spreadsheet Generator
This script pulls stock/share data from Yahoo Finance (including dividend data) and build an Excel spreadsheet. This is especially handy to track your investments.

## Modules required
There are a number of modules required in order to run the script. These are:
yfinance
pandas
openpyxl

You can install these with python pip:
```
pip install <module name>
```

## Usage
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
It may take a while to generate depending on the number of ticker symbols added to your data file. Once the script has generated the excel file it will open it.

## Todo
Currently, the script only will show the last dividend so estimated income is just based on this value (number of shares * last dividend). I'm planning on adding functionality to get all dividends in a financial year.
