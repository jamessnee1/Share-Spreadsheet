import os
import sys
import time
import datetime
import yfinance as yf
import math
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

AUS_TICKERS = {
"ANZ.AX" : 48,
"APA.AX" : 45,
"APT.AX" : 76,
"ARF.AX" : 230,
"BOQ.AX" : 77,
"BSL.AX" : 188,
"CBA.AX" : 29,
"DTS.AX" : 3333,
"FMG.AX" : 200,
"GEM.AX" : 500,
"PL8.AX" : 1351,
"RMD.AX" : 20,
"VHY.AX" : 9,
"Z1P.AX" : 262
}

#Debugging
#print(AUS_TICKERS)


start = datetime.datetime.now()

#Create excel
wb = Workbook()  
ws = wb.active
strDate = start.strftime("%d_%m_%Y")
dest_filename = 'SharePortfolio_' + strDate + '.xlsx'
ws.title = "Shares " + start.strftime("%d_%m_%Y")
#Write excel column headers
ws.column_dimensions['A'].width = 55
ws.column_dimensions['B'].width = 25
ws.column_dimensions['C'].width = 25
ws.column_dimensions['D'].width = 25
ws.column_dimensions['E'].width = 25
ws.column_dimensions['F'].width = 25
ws.column_dimensions['G'].width = 25
ws.column_dimensions['H'].width = 25
ws.column_dimensions['I'].width = 25
ws['A1'] = "Company Name"
ws['B1'] = "Sector"
ws['C1'] = "Current Market Price"
ws['D1'] = "Number of shares"
ws['E1'] = "Total share value"
ws['F1'] = "Last dividend per share"
ws['G1'] = "Dividend yield"
ws['H1'] = "Ex-dividend date"
ws['I1'] = "Estimated Income"
ws['A1'].font = Font(bold=True)
ws['B1'].font = Font(bold=True)
ws['C1'].font = Font(bold=True)
ws['D1'].font = Font(bold=True)
ws['E1'].font = Font(bold=True)
ws['F1'].font = Font(bold=True)
ws['G1'].font = Font(bold=True)
ws['H1'].font = Font(bold=True)
ws['I1'].font = Font(bold=True)
    
print("Downloading stock data from Yahoo Finance...")
rowCount = 2

totalDividendIncome = 0.00
totalNumOfShares = 0
totalPortfolioValue = 0.00

for company in AUS_TICKERS:
    print("Getting data for " + company)
    ticker = yf.Ticker(company)
    #print("next event: " + str(ticker.calendar))
    
    try:
        sector = str(ticker.info['sector'])
    except KeyError: sector = "N/A"
    
    longName = str(ticker.info['longName'])
    lastDividendValue = str(ticker.info['lastDividendValue'])
    marketPrice = str(ticker.info['regularMarketPrice'])
    dividendYield = str(ticker.info['dividendYield'])
    
    if dividendYield == 'None':
        dividendYield = 0.00
    
    if lastDividendValue == 'None':
        lastDividendValue = 0.00
        
        
    estimatedIncome = round(float(AUS_TICKERS[company]) * float(lastDividendValue))
    totalDividendIncome = round(totalDividendIncome + estimatedIncome)
    
    totalShareValue = round(float(AUS_TICKERS[company]) * float(marketPrice))
    totalNumOfShares = round(totalNumOfShares + AUS_TICKERS[company])
    totalPortfolioValue = round(totalPortfolioValue + totalShareValue)
    
    #Data from API
    ws.cell(row=rowCount, column=1, value=str(longName))
    ws.cell(row=rowCount, column=2, value=str(sector))
    ws.cell(row=rowCount, column=3, value="$" + str(marketPrice))
    ws.cell(row=rowCount, column=4, value=str(AUS_TICKERS[company]))
    ws.cell(row=rowCount, column=5, value="$" + str(totalShareValue))
    ws.cell(row=rowCount, column=6, value="$" + str(lastDividendValue))
    yieldColumn = ws.cell(rowCount, 7)
    yieldColumn.number_format = '0.00%'
    yieldColumn.value=str(dividendYield)
    #ws.cell(row=rowCount, column=7, value=str(dividendYield))
    ws.cell(row=rowCount, column=9,value="$" + str(estimatedIncome))
    rowCount = rowCount + 1

#Write totals to spreadsheet
ws.cell(row=rowCount, column=1, value=str("TOTAL:"))
ws.cell(row=rowCount, column=4, value=str(totalNumOfShares))
ws.cell(row=rowCount, column=5, value="$" + str(totalPortfolioValue))
ws.cell(row=rowCount, column=9,value="$" + str(totalDividendIncome))

print("Saving spreadsheet...")
wb.save(filename = dest_filename)

end = datetime.datetime.now()
elapsed = end - start

print("Script took " + str(elapsed) + " to execute")