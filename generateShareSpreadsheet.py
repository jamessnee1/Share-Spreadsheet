import os
import sys
import time
from datetime import datetime
import yfinance as yf
import math
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

#Format of stock file: Ticker Symbol, Number of shares

def loadTickerFile(fileName):
    TICKERS = {}
    print("Loading stocks file...")
    with open(fileName, "r") as fileHandler:
        for line in fileHandler:
            fileLine = line.split(",")
            TICKERS[fileLine[0]] = int(fileLine[1])
        fileHandler.close()
    return TICKERS

def writeTitle(ws, title, width, column):
    cell = column + "1"
    ws.column_dimensions[column].width = width
    ws[cell] = title
    ws[cell].font = Font(bold=True)
    

def writeData(ws, row, column, value):
    ws.cell(row=row, column=column, value=str(value))
    
AUS_TICKERS = loadTickerFile("exampleStockFile.csv")
#Debugging
#print(AUS_TICKERS)

start = datetime.now()

#Create excel
wb = Workbook()
wb.guess_types = True
ws = wb.active
strDate = start.strftime("%d_%m_%Y")
dest_filename = 'SharePortfolio_' + strDate + '.xlsx'

#If file exists, delete it
if os.path.exists(dest_filename):
    os.remove(dest_filename)
    
ws.title = "Shares " + start.strftime("%d_%m_%Y")
#Write excel column headers
writeTitle(ws, "Company Name", 55, "A")
writeTitle(ws, "Ticker Symbol", 12, "B")
writeTitle(ws, "Sector", 25, "C")
writeTitle(ws, "Current Market Price (ask)", 25, "D")
writeTitle(ws, "Number of shares", 25, "E")
writeTitle(ws, "Total share value", 25, "F")
writeTitle(ws, "Last dividend per share", 25, "G")
writeTitle(ws, "Dividend yield", 25, "H")
writeTitle(ws, "Ex-dividend date", 25, "I")
writeTitle(ws, "Estimated Income", 25, "J")
    
print("Downloading stock data from Yahoo Finance...")
rowCount = 2

totalDividendIncome = 0.00
totalNumOfShares = 0
totalPortfolioValue = 0.00

for company in AUS_TICKERS:
    print("Getting data for " + company)
    ticker = yf.Ticker(company)
    #print("next event: " + str(ticker.calendar))
    #print(ticker.dividends)
    
    try:
        sector = str(ticker.info['sector'])
    except KeyError: sector = "N/A"
    
    longName = str(ticker.info['longName'])
    lastDividendValue = str(ticker.info['lastDividendValue'])
    marketPrice = str(ticker.info['ask'])
    dividendYield = str(ticker.info['dividendYield'])
    
    if dividendYield == 'None':
        dividendYield = 0.00
    
    if lastDividendValue == 'None':
        lastDividendValue = 0.00
        
        
    estimatedIncome = round(float(AUS_TICKERS[company]) * float(lastDividendValue), 3)
    totalDividendIncome = round(totalDividendIncome + estimatedIncome, 3)
    
    totalShareValue = round(float(AUS_TICKERS[company]) * float(marketPrice), 3)
    totalNumOfShares = round(totalNumOfShares + AUS_TICKERS[company])
    totalPortfolioValue = round(totalPortfolioValue + totalShareValue, 3)
    
    try:
        exDate = ticker.calendar['Value'][0]
    except KeyError: exDate = "N/A"
    except TypeError: exDate = "N/A"
    
    if exDate != "N/A":
        #convert to proper date
        dateComponents = str(exDate).split(" ")
        exDate = dateComponents[0]
    
    
    #Data from API
    writeData(ws, rowCount, 1, longName)
    writeData(ws, rowCount, 2, company)
    writeData(ws, rowCount, 3, sector)
    writeData(ws, rowCount, 4, float(marketPrice))
    writeData(ws, rowCount, 5, str(AUS_TICKERS[company]))
    writeData(ws, rowCount, 6, float(totalShareValue))
    writeData(ws, rowCount, 7, float(lastDividendValue))
    writeData(ws, rowCount, 8, float(dividendYield))
    writeData(ws, rowCount, 9, exDate)
    writeData(ws, rowCount, 10, float(estimatedIncome))
    rowCount = rowCount + 1

#Write totals to spreadsheet
writeData(ws, rowCount, 1, "TOTAL:")
writeData(ws, rowCount, 5, totalNumOfShares)
writeData(ws, rowCount, 6, float(totalPortfolioValue))
writeData(ws, rowCount, 10, float(totalDividendIncome))

#Save spreadsheet
wb.save(filename = dest_filename)

end = datetime.now()
elapsed = end - start

print("Script took " + str(elapsed) + " to execute")

#Open excel file
os.startfile(dest_filename)