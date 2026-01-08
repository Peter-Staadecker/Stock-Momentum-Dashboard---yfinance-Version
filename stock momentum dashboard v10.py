# This program is not investment advice nor is it guaranteed to be correct or accurate. This program uses
# the python yfinance library to download Yahoo data. This program is intended for personal use only.
# You should refer to Yahoo!’s terms of use for details on your rights to use the actual data downloaded.

#  This program uses python and the yfinance library to download adjusted closing prices for various selected stock
# and ETF prices from Yahoo for the last previous trading day. In addition to the closing prices
# the program downloads the closing prices for 1, 3, 6 and 12 months ago. [These are by the same date as best possible
#  within each of those previous months, regardless of the length of the months.] The prices from Yahoo are
#  dividend-adjusted. The program then calculates the ANNUALIZED percentage price changes from each of these
# prior dates until the most Recent Price. [The annualization is done by raising the price change
# to the power of 12 divided by the number of elapsed months, without regards to the length of each month.]
# The 1, 3, 6 and 12-month changes give a colour-coded "momentum" dashboard for stock price changes.
# The program saves the dashboard to an Excel file and opens the file for view.
# The inputs to the program - the list of stock tickers and the path and name of the output file - are hard coded
# in the python program, and can be changed there by the user. Note that any stock or ETF that has a Yahoo listing,
# regardless of whether Canadian or US or other can be included.


# Some additional technical details:
# yfinance retrieves stocks by specifying a ticker name or list of tickers and a date range. If you provide a
# single date, yfinance will return #NA or None if there was no trading on that date. This program
# therefore provides a range of 5 days close to the target date, and yFinance will find at least one date with closing
# data within that range. It was tempting to let yfinance do this for the entire list of stocks all at once, but
# that creates a second problem. Example: On Sunday 28'th December I asked yFinance to retrieve closing prices for a
# batch of stocks including US listings (e.g. AAAU) and Canadian listings (e.g. CNQ.TO). The most recent close for AAAU
# was Friday 26 December. However the Canadian exchanges were closed on that day so yfinance retrieved a valid
# closing price for AAAU and #NA for CNQ.TO. To avoid this problem I let yfinance look up each stock separately, rather
# than in a combined list of stocks. That way I got the close for AAAU from 26 December, and for CNQ.TO from
# 24 December.

# v8 - test for the existence of an output file, and instead of laboriously updating it, delete it and write it anew.
# v9 - instead of hardcoding the tickers, take them from the output spreadsheet, and let the spreadsheet include
#      a last purchase date and price. Add a calculation for the annualized % price change since last purchase.
# v10   - find the most recent ticker price rather than yesterday's close

import yfinance as yf
import pandas as pd
from datetime import datetime # for timestamping the file name
from datetime import date
from dateutil.relativedelta import relativedelta # for calculating 1,3,6,12 months date before today
import os
import platform
import subprocess

# -- make sure all pandas dataframe columns are shown when printing to console
pd.set_option('display.max_columns', None)   # show all columns
pd.set_option('display.float_format', '{:.2f}'.format) # 2 decimal places
pd.set_option('display.width', 200)         # don't wrap lines
pd.set_option("display.max_colwidth", 15)  # max sub column width
pd.set_option('display.float_format', '{:.2f}'.format) # print data frames to two decimals

# - validate yfinance version if validation is needed
# - print(yf.__version__)

#-----------------------------------------------------------------------------------
# function
def get_price_on_or_before(ticker: str, targetdate_str: str):
    # Fetch the closing price for one ticker on a given date.
    # Note that by design, when calling yfinance with start date s and end date e, and daily data,
    # yfinance returns days s to e-1 inclusive, not e.
    # and you can't work around it by putting in tomorrow's date for e, and hope to get today's data
    # unless today's markets have already closed.
    # The function will return the closing price for the ticker, on the date on or before the given date
    # according to when a closing price was last available. The function will check for an available
    # closing price up to 5 days prior. E.g. on Monday 29-Dec 2025, the most recent TSX closing date will be
    # Wed 24 December. Thurs 25 is Xmas, Fri 26 is Boxing Day, and no closes on the weekend.

    # Convert to date object
    targetdate_dt = datetime.strptime(targetdate_str,"%Y-%m-%d").date()

    # Try up to 4 days earlier if necessary
    targetdate_minus5dys_dt = targetdate_dt - relativedelta(days=5)
    targetdate_minus5dys_str = targetdate_minus5dys_dt.strftime("%Y-%m-%d")

    historyDF = yf.download(ticker, start = targetdate_minus5dys_str, end = targetdate_str, interval= "1d")

    # usually the last row of historyDF will be a date with valid date, though it may be earlier than the
    # target date.
    # I run these one ticker at a time, because e.g. US tickers are available on 26 December,
    # Canadian tickers will not be available on the 26'th, but only up to the 24'th.

    ln=len(historyDF)
    # here's the date for the last row of historyDF
    returndate_dt = (historyDF.index[ln-1])
    close_price = historyDF.loc[returndate_dt,("Close", ticker)]
    return returndate_dt, close_price


# --------------------------------------------------------------------------------------------------
# - Main Program

# Set up yesterday's date and 1,3,6, 12 months prior
# Note that by design, when calling yfinance with start date s and end date e, and daily data,
# yfinance returns days s to e-1 inclusive, not e.
# and you can't work around it by putting in tomorrow's date for e, and hope to get today's data
# unless today's markets have already closed



# hard code the Excel file path and name to use
file_name = r"D:\Users\rpsco\Documents\software\python\python stock screeners\stock momentum dashboard for github\stock momentum dashboard v10.xlsx"

# read the stock tickers from this file
tickersDF = pd.read_excel(file_name, header=2, usecols="A")
tickersDF.columns=['Tickers']
tickersLST = tickersDF["Tickers"].tolist()

# read the tickers, last purchase date and price from this file using tickers as the index
purchaseDF = pd.read_excel(file_name, header=2, usecols="A, F, G")
purchaseDF = purchaseDF.set_index("Tickers") # Tickers is the Excel column header for column A
purchaseDF.columns = ['Last Purchase Date', 'Last Purchase Price']


date_today=date.today()
# ==========================================================================================

# note that relativedelta(months=x) goes back x months and picks the same date in that month
# if possible, it does not go back a fixed number of days. Going back one month from March 15
# would be Feb 15, regardless of the fact that Feb is a shorter month.

date_minus_1=date_today-relativedelta(months=1)
date_minus_3=date_today-relativedelta(months=3)
date_minus_6=date_today-relativedelta(months=6)
date_minus_12=date_today-relativedelta(months=12)

# set up the string equivalent of the dates for calling from yfinance
date_todayString= date_today.strftime("%Y-%m-%d")
# date_yesterdayString= date_yesterday.strftime("%Y-%m-%d")
date_minus_1Str= date_minus_1.strftime("%Y-%m-%d")
date_minus_3Str= date_minus_3.strftime("%Y-%m-%d")
date_minus_6Str= date_minus_6.strftime("%Y-%m-%d")
date_minus_12Str= date_minus_12.strftime("%Y-%m-%d")

# The resultDF dataframe will hold stock tickers in the rows, and the price columns will be as follows"
# Define the columns
resultDF = pd.DataFrame(index = tickersLST, columns = ['Date','Name','Beta', 'Mkt Cap Bllns', "Last Purchase Date",
                                                       'Last Purchase Price','Recent Price',
                                                       'Price -1mo', 'Price -3mo','Price -6mo', 'Price -12mo'])

# to get the ticker name and market cap loop through the ticker list
for ticker in tickersLST:
    stock = yf.Ticker(ticker)
    # Retrieve all available information as a dictionary
    info = stock.info
    # Extract and print specific information
    company_name = info.get('longName')
    market_cap = info.get('marketCap')
    recent_price = info.get('regularMarketPrice')
    price_timestamp = info.get("regularMarketTime")
    price_timestamp_dt = datetime.fromtimestamp(price_timestamp)

    if market_cap is not None:
        market_cap = market_cap/1000000000
    beta = info.get('beta')
    resultDF.loc[ticker,'Name'] = company_name
    resultDF.loc[ticker, 'Beta'] = beta
    resultDF.loc[ticker, 'Mkt Cap Bllns'] = market_cap
    resultDF.loc[ticker, 'Recent Price'] = recent_price
    resultDF.loc[ticker, 'Date'] = price_timestamp_dt

    # Transfer last purchase date and price into resultDF
    resultDF.loc[ticker,'Last Purchase Date'] = purchaseDF.loc[ticker, 'Last Purchase Date']
    resultDF.loc[ticker,'Last Purchase Price'] = purchaseDF.loc[ticker, 'Last Purchase Price']

    # --------------------------------------------------------------------------------------------------------

    # get price for approx 1 month prior
    returndate_dt,cl_price = get_price_on_or_before(ticker, date_minus_1Str)
    # move returndate and price from historyDF to resultDF
    resultDF.loc[ticker, 'Price -1mo'] = cl_price

    # get price for approx 3 month prior
    returndate_dt,cl_price = get_price_on_or_before(ticker, date_minus_3Str)
    resultDF.loc[ticker, 'Price -3mo'] = cl_price

    # get price for approx 6 month prior
    returndate_dt,cl_price = get_price_on_or_before(ticker, date_minus_6Str)
    resultDF.loc[ticker, 'Price -6mo'] = cl_price

    # get price for approx 1 year prior
    returndate_dt, cl_price = get_price_on_or_before(ticker, date_minus_12Str )
    resultDF.loc[ticker, 'Price -12mo'] = cl_price

print(" ")
print(resultDF)
print("")

# Do calculations with resultDF
resultDF["Annualized 1mo % Price Change"] = ((resultDF["Recent Price"]/resultDF["Price -1mo"])**12-1)*100
resultDF["Annualized 3mo % Price Change"] = ((resultDF["Recent Price"]/resultDF["Price -3mo"])**4-1)*100
resultDF["Annualized 6mo % Price Change"] = ((resultDF["Recent Price"]/resultDF["Price -6mo"])**2-1)*100
resultDF["12mo % Price Change"] = ((resultDF["Recent Price"]/resultDF["Price -12mo"])-1)*100

# calculate the days since last purchased, transforming this from time delta to days
resultDF["Date"] = pd.to_datetime(resultDF["Date"], errors="coerce")
resultDF["Last Purchase Date"] = pd.to_datetime(resultDF["Last Purchase Date"], errors="coerce")
resultDF["Days Since Last Purchase"] = ( (resultDF["Date"] - resultDF["Last Purchase Date"]) .where(resultDF["Last Purchase Date"].notna()) .dt.days )



resultDF["Annualized % Price Change Since Last Purchase"] = \
    (resultDF["Recent Price"]/resultDF["Last Purchase Price"])**(365/resultDF["Days Since Last Purchase"])*\
    100-100

print("print resultDF ")
print(resultDF)
print("")
#
# ---------------------------------------------------------------------------------------------------------------
# create outputDF dataframe for output to excel. It drops some columns from resultDF.
# ***********************************************************************************************************

# outputDF is an independent copy. Changes in outputDF will not affect resultDF
outputDF= resultDF.copy(deep=True)

# drop rows that I won't need in output
outputDF = outputDF.drop(columns=["Price -1mo", "Price -3mo", "Price -6mo", "Price -12mo", "Days Since Last Purchase"])

print('output after copying and dropping values')
print(outputDF)
print("")

# Write to an excel file ********************************************************************************************

# each time the old file - if it already exists - is deleted in order to be replaced by the updated version
if os.path.exists(file_name): os.remove(file_name)

timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
timestamp_str = str(timestamp)

# Build full file path (Windows directory + timestamped filename)
file_name = fr"{file_name}"

# Write ouputDF to Excel. Note the nan_inf_to_errors': True. nan_inf_to_errors: This enables the worksheet.write()
# and write_number() methods to convert nan, inf and -inf to Excel errors. Excel doesn’t handle NAN/INF as numbers
# so as a workaround they are mapped to formulas that yield the error codes #NUM! and #DIV/0!

with pd.ExcelWriter(file_name, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
    outputDF.to_excel(writer, sheet_name='Yfin Stock Momentum', index=True, startrow=2)

    workbook=writer.book
    worksheet=writer.sheets["Yfin Stock Momentum"]

    message= "Data pulled Yahoo by yfinance with Peter's stock momentum program. No guarantees of " \
             "correctness. The prices and changes are DIVIDEND ADJUSTED.  Not for trading purposes or advice. " \
             "Last run at " + timestamp_str
    worksheet.write("A1",message)

    worksheet.write("A3", "Tickers")

    # column widths
    worksheet.set_column(5,12,15) # all
    worksheet.set_column(0,0,10) # ticker
    worksheet.set_column(1,1,12) # date
    worksheet.set_column(2,2,27)  # name
    worksheet.set_column(3,4,9) # beta, mkt cap
    worksheet.set_column(5,5,11)  # last purchase date
    worksheet.set_column(8,12,11)  # last purchase date

    # yellow background
    yellow_bg = workbook.add_format({"bg_color": "#FFFFCC"})

    # set word wrap format for Excel row 2, the titles for annual % change etc
    wrap_format = workbook.add_format({"text_wrap": True, "bold": True, "align": "center", "valign": "vcenter"})
    # Example: wrap entire third Excel row (row index 2)
    # for col in range(len(outputDF.columns)):
    i=0
    for col in outputDF.columns:
        cell_value = col
        worksheet.write(2, i+1, cell_value, wrap_format)
        i=i+1

    #Add 'Tickers' as a column header
    worksheet.write(2,0, 'Tickers', wrap_format)

    # --- Add AutoFilter to the heading row and data below---
    # first two digits are the Excel row and col where the filter starts
    worksheet.autofilter(2, 0, len(outputDF), len(outputDF.columns))

    # 1 decimal place for annualized percentage changes
    one_decimal = workbook.add_format({
        "num_format": "0.0"  # ← one decimal place
    })

    # 2 decimal places for beta, market cap
    two_decimal = workbook.add_format({
        "num_format": "0.00"  # ← two decimal places
    })

    # for last purchase date
    date_format = workbook.add_format({"num_format": "yyyy-mm-dd"})


    for dfrow in range(0,len(outputDF)):

        # 1 decimal for annualized percentage changes except since last purchase
        for dfcol in range(6, len(outputDF.columns)-1):
            value = outputDF.iloc[dfrow, dfcol]
            worksheet.write(dfrow+3, dfcol+1, value, one_decimal)

        # 1 decimal for annualized price chnge since last purchase
        value = outputDF.iloc[dfrow,11]
        if pd.isna(value):
            worksheet.write_blank(dfrow + 3, 12, None)
        else:
            worksheet.write(dfrow + 3, 12, value, one_decimal)

        # 2 decimal places for beta, market cap
        for dfcol in range(2, 4):
            value = outputDF.iloc[dfrow, dfcol]
            worksheet.write(dfrow+3, dfcol+1, value, two_decimal)

        # 2 decimal places for last purch price and Recent Price
        for dfcol in range(5, 7):
            value = outputDF.iloc[dfrow, dfcol]
            if pd.isna(value):
                worksheet.write_blank(dfrow + 3, dfcol+1, None)
            else:
                worksheet.write(dfrow + 3, dfcol+1, value, two_decimal)

        # write date in date_format
        value = outputDF.iloc[dfrow,0]
        if pd.isna(value):
            worksheet.write_blank(dfrow+3, 1, None, date_format)
        else:
            worksheet.write_datetime(dfrow+3, 1, value, date_format)

        # write last purchase date in date_format
        value = outputDF.iloc[dfrow,4]
        if pd.isna(value):
            worksheet.write_blank(dfrow+3, 5, None, date_format)
        else:
            worksheet.write_datetime(dfrow+3, 5, value, date_format)


    # --- Conditional formatting: (red) parameters 1st Excel row, 1st Excel col, last Excel row, last Excel col---
    red_format = workbook.add_format({"bg_color": "#FFC7CE"})
    worksheet.conditional_format(3, 8, 2+len(outputDF), len(outputDF.columns)-1, {"type": "cell", "criteria": "<", "value": -10, "format": red_format} )

    # --- Conditional formatting: (green) ---
    green_format = workbook.add_format({"bg_color": "#C6EFCE"})
    worksheet.conditional_format(3, 8, 2+len(outputDF), len(outputDF.columns)-1, {"type": "cell", "criteria": ">", "value": 10, "format": green_format} )

    range_to_format_conditionally = "C4:C" + str((len(outputDF)+3))
    worksheet.conditional_format( range_to_format_conditionally,
                                  { "type": "formula", "criteria": "=AND(I4>10, J4>10, K4>10, L4>10)",
                                    "format": green_format })
    worksheet.conditional_format(range_to_format_conditionally,
                                  { "type": "formula", "criteria": "=AND(I4<-10, J4<-10, K4<-10, L4<-10)",
                                  "format": red_format})

    print(f"DataFrame saved to: {file_name}")

    # --- Open the spreadsheet automatically ---
    system = platform.system()
    if system == "Windows":
        os.startfile(file_name)