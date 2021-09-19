# Press Shift+F10 to execute
# Press Double Shift to search everywhere for classes, files, etc.

# Import libraries
import numpy as np
import pandas as pd
import requests
import xlsxwriter
import math

# Import API token to be able to access IEX database with information on stock
# Fake/randomized data for now
from secrets import IEX_CLOUD_API_TOKEN

# Example with first stock: AAPL
def aaplExample(stocks):
    ticker = 'AAPL'
    api_url = f'https://sandbox.iexapis.com/stable/stock/{ticker}/quote/?token={IEX_CLOUD_API_TOKEN}'

    # Check that the request went through
    data = requests.get(api_url)
    print(data.status_code == 200)

    # AAPL data
    data = requests.get(api_url).json()
    price = data['latestPrice']
    market_cap = data['marketCap']

    # Create a dataframe
    my_columns = ['Ticker', 'Stock Price', 'Market Capitalization', 'Number of Shares to Buy']
    final_df = pd.DataFrame(columns=my_columns)

    # Fill dataframe with stock info
    aapl_series = pd.Series([ticker, price, market_cap, 'N/A'], index=my_columns)
    final_df = final_df.append(aapl_series, ignore_index=True)

    return final_df

# Looping through all the tickers of the S&P 500 (inefficient)
def allStocksExample():
    # Import file with S&P 500 tickers
    stocks = pd.read_csv('sp_500_stocks.csv')

    # Create a dataframe
    my_columns = ['Ticker', 'Stock Price', 'Market Capitalization', 'Number of Shares to Buy']
    final_df = pd.DataFrame(columns = my_columns)

    # Fill dataframe with stock info
    for ticker in stocks['Ticker']:
        api_url = f'https://sandbox.iexapis.com/stable/stock/{ticker}/quote/?token={IEX_CLOUD_API_TOKEN}'
        data = requests.get(api_url).json()
        each_stock = pd.Series([ticker,data['latestPrice'],data['marketCap'],'N/A'],index=my_columns)
        final_df = final_df.append(each_stock,ignore_index=True)


# Generic function to create sublists from list of size n
def chunks(lst, n):
    for i in range(0,len(lst),n):
        yield lst[i:i+n]

# Batch API calls
def batchCalls(stocks):
    # Create a dataframe
    my_columns = ['Ticker', 'Stock Price', 'Market Capitalization', 'Number of Shares to Buy']
    final_df = pd.DataFrame(columns = my_columns)

    # Create strings of 100 tickers to be batches
    ticker_groups = list(chunks(stocks['Ticker'], 100))
    ticker_strings = []
    for i in range(len(ticker_groups)):
        ticker_strings.append(','.join(ticker_groups[i]))

    # Loop through each batch (each string of 100 tickers)
    for ticker_string in ticker_strings:
        batch_api_call_url = f'https://sandbox.iexapis.com/stable/stock/market/batch?symbols={ticker_string}&types=quote&token={IEX_CLOUD_API_TOKEN}'

        # Check that the request went through
        data_check = requests.get(batch_api_call_url)

        # If so, retrieve data
        if data_check.status_code == 200:
            data = requests.get(batch_api_call_url).json()
        else:
            print('Error in retrieval!')

        # Fill dataframe with stock info
        for ticker in ticker_string.split(','): # Parse each ticker_string of 100 tickers for individual tickers
            each_stock = pd.Series([ticker,data[ticker]['quote']['latestPrice'],data[ticker]['quote']['marketCap'],'N/A'],index=my_columns)
            final_df = final_df.append(each_stock,ignore_index=True)

    return final_df

# Determine the numbers of shares to buy so that you have an equal position in each company
def calculateSharesToBuy(final_df):
    value = None
    while value is None:
        portfolio_size = input('Enter the value of your portfolio: ')
        try:
            value = float(portfolio_size)
        except ValueError:
            print('Not a valid value for portfolio! \nPlease try again. ')

    for ticker in range(len(final_df.index)):
        position_size_for_each = value / len(final_df.index)
        final_df.loc[ticker,'Number of Shares to Buy'] = position_size_for_each/final_df.loc[ticker,'Stock Price']

    return final_df

# Write to excel to see how many shares of each stock to buy to mimic S&P500 Index
def writeToExcel(final_df):
    # Intializing writer object
    writer = pd.ExcelWriter('S&P500 Indices.xlsx', engine='xlsxwriter')
    final_df.to_excel(writer,'S&P500 Indices', index = False)

    # Formatting
    background_color = '#e1ebaf'
    font_color = '#071F60'
    attributes = {'font_color': font_color,
                  'bg_color': background_color,
                  'border': 1}
    string_format = writer.book.add_format(attributes)
    attributes['num_format'] = '$0.00'
    dollar_format = writer.book.add_format(attributes)
    attributes['num_format'] = '0.00'
    integer_format = writer.book.add_format(attributes)

    writer.sheets

    column_formats = {
        'A':['Ticker',string_format],
        'B':['Stock Price', dollar_format],
        'C':['Market Capitalization', dollar_format],
        'D':['Number of Shares to Buy', integer_format]
    }

    # Column spacing
    # writer.sheets['Recommended Trades'].set_column('C:C', 'D:D', 200)

    # Add formatting
    for column in column_formats.keys():
        writer.sheets['Recommended Trades'].set_column(f'{column}:{column}', 18, column_formats[column][1])
        writer.sheets['Recommended Trades'].write(f'{column}1', column_formats[column][0], column_formats[column][1])



    writer.save()

# Main
if __name__ == '__main__':
    # Import file with S&P 500 tickers
    stocks = pd.read_csv('sp_500_stocks.csv')
    final_df = batchCalls(stocks)
    final_df = calculateSharesToBuy(final_df)
    print(final_df)
    writeToExcel(final_df)