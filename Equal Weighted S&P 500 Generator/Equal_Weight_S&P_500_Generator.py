# Generates an equal-weighted S&P 500 index fund based on the user's input portfolio value.
# Developed on 09/01/22 by Oscar David Strachan.
# Credit to Nick McCullum for the project idea.

# Improvements:
#   Modify to implement a MarketCap-Weighted S&P 500 index fund.
#   Implement dynamic version of S&P 500 constituents.
#   Potentially include full-version of IEX Cloud API.


# A mathematical computing library executing in C.
import numpy as np
# Allows tabular data manipulation.
import pandas as pd
# Executes our API calls to the IEXCloud library for stock market data.
import requests
# Allows saving well-formatted Excel documents.
import xlsxwriter
# Provides basic mathematical functions for script operations.
import math

stocks = pd.read_csv('sp_500_stocks.csv')

from secrets import IEX_CLOUD_API_TOKEN

symbol = 'AAPL'
# ?token tells the API we have permission to access its data.
api_url = f'https://sandbox.iexapis.com/stable/stock/{symbol}/quote?token={IEX_CLOUD_API_TOKEN}'
# .json() actually fetches the data and represents it as a dictionary for us to use.
data = requests.get(api_url).json()

price = data['latestPrice']
marketCap = data['marketCap']

# Setting the columns of the final dataframe as a Python list of column names.
my_columns = ['Ticker', 'Stock Price', 'Market Capitalisation', 'Number of Shares to Buy']
final_dataframe = pd.DataFrame(columns=my_columns)


# Appending the rows of the dataframe with a Series of corresponding market data.
final_dataframe = final_dataframe.append(
        pd.Series([
          symbol,
          price,
            marketCap,
            'N/A'
        ],
        index=my_columns
    ),
    ignore_index=True
)


# Splits our list of 500 stocks into sublists of 100 stocks.
def chunks(lst, n):
    """Yield successive n-sized chunks from lst."""
    for i in range(0, len(lst), n):
        yield lst[i:i + n]


# Splitting our list of 500 stocks into sublists of 100 stocks.
symbol_groups = list(chunks(stocks['Ticker'], 100))
symbol_strings = []
# Creates a list of our 100 stock sublists.
for i in range(0, len(symbol_groups)):
    symbol_strings.append(','.join(symbol_groups[i]))

final_dataframe = pd.DataFrame(columns=my_columns)

# Issuing a batch of API calls for each 100 stock symbol_string we iterate through, appending its financial data to our
# table.
for symbol_string in symbol_strings:
    batch_api_call_url = f'https://sandbox.iexapis.com/stable/stock/market/batch/?types=quote&symbols={symbol_string}&token={IEX_CLOUD_API_TOKEN}'
    data = requests.get(batch_api_call_url).json()
    # Splitting the string back into a list
    for symbol in symbol_string.split(','):
        final_dataframe = final_dataframe.append(
            pd.Series(
                [
                    # Multiple levels of parsing.
                    symbol,
                    data[symbol]['quote']['latestPrice'],
                    data[symbol]['quote']['marketCap'],
                    'N/A'
                ],
                index=my_columns
            ),
            ignore_index=True
        )



portfolio_size = input('Enter the value of your portfolio: ')

# Using try/except arguments to ensure a valid input.
try:
    val = float(portfolio_size)
except:
    print("That's not a number! \nPlease try again.")
    portfolio_size = input('Enter the value of your portfolio: ')
    val = float(portfolio_size)

# Accessing the 'Shares to Buy' column and calculating the number of shares to purchase (rounded down).
position_size = val / len(final_dataframe.index)
for i in range(0, len(final_dataframe.index)):
    final_dataframe.loc[i, 'Number of Shares to Buy'] = math.floor(
        position_size / final_dataframe.loc[i, 'Stock Price'])



# Setting up our excel document export.
writer = pd.ExcelWriter('recommended_trades.xlsx', engine='xlsxwriter')
final_dataframe.to_excel(writer, 'Recommended Trades', index=False)

background_color = '#0a0a23'
font_color = '#ffffff'

# String format for tickers.
string_format = writer.book.add_format(
    {
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    }
)

# Dollar formats for stock prices.
dollar_format = writer.book.add_format(
    {
        'num_format': '$0.00',
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    }
)

# Integer format for the number of shares to purchase.
integer_format = writer.book.add_format(
    {
        'num_format': '0',
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    }
)

column_formats = {
    'A': ['Ticker', string_format],
    'B': ['Stock Price', dollar_format],
    'C': ['Market Capitalisation', dollar_format],
    'D': ['Number of Shares to Buy', integer_format]
}

# Iterating through our dictionary by key to format our cells & columns properly in the Excel document.
for column in column_formats.keys():
    writer.sheets['Recommended Trades'].set_column(f'{column}:{column}', 18, column_formats[column][1])
    writer.sheets['Recommended Trades'].write(f'{column}1', column_formats[column][0], column_formats[column][1])

writer.save()
