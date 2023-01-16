# Equal Weight S&P 500 Generator
Generates an equal-weighted S&P 500 index fund based on the user's input portfolio value using Python and relevant dependancies.
&nbsp;  
Through batch API requests to IEX Cloud, the generator pulls relevant market data on a list of the top five-hundred companies in America, and organises it into a pandas dataframe. An excel spreadsheet tabulating the data is then formatted from the dataframe using XlsxWriter before being exported to the current directory.

**Dependencies used:**
1. pandas
2. Requests
3. Numpy
4. XlsxWriter
