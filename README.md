# Alpha_Vantage_Interface
VBA code for Excel to download and analyze historic financial security prices using the Alpha Vantage API

This is a simple tool to help determine whether the price of a security is high or low compared to the previous year's closing prices.
It displays the following metrics: % current price compared to high / low of previous year, 1-year growth, current price % above average for previous year, average % growth per year.

This code does the following:
- Connects to the Alpha Vantage API by clicking an Excel button
- Downloads historic prices of a provided financial security name
- Puts the data in another Excel tab
- Finds the max, min, current, and current - 1 year prices for the previous year
- Finds the price range between min, max and determines the percentage of the current - min compared to range
- Determines the 1-year price growth of the current vs closing price one year prior
- Determines the percentage above 1-year average for the current price
- Determines the average growth per year for a security based on available historical data
- Compares the current year's growth to the average annual growth
- Displays the 1-year growth for the adjusted price

The code found in alpha_vantage_interface.vb can be copied into the VBA editor within an Excel file.
- Macros must be enabled for the Excel file
- Developer mode must be enabled in Excel

Excel worksheet setup:
- The Stock Exchange value of 'XTSE' can be added to cell B1 if accessing TSX securities. Leave blank for US exchanges.
- The Stock Ticker value (such as 'VCE' for Vanguard Canadian ETF on TSX) should be added to cell B2.
- The Report Type value of 'cf' for .cvs responses (as opposed to JSON) can be added to cell B3.
- The Function value of 'TIME_SERIES_DAILY_ADJUSTED' for the type of query should be added to cell B4.
- The unique APIKey value acquired at https://www.alphavantage.co/ should be added to cell B7.

Running the tool:
- Option 1: Run the code directly from the VBA editor. This is the simpler way albeit less elegant.
- Option 2: Create a 'Run' button in the worksheet and link it to the VBA macro function.

Please note, the interface code to Alpha Vantage was inspired by: https://thomasrainvillelapointe.blogspot.com/2017/12/equity-valuation-excel-vba-code-to-get_16.html

The parsing, analysis, and metrics display code is custom.
