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

To use this code it can be copied into the VBA editor within an Excel file.
- Macros must be enabled for the Excel file
- Developer mode must be enabled in Excel

Please note, the interface code to Alpha Vantage was inspired by: https://thomasrainvillelapointe.blogspot.com/2017/12/equity-valuation-excel-vba-code-to-get_16.html

The parsing, analysis, and metrics display code is custom
