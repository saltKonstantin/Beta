# Beta
## Description

Simple Excel file to calculate 2Y weekly and 5Y monthly beta factor against the S&P 500

## Warnings!

Make sure to get the right ticker by using the Stocks data type in Excel. Preferably use the stock market prefix (e.g. XNAS:AAPL)

## Constraints

Main constraints come from using the =STOCKHISTORY function in Excel:

    - Microsoft 365 account needed
    - International stocks are not always included
    - Data constraints limit country-specific index comparisons (e.g., MSCI World).

## Usage

1.  **Open the Excel file.**
2.  **Enter the stock ticker:** Navigate to the designated cell (e.g., a cell labeled "Ticker Input") and enter the stock ticker symbol for the company you want to analyze. For best results, use the exchange prefix (e.g., XNAS:AAPL for Apple Inc. on NASDAQ).
3.  **Allow data to refresh:** The file uses the `STOCKHISTORY` function, which may take a moment to fetch the latest stock data. Ensure you have an active internet connection.
4.  **View results:** The calculated 2-year weekly beta and 5-year monthly beta factors will be displayed in their respective cells (e.g., "2Y Weekly Beta" and "5Y Monthly Beta").