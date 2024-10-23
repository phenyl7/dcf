import requests
import webbrowser
import yfinance as yf
import mplfinance as mpf
import pandas as pd
import matplotlib.pyplot as plt

def open_sec_link(ticker, form_type):
    if form_type == "10k":
        url = f"https://www.sec.gov/edgar/search/?r=el#/dateRange=all&entityName={ticker}&filter_forms=10-K"
    elif form_type == "10q":
        url = f"https://www.sec.gov/edgar/search/?r=el#/dateRange=all&entityName={ticker}&filter_forms=10-Q"
    elif form_type == "cf":
        url = f"https://www.sec.gov/edgar/search/?r=el#/dateRange=all&entityName={ticker}"
    elif form_type == "sta":
        url = f'https://stockanalysis.com/stocks/{ticker}'
    elif form_type == "fa": 
        url = f'https://stockanalysis.com/stocks/{ticker}/financials/balance-sheet/?p=quarterly'
    elif form_type == 'sa':
        url = f'https://seekingalpha.com/symbol/{ticker}'
    elif form_type == 'hds':
        url = f'https://whalewisdom.com/stock/{ticker}'
    elif form_type == 'ins':
        url = f'http://openinsider.com/search?q={ticker}'
    else:
        print("Invalid form type.")
        return

    webbrowser.open(url)

def get_cik_for_ticker(ticker):
    """Fetches the CIK for a given ticker."""
    base_url = "https://www.sec.gov/files/company_tickers.json"
    headers = {
        "User-Agent": "Your Name (your_email@example.com)"
    }

    response = requests.get(base_url, headers=headers)

    if response.status_code == 200:
        data = response.json()
        for company in data.values():
            if company['ticker'].lower() == ticker.lower():
                return company['cik_str']
    else:
        print(f"Failed to retrieve CIK data. Status code: {response.status_code}")
        return None

def get_edgar_filings(cik, filing_type=None, count=10):
    """Fetches recent filings for a company with a given CIK."""
    base_url = "https://data.sec.gov/submissions/"
    headers = {
        "User-Agent": "Your Name (your_email@example.com)"
    }

    url = f"{base_url}CIK{cik:0>10}.json"
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        data = response.json()
        filings = data.get("filings", {}).get("recent", {})
        if not filings:
            print("No filings found.")
            return []

        indices = range(len(filings.get("form", [])))
        if filing_type:
            indices = [i for i in indices if filings['form'][i] == filing_type]
        indices = indices[:count]

        filing_links = []
        for i in indices:
            accession_number = filings['accessionNumber'][i].replace('-', '')
            primary_document = filings['primaryDocument'][i]
            filing_link = f"https://www.sec.gov/Archives/edgar/data/{cik}/{accession_number}/{primary_document}"
            
            print(f"{filings['form'][i]}")
            print(f"{filings['filingDate'][i]}")
            print(f"{filings['accessionNumber'][i]}")
            print(f"{filings['primaryDocument'][i]}")
            #print(f"Filing Link: {filing_link}")
            print("-" * 3)
            
            filing_links.append(filing_link)

        # Open all filing links in a web browser
        for link in filing_links:
            webbrowser.open(link)

        return filing_links

    else:
        print(f"Failed to retrieve data. Status code: {response.status_code}")
        return []
    


def plot_stock_chart(ticker):
    """Fetches and plots the stock price for the given ticker with a default period of 5 years."""
    # Use a default period of 5 years with a weekly interval
    period = "5y"
    interval = "1wk"
    
    try:
        # Fetch the stock data using yfinance
        data = yf.download(ticker, period=period, interval=interval)

        if data.empty:
            print(f"No data found for {ticker} over the past 5 years.")
            return

        # Convert data to appropriate format for mplfinance
        data.index = pd.to_datetime(data.index)
        mpf_data = data[['Open', 'High', 'Low', 'Close', 'Volume']]

        percent_change = ((data['Close'].iloc[-1] - data['Close'].iloc[0]) / data['Close'].iloc[0]) * 100
        market_colors = mpf.make_marketcolors(
            up='green',    # Custom color for up candles
            down='red',    # Custom color for down candles
            wick='inherit',  # Custom color for wicks
            edge='white',
            volume='gray'
        )

        # Define a custom style to adjust font sizes and set plot area background color
        custom_style = mpf.make_mpf_style(
            base_mpf_style='default',
            marketcolors=market_colors,
            rc={'font.size': 8, 
                'axes.labelsize': 8, 
                'xtick.labelsize': 7, 
                'ytick.labelsize': 8,
                'text.color': 'white',
                'axes.labelcolor': 'white',
                'xtick.color': 'white',
                'ytick.color': 'white', 
                'grid.color': '#2e2e2e'},
            facecolor='#000000',  # Black background for the plot area
        )

        # Plot the data with reduced text size and return the figure
        fig, axes = mpf.plot(
            mpf_data,
            type='candle',
            style=custom_style,
            volume=True,
            tight_layout=True,
            scale_padding=dict(left=0.7, top=1, right=1.5),
            figsize=(8, 6),  # Adjust the figure size as needed
            returnfig=True,  # This is important to get the figure object
        )
        fig.suptitle(f'{ticker.upper()} 5Y: {percent_change:.2f}%', fontsize=8, color='orange')
        # Set the figure background color (change '#E6E6FA' to any color you prefer)
        fig.patch.set_facecolor('#000000')  # Black background for the margins

        # Display the plot
        plt.show(block=False)

    except Exception as e:
        print(f"Failed to download: {ticker}. Error: {e}")


def search_investor_relations(ticker):
    """Searches for the investor relations page for a given stock ticker."""
    # Get the stock info
    stock_info = yf.Ticker(ticker).info
    long_name = stock_info.get('longName', ticker)  # Use longName or fallback to ticker if not found
    
    # Prepare the search URL using the long name
    search_url = f"https://www.google.com/search?q={long_name}+investor+relations"
    
    webbrowser.open(search_url)
    print(f"Searching Investor Relations page for {long_name} ({ticker.upper()})...")

def get_stock_info(ticker):
    CYAN = "\033[36m"
    RED = "\033[31m"
    RESET = "\033[0m"
    import yfinance as yf
    import math  # Import math module for ceiling function

    try:
        stock = yf.Ticker(ticker)
        info = stock.info

        # Extract desired metrics
        long_name = info.get('longName', 'N/A')
        price = info.get('currentPrice', 'N/A')
        shares = info.get('sharesOutstanding', 'N/A')
        market_cap = info.get('marketCap', 'N/A')
        cash = info.get('totalCash', 'N/A')
        debt = info.get('totalDebt', 'N/A')
        enterprise_value = info.get('enterpriseValue', 'N/A')

        # Convert relevant values to millions and round up
        def format_millions(value):
            if isinstance(value, (int, float)):
                return f"{math.ceil(value / 1_000_000):,}"  # Round up and format with commas
            return value

        # Format all the values
        price = (price)
        shares = format_millions(shares)
        market_cap = format_millions(market_cap)
        cash = format_millions(cash)
        debt = format_millions(debt)
        enterprise_value = format_millions(enterprise_value)

        # Inform the user
        print(f"{RED}{long_name}")
        print(f"{CYAN}Price:{RESET} ${price}")  # Rounded to whole number with commas
        print(f"{CYAN}Shares:{RESET} {shares} ")  # Rounded to whole number with commas
        print(f"{CYAN}MC:{RESET} {market_cap} ")  # Rounded to whole number with commas
        print(f"{CYAN}Cash:{RESET} {cash} ")  # Rounded to whole number with commas
        print(f"{CYAN}Debt:{RESET} {debt} ")  # Rounded to whole number with commas
        print(f"{CYAN}EV:{RESET} {enterprise_value}")  # Rounded to whole number with commas

    except Exception as e:
        print(f"Failed to retrieve data for {ticker}. Error: {e}")


def print_commands():
    """Prints a list of available commands."""
    commands = {
        "ticker cf": "opens link for all filings (e.g., 'AAPL CF').",
        "ticker 10-k #": "opens links for 10k filings (e.g., 'AAPL 10-K 5').",
        "ticker 10-q #": "opens links for 10q filings (e.g., 'AAPL 10-Q 5').",
        "ticker sta": "Open Stock Analysis page for the specified ticker.",
        "ticker g": "Open StockCharts for the specified ticker.",
        "ticker q": "Open Yahoo Finance for the specified ticker.",
        "ticker n": "Open Stock Analysis page for the specified ticker.",
        "ticker fa": "Open Financials page for the specified ticker.",
        "ticker sa": "open seeking alpha link for stock",
        "ticker fv": "open finviz link for stock",
        "ticker hds": "open whalewisdom link for stock",
        "ticker ins": "open openinsider link for stock",
        "ticker g 1d/5d/1mo/3m/6m/ytd/1y/5y": "Plot a candlestick chart for the specified period",
        "?": "Display this help message."
    }
    print("Available commands:")
    for command, description in commands.items():
        print(f"{command}: {description}")

def main():
    while True:
        user_input = input("> ").strip().split()
        if user_input[0].lower() == 'q':
            break
        elif user_input[0] == '?':
            print_commands()
            continue
        elif len(user_input) < 1:
            print("No input provided.")
            continue

        ticker = user_input[0].upper()

        # Check for "ticker q" command
        if len(user_input) == 2 and user_input[1].lower() == 'q':
            get_stock_info(ticker)  # Call the new function for stock info
            continue

        # Handle SEC links or filings
        if len(user_input) == 2:
            form_type = user_input[1].lower()
            if form_type == 'ir':
                search_investor_relations(ticker)
            elif user_input[1].lower() == 'g':
                # Call plot_stock_chart with default 5y period
                plot_stock_chart(ticker)  # No need for period argument anymore
            else:
                open_sec_link(ticker, form_type)
        elif len(user_input) == 3:
            filing_type = user_input[1].upper()
            count = int(user_input[2]) if user_input[2].isdigit() else 10

            cik = get_cik_for_ticker(ticker)
            if cik:
                print(f"CIK for {ticker}: {cik}")
                get_edgar_filings(cik=cik, filing_type=filing_type, count=count)
            else:
                print(f"CIK for ticker '{ticker}' not found.")
        else:
            print("Invalid command format. Please check your input.")

if __name__ == "__main__":
    main()
