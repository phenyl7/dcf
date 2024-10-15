import requests
from bs4 import BeautifulSoup
import pandas as pd
import os
import pyperclip  # Ensure pyperclip is installed: pip install pyperclip
import yfinance as yf
from openpyxl import load_workbook

# Function to generate the list of quarters
def generate_quarters(start_quarter, num_quarters=20):
    q, y = start_quarter.split()
    current_quarter = int(q[1])
    current_year = int(y)
    
    quarters = []
    
    for _ in range(num_quarters):
        quarter_label = f"Q{current_quarter} {current_year}"
        quarters.append(quarter_label)
        
        current_quarter -= 1
        if current_quarter == 0:
            current_quarter = 4
            current_year -= 1
    
    return quarters

def generate_future_quarters(start_quarter, num_quarters=12):
    q, y = start_quarter.split()
    current_quarter = int(q[1])
    current_year = int(y)
    
    quarters = []

    # Start generating from the next quarter
    current_quarter += 1
    if current_quarter > 4:
        current_quarter = 1
        current_year += 1

    for _ in range(num_quarters):
        quarter_label = f"Q{current_quarter} {current_year}"
        quarters.append(quarter_label)
        
        current_quarter += 1
        if current_quarter > 4:
            current_quarter = 1
            current_year += 1
    
    return quarters

def scrape_table(url):
    try:
        response = requests.get(url)
        response.raise_for_status()
    except requests.RequestException as e:
        print(f"Error fetching data from {url}: {e}")
        return []
    
    soup = BeautifulSoup(response.content, 'html.parser')
    table = soup.find('table')
    
    if not table:
        print(f"No table found on {url}.")
        return []
    
    data = []
    for tr in table.find_all('tr'):
        cells = tr.find_all('td')
        if cells:
            row = [cell.text.strip() for cell in cells]
            data.append(row)
    
    return data

def scrape_financials(ticker, save_folder):
    financials_url = f"https://stockanalysis.com/stocks/{ticker}/financials/?p=quarterly"
    balance_sheet_url = f"https://stockanalysis.com/stocks/{ticker}/financials/balance-sheet/?p=quarterly"
    cash_flow_url = f"https://stockanalysis.com/stocks/{ticker}/financials/cash-flow-statement/?p=quarterly"
    
    financials_data = scrape_table(financials_url)
    balance_sheet_data = scrape_table(balance_sheet_url)
    cash_flow_data = scrape_table(cash_flow_url)

    separator = [[""], [""], [""]]
    combined_data = financials_data + separator + balance_sheet_data + separator + cash_flow_data
    
    if not combined_data:
        print("No financial data found.")
        return
    
    df = pd.DataFrame(combined_data)
    
    quarter_input = input("Enter the quarter (e.g., 'Q3 2024'): ").strip()
    quarters = generate_quarters(quarter_input)
    quarters.reverse()

    if df.shape[1] > 1:
        df = df.iloc[:, :-1]
        df = df.iloc[:, [0] + list(range(1, df.shape[1]))[::-1]]
    
    custom_headers = ['Metric'] + quarters[:df.shape[1] - 1]
    df.columns = custom_headers
    
    quarter_row = "\t".join(quarters)
    pyperclip.copy(quarter_row)

    # Call the function to calculate average free cash flow
    calculate_stuff(df, ticker, quarter_input, save_folder)

def calculate_stuff(df, ticker, quarter_input, save_folder):
     # Search for the rows containing relevant metrics
    fcf_row = df[df['Metric'].str.contains('Free Cash Flow', case=False, na=False)]
    fcf_growth_row = df[df['Metric'].str.contains('Free Cash Flow Growth', case=False, na=False)]
    net_cash_row = df[df['Metric'].str.contains(r'Net Cash \(Debt\)', case=False, na=False)]
    shares_outstanding_row = df[df['Metric'].str.contains(r'Shares Outstanding \(Basic\)', case=False, na=False)]

    if fcf_row.empty or fcf_growth_row.empty or net_cash_row.empty or shares_outstanding_row.empty:
        print("Required financial data not found.")
        return

    # Extract the numeric values (skip the 'Metric' column)
    def clean_and_convert(x):
        if pd.isna(x) or x == '' or x == '-':
            return 0.0
        try:
            return float(x.replace('%', '').replace(',', '').strip() or 0)
        except ValueError:
            print(f"Warning: Could not convert '{x}' to float. Using 0.0 instead.")
            return 0.0

    fcf_growth = fcf_growth_row.iloc[0, 1:].apply(clean_and_convert)
    fcf_growth_1q = clean_and_convert(fcf_growth_row.iloc[0, -1])
    last_q_fcf = clean_and_convert(fcf_row.iloc[0, -1])

    # Get the last value from the Net Cash (Debt) and Shares Outstanding (Basic) rows
    last_net_cash = clean_and_convert(net_cash_row.iloc[0, -1])
    last_shares_outstanding = clean_and_convert(shares_outstanding_row.iloc[0, -1])

    # Fetch current stock price using yfinance
    stock_data = yf.Ticker(ticker)
    current_price = stock_data.history(period='1d')['Close'].iloc[-1]

    # Calculate stuff
    avg_fcf_growth_1q = fcf_growth_1q  # This is already a single value, no need for mean()
    fcf_gr = avg_fcf_growth_1q / 2
    discount = 3.75
    tgr = 0.5

    # Create DCF model data
    fcf_data = pd.DataFrame({
        'Metric': ['last q', 'lq fcf', 'lq fcf gr %', 'fgr', 'wacc', 'tgr', 'npv of fcf', 'tv', 'pv of tv', 'ev', 'net cash', 'eq val', 'shares', 'implied sh $', 'curr sh $', '% upside'],
        'Value': [quarter_input, last_q_fcf, f"{avg_fcf_growth_1q:,.2f}", f"{fcf_gr:,.2f}%", f"{discount:,.2f}%", f"{tgr:,.2f}%", "=NPV(wacc,C22:N22)", "=n22*(1+tgr)/(wacc-tgr)", "=tv/(1+wacc)^12", "=b10+b8", last_net_cash, "=b11+b12", last_shares_outstanding, "=b13/b14", " ", "=B15/B16-1"]
    })
    
    # Generate future quarters for the forecast
    future_quarters = generate_future_quarters(quarter_input, 12)
    forecast_quarters = [quarter_input] + future_quarters
    
    # Forecast FCF for the next 12 quarters using fcf_gr
    fcf_forecast = [last_q_fcf]  # Start with last_q_fcf
    for i in range(12):
        next_fcf = fcf_forecast[-1] * (1 + fcf_gr / 100)  # Increase by growth rate
        fcf_forecast.append(next_fcf)

    # Create forecast DataFrame
    forecast_data = {
        'index': list(range(13)),
        'quarter': forecast_quarters,
        'fcf': fcf_forecast
    }
    
    forecast_df = pd.DataFrame(forecast_data)

    # Combine all data into an Excel file with separate sheets
    excel_file = os.path.join(save_folder, f"${ticker}.xlsx")

    with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='financials', index=False)
        fcf_data.to_excel(writer, sheet_name='dcf_model', index=False)
        forecast_df.to_excel(writer, sheet_name='forecast', index=False)
        
        # Get the xlsxwriter workbook and worksheet objects
        workbook = writer.book
        dcf_worksheet = writer.sheets['dcf_model']
        
        # Define named ranges for fgr, wacc, tgr, and tv
        workbook.define_name('fgr', f'=dcf_model!$B$5')
        workbook.define_name('wacc', f'=dcf_model!$B$6')
        workbook.define_name('tgr', f'=dcf_model!$B$7')
        workbook.define_name('tv', f'=dcf_model!$B$9')
        
        # Define a percentage format
        percent_format = workbook.add_format({'num_format': '0.00%'})

        right_align_format = workbook.add_format({'align': 'right'})  # Right-aligned format
        
        # Apply right alignment to column B in the dcf_model sheet
        dcf_worksheet.set_column('B:B', None, right_align_format)
        

        # Format the cells to display as percentages
        dcf_worksheet.write('B5', fcf_gr / 100, percent_format)
        dcf_worksheet.write('B6', discount / 100, percent_format)
        dcf_worksheet.write('B7', tgr / 100, percent_format)
        
        # Write the current stock price to B16
        dcf_worksheet.write('B16', current_price)
        dcf_worksheet.write('A22', 'PROJ FCF')
        dcf_worksheet.write('A21', 'INDEX')
        # Set the formula for B17
        dcf_worksheet.write_formula('B17', '=B15/B16-1')
        dcf_worksheet.write_formula('B22', '=b3')
        dcf_worksheet.write_formula('C22', '=B22*(1+fgr)')
        dcf_worksheet.write_formula('D22', '=C22*(1+fgr)')
        dcf_worksheet.write_formula('E22', '=D22*(1+fgr)')
        dcf_worksheet.write_formula('F22', '=E22*(1+fgr)')
        dcf_worksheet.write_formula('G22', '=F22*(1+fgr)')
        dcf_worksheet.write_formula('H22', '=G22*(1+fgr)')
        dcf_worksheet.write_formula('I22', '=H22*(1+fgr)')
        dcf_worksheet.write_formula('J22', '=I22*(1+fgr)')
        dcf_worksheet.write_formula('K22', '=J22*(1+fgr)')
        dcf_worksheet.write_formula('L22', '=K22*(1+fgr)')
        dcf_worksheet.write_formula('M22', '=L22*(1+fgr)')
        dcf_worksheet.write_formula('N22', '=M22*(1+fgr)')
        dcf_worksheet.write_formula('B21', '0')
        dcf_worksheet.write_formula('C21', '=B21+1')
        dcf_worksheet.write_formula('D21', '=C21+1')
        dcf_worksheet.write_formula('E21', '=D21+1')
        dcf_worksheet.write_formula('F21', '=E21+1')
        dcf_worksheet.write_formula('G21', '=F21+1')
        dcf_worksheet.write_formula('H21', '=G21+1')
        dcf_worksheet.write_formula('I21', '=H21+1')
        dcf_worksheet.write_formula('J21', '=I21+1')
        dcf_worksheet.write_formula('K21', '=J21+1')
        dcf_worksheet.write_formula('L21', '=K21+1')
        dcf_worksheet.write_formula('M21', '=L21+1')
        dcf_worksheet.write_formula('N21', '=M21+1')
        # Apply percentage format to cell B17
        dcf_worksheet.set_row(16, None, percent_format)  # Format the specific row for cell B17


    print(f"Excel file '{excel_file}' created")



# Main execution
ticker = input("Enter ticker symbol (e.g., 'aapl'): ")
save_folder = "/Users/rr/fa"
scrape_financials(ticker.lower(), save_folder)