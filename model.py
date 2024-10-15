import requests
from bs4 import BeautifulSoup
import pandas as pd
import os
import pyperclip  # Ensure pyperclip is installed: pip install pyperclip
from openpyxl import load_workbook  # Ensure openpyxl is installed: pip install openpyxl

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

    # Create an Excel file with two sheets
    excel_file = os.path.join(save_folder, f"${ticker}.xlsx")

    with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='financials', index=False)
        # Create an empty 'model' sheet
        pd.DataFrame().to_excel(writer, sheet_name='model', index=False)

    # Load the workbook to modify the model sheet
    wb = load_workbook(excel_file)
    model_sheet = wb['model']

    # Set specific values in certain cells
    model_sheet['A1'] = "last fcf "
    model_sheet['B1'] = "last 4q fcf"
    model_sheet['C1'] = "last 8q fcf "
    model_sheet['D1'] = "wacc"
    model_sheet['E1'] = "tgr"
    model_sheet['F1'] = "npv of fcf"
    model_sheet['G1'] = "tv"
    model_sheet['H1'] = "pv of tv"
    model_sheet['I1'] = "ev"
    model_sheet['J1'] = "net cash"
    model_sheet['K1'] = "eq val"
    model_sheet['L1'] = "shares"
    model_sheet['M1'] = "implied $"
    model_sheet['N1'] = "current $"
    model_sheet['O1'] = "upside"
    model_sheet['G2'] = "(last fcf proj)*(1+tgr)/(wacc-tgr)"
    model_sheet['H2'] = "tv/(1+wacc)^n"

    # Save the workbook
    wb.save(excel_file)
    wb.close()

    print(f"Excel file '{excel_file}' created")

# Main execution
ticker = input("Enter ticker symbol (e.g., 'aapl'): ")
save_folder = "/Users/rr/fa"  # Update this to your desired save path
scrape_financials(ticker, save_folder)
