import requests
import pandas as pd
import os
from openpyxl import Workbook
from openpyxl.styles import numbers

# Your API key
api_key = "cbqZ2QHLPY7B5zq0SVrBDxtLcri82HCQ"

# Base URL for FMP API
base_url = "https://financialmodelingprep.com/api/v3"

# Folder to save the Excel file
save_folder = "/Users/rr/fa"

# Function to fetch financial statements
def fetch_financial_statements(ticker):
    endpoints = {
        "income_statement": f"{base_url}/income-statement/{ticker}?apikey={api_key}",
        "cash_flow": f"{base_url}/cash-flow-statement/{ticker}?apikey={api_key}"
    }

    data_frames = {}
    
    # Loop through the endpoints and store the data in pandas DataFrames
    for statement, url in endpoints.items():
        response = requests.get(url)
        if response.status_code == 200:
            data = response.json()
            df = pd.DataFrame(data)
            data_frames[statement] = df
        else:
            print(f"Failed to fetch {statement} data for {ticker}")
            return None

    return data_frames

# Function to apply custom formatting for large numbers
def apply_custom_format(ws_model):
    # Loop through all cells in the sheet
    for row in ws_model.iter_rows(min_row=1, max_row=ws_model.max_row, min_col=1, max_col=ws_model.max_column):
        for cell in row:
            if isinstance(cell.value, (int, float)) and abs(cell.value) >= 1000000:
                cell.number_format = '#,###,,'

# Function to save the financial statements to an Excel file using openpyxl
def save_to_excel(ticker, data_frames, income_start_row=5, cashflow_start_row=None):
    # Create the full path for the Excel file
    file_name = os.path.join(save_folder, f"${ticker}.xlsx")
    
    # Create an openpyxl Workbook
    wb = Workbook()
    
    # Create the first sheet named "model"
    ws_model = wb.active
    ws_model.title = "model"
    
    # Create the second sheet named "main" (empty)
    ws_main = wb.create_sheet(title="main")
    
    # Get the income statement and cash flow statement as DataFrames
    income_statement = data_frames['income_statement'].drop(columns=['symbol'], errors='ignore').T
    cash_flow = data_frames['cash_flow'].drop(columns=['symbol'], errors='ignore').T

    # Set the default starting row for the cash flow statement if not provided
    if cashflow_start_row is None:
        cashflow_start_row = income_start_row + len(income_statement) + 5  # Add 5 empty rows after the income statement

    # Write the income statement to the specified start row
    for r_idx, row in enumerate(income_statement.itertuples(), income_start_row):
        for c_idx, value in enumerate(row, 1):
            ws_model.cell(row=r_idx, column=c_idx, value=value)

    # Add empty rows
    for _ in range(5):
        ws_model.append([])  # Add 5 empty rows

    # Write the cash flow statement to the specified start row
    for r_idx, row in enumerate(cash_flow.itertuples(), cashflow_start_row):
        for c_idx, value in enumerate(row, 1):
            ws_model.cell(row=r_idx, column=c_idx, value=value)

    # Set the width of column A to 30
    ws_model.column_dimensions['A'].width = 30

    # Set the width of all other columns to 10
    for col in ws_model.iter_cols(min_col=2, max_col=ws_model.max_column):
        ws_model.column_dimensions[col[0].column_letter].width = 10

    # Specify names to remove
    names_to_remove = [
        "reportedCurrency", 
        "cik", 
        "fillingDate", 
        "acceptedDate", 
        "link", 
        "finalLink", 
        "period", 
        "grossProfitRatio", 
        "ebitdaratio", 
        "operatingIncomeRatio", 
        "incomeBeforeTaxRatio", 
        "netIncomeRatio", 
        "sellingGeneralAndAdministrativeExpenses", 
        "otherExpenses", 
        "costAndExpenses", 
        "interestIncome", 
        "interestExpense", 
        "depreciationAndAmortization", 
        "ebitda", 
        "epsdiluted", 
        "weightedAverageShsOutDil",
        "deferredIncomeTax",
        "stockBasedCompensation",
        "stockBasedCompensation",
        "accountsReceivables",
        "inventory",
        "accountsPayables",
        "otherWorkingCapital",
        "otherNonCashItems",
        "acquisitionsNet",
        "purchasesOfInvestments",
        "salesMaturitiesOfInvestments",
        "otherInvestingActivites",
        "debtRepayment",
        "commonStockIssued",
        "commonStockRepurchased",
        "dividendsPaid",
        "otherFinancingActivites",
        "effectOfForexChangesOnCash",
        "operatingCashFlow",
        "capitalExpenditure", 
        "changeInWorkingCapital"
    ]

    # Iterate through the rows and collect rows to delete based on column A
    rows_to_delete = []
    for row in ws_model.iter_rows(min_row=1, max_row=ws_model.max_row, min_col=1, max_col=1):
        for cell in row:
            if cell.value in names_to_remove:
                rows_to_delete.append(cell.row)

    # Delete rows in reverse order to maintain correct indexing
    for row in sorted(rows_to_delete, reverse=True):
        ws_model.delete_rows(row)

    # Reverse the order of columns B:F while keeping column A intact
    max_row = ws_model.max_row
    for r_idx in range(1, max_row + 1):
        values = [ws_model.cell(row=r_idx, column=c_idx).value for c_idx in range(2, 7)]  # Get values from B to F
        for c_idx, value in enumerate(reversed(values), start=2):  # Reverse and set back
            ws_model.cell(row=r_idx, column=c_idx, value=value)

    # Add your specified values and formulas in row 1 and row 2
    ws_model['C1'] = "fgr"
    ws_model['D1'] = "wacc"
    ws_model['E1'] = "tgr"
    ws_model['F1'] = "npv of fcf"
    ws_model['G1'] = "tv"
    ws_model['H1'] = "pv of tv"
    ws_model['I1'] = "ev"
    ws_model['J1'] = "net cash"
    ws_model['K1'] = "eq val"
    ws_model['L1'] = "shares"
    ws_model['M1'] = "implied $"
    ws_model['N1'] = "current $"
    ws_model['O1'] = "upside"
    ws_model['G5'] = "1"
    ws_model['H5'] = "=G5+1"
    ws_model['G6'] = "=F6+1"

    

    # Add formulas in G2 and H2
    ws_model['G2'] = "=(last fcf proj)*(1+tgr)/(wacc-tgr)"
    ws_model['H2'] = "=tv/(1+wacc)^n"

    # Apply custom number formatting for large numbers
    apply_custom_format(ws_model)

    # Save the workbook
    wb.save(file_name)
    print(f"Financial statements for {ticker} saved to {file_name}")

# Main function to run the script
def main():
    # Ask the user to input a ticker symbol
    ticker = input("Enter the ticker symbol: ").upper()
    
    # Fetch financial statements for the given ticker
    data_frames = fetch_financial_statements(ticker)
    
    if data_frames:
        # Save the data to an Excel file
        save_to_excel(ticker, data_frames)
    else:
        print("No financial data found.")

if __name__ == "__main__":
    main()
