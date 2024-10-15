import os
import pandas as pd
from openpyxl import load_workbook

# Function to create the financials and model sheets with specific cell values
def create_financials_excel(ticker, save_folder):
    # Create an empty DataFrame for the 'financials' sheet
    df = pd.DataFrame()

    # Set the file path for saving
    excel_file = os.path.join(save_folder, f"${ticker}.xlsx")

    # Create an Excel file with two sheets
    with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
        # Write the empty DataFrame to the 'financials' sheet (will be updated later)
        df.to_excel(writer, sheet_name='financials', index=False)
        # Create an empty 'model' sheet
        pd.DataFrame().to_excel(writer, sheet_name='model', index=False)

    # Load the workbook to modify both sheets
    wb = load_workbook(excel_file)

    # Update the 'financials' sheet with specific values
    financials_sheet = wb['financials']
    financials_sheet['A1'] = "ticker"
    financials_sheet['A2'] = "price"
    financials_sheet['A3'] = "shares"
    financials_sheet['A4'] = "mc"       # Market Capitalization
    financials_sheet['A5'] = "cash"
    financials_sheet['A6'] = "debt"
    financials_sheet['A7'] = "ev"       # Enterprise Value

    # Update the 'model' sheet with specific values
    model_sheet = wb['model']
    model_sheet['C1'] = "fgr"
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

    print(f"Excel file '{excel_file}' created ")

# Main execution
ticker = input("Enter ticker symbol (e.g., 'aapl'): ")
save_folder = "/Users/rr/fa"  # Update this to your desired save path
create_financials_excel(ticker, save_folder)
