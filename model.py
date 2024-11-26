import openpyxl
import os

# Prompt for the ticker symbol
ticker = input("Enter the stock ticker symbol: ")

# Create a new Excel workbook
wb = openpyxl.Workbook()

# Add a new sheet named "model"
ws_model = wb.create_sheet(title="model")

# Access the default sheet and rename it to "main"
ws_main = wb.active
ws_main.title = "main"

# Populate the "main" sheet as per your instructions
ws_main['B2'] = "price"
ws_main['B3'] = "shares"
ws_main['B4'] = "mc"
ws_main['B5'] = "cash"
ws_main['B6'] = "debt"
ws_main['B7'] = "ev"
ws_main['B9'] = "net cash"
ws_main['C9'] = "=C5-C6"

# Populate the "model" sheet according to your specifications
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

# Add data and formulas
ws_model['D2'] = "15%"
ws_model['E2'] = "2%"
ws_model['I2'] = "=F2+H2"
ws_model['J2'] = "=main!C9"
ws_model['K2'] = "=I2+J2"
ws_model['L2'] = "=main!C3"
ws_model['M2'] = "=K2/L2"
ws_model['N2'] = "=main!C2"
ws_model['O2'] = "=M2/N2-1"

# Add formulas in G2 and H2
ws_model['G2'] = "=(last fcf proj)*(1+tgr)/(wacc-tgr)"
ws_model['H2'] = "=tv/(1+wacc)^n"

# Define the folder path where the file should be saved
folder_path = "/Users/rr/fa"

# Save the workbook with the ticker symbol as part of the filename
filename = os.path.join(folder_path, f"{ticker}.xlsx")

try:
    wb.save(filename)
    print(f"Excel file saved to '{filename}' successfully!")
except Exception as e:
    print(f"Failed to save the file: {e}")
