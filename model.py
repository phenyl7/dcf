import openpyxl
from openpyxl import Workbook
import os

# Prompt user for the stock ticker
ticker = input("Enter the stock ticker: ")

# Create a new workbook and get the active worksheet
wb = Workbook()

# Rename the active sheet to "main"
ws_main = wb.active
ws_main.title = "main"

# Create a new sheet named "model"
ws_model = wb.create_sheet(title="model")

# Fill the "main" sheet with the provided values
ws_main['B2'] = "price"
ws_main['B3'] = "shares"
ws_main['B4'] = "mc"
ws_main['B5'] = "cash"
ws_main['B6'] = "debt"
ws_main['B7'] = "ev"

# Fill the "model" sheet with the provided values
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
ws_model['G2'] = "=(last fcf proj)*(1+tgr)/(wacc-tgr)"
ws_model['H2'] = "=tv/(1+wacc)^n"

# Define the path to save the file
save_path = "/Users/rr/fa"
filename = os.path.join(save_path, f"${ticker}.xlsx")

# Save the workbook to the specified directory
wb.save(filename)

print(f"Workbook saved as {filename}")
