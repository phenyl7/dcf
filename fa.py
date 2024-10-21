import matplotlib.pyplot as plt

# Define the function to calculate values
def calculate_dcf(wacc, tgr, npv_fcf, last_proj_fcf, net_cash, n, shares, current_price=None):
    # Terminal Value (TV)
    tv = last_proj_fcf * (1 + tgr) / (wacc - tgr)

    # Present Value of Terminal Value (PVTV)
    pvtv = tv / (1 + wacc) ** n

    # Enterprise Value (EV)
    ev = npv_fcf + pvtv

    # Equity Value (Equity Value = EV + Net Cash)
    equity_value = ev + net_cash

    # Implied Share Price
    implied_share_price = equity_value / shares

    if current_price is not None:
        # Calculate percentage upside
        percent_upside = ((implied_share_price - current_price) / current_price) * 100
    else:
        percent_upside = None

    return tv, pvtv, ev, equity_value, implied_share_price, percent_upside

# Prompt the user for inputs
wacc = float(input("enter wacc as %: ")) / 100
tgr = float(input("enter tgr as %: ")) / 100
npv_fcf = float(input("enter npv of future fcf: "))
last_proj_fcf = float(input("enter the last projected fcf: "))
net_cash = float(input("enter net cash: "))
n = int(input("enter # of years projected (n): "))
shares = float(input("enter # of shares: "))
current_price = float(input("enter current stock price: "))

# Perform calculations for the user-provided WACC
tv, pvtv, ev, equity_value, implied_share_price, percent_upside = calculate_dcf(
    wacc, tgr, npv_fcf, last_proj_fcf, net_cash, n, shares, current_price
)

# Display the results
print(f"\nResults:")
print(f"Terminal Value (TV): ${tv:,.2f}")
print(f"Present Value of Terminal Value (PVTV): ${pvtv:,.2f}")
print(f"Enterprise Value (EV): ${ev:,.2f}")
print(f"Equity Value: ${equity_value:,.2f}")
print(f"Implied Share Price: ${implied_share_price:,.2f}")
print(f"Percentage Upside: {percent_upside:.2f}%")

# Range of WACC values from 5% to 50%
wacc_values = [w / 100 for w in range(1, 51)]

# List to store the implied share prices for different WACC values
implied_share_prices = []

# Calculate implied share price for each WACC value
for wacc in wacc_values:
    _, _, _, _, implied_share_price, _ = calculate_dcf(wacc, tgr, npv_fcf, last_proj_fcf, net_cash, n, shares)
    implied_share_prices.append(implied_share_price)

# Plotting the distribution of implied share prices
plt.figure(figsize=(10, 6))
plt.plot(wacc_values, implied_share_prices, color='blue', marker='o', linestyle='-', linewidth=2, markersize=5)
plt.title('Distribution of Implied Share Price by WACC')
plt.xlabel('WACC (%)')
plt.ylabel('Implied Share Price ($)')
plt.grid(True)
plt.show()
