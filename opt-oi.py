from ib_insync import *
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
import numpy as np

# Connect to Interactive Brokers
ib = IB()
ib.connect('127.0.0.1', 7496, clientId=1)

def get_option_chain(symbol, expiry, min_strike=None, max_strike=None):
    stock = Stock(symbol, 'SMART', 'USD')
    ib.qualifyContracts(stock)
    
    chains = ib.reqSecDefOptParams(stock.symbol, '', stock.secType, stock.conId)
    chain = next(c for c in chains if c.exchange == 'SMART')
    
    #print(chain)  # Add this line to inspect the attributes
    
    # Filter strikes based on min_strike and max_strike if provided
    strikes = [strike for strike in chain.strikes
               if (min_strike is None or strike >= min_strike) and
                  (max_strike is None or strike <= max_strike)]
    
    contracts = [Option(symbol, expiry, strike, right, 'SMART')
                 for right in ['C', 'P']
                 for strike in strikes]
    
    contracts = ib.qualifyContracts(*contracts)
    
    return contracts


def get_option_oi(contract):
    oi_data = []

    def on_tick_price(ticker, field, price, attrib):
        if field == 86:  # 86 is the field for Open Interest
            oi_data.append(price)

    ib.reqMktData(contract, '', False, False)
    ib.sleep(2)  # Wait for data to be received
    ib.cancelMktData(contract)

    return oi_data[0] if oi_data else 0  # Return 0 if no data



# Example usage
symbol = 'SMCI'
expiry = '20240920'  # Format: YYYYMMDD
min_strike = 450.0
max_strike = 500.0

# Get option chain with filtered strikes
option_chain = get_option_chain(symbol, expiry, min_strike, max_strike)

# Get OI data for all strikes
calls_oi = {}
puts_oi = {}
for contract in option_chain:
    oi = get_option_oi(contract)
    print(f"oi={oi}")
    if contract.right == 'C':
        calls_oi[contract.strike] = oi if oi is not None else 0
        print(f"calls_oi={calls_oi}")
    elif contract.right == 'P':
        puts_oi[contract.strike] = oi if oi is not None else 0
        print(f"puts_oi={puts_oi}")
        
# Ensure no None values in the OI values
strikes = sorted(set(calls_oi.keys()).union(puts_oi.keys()))
puts_oi_values = [puts_oi.get(strike, 0) for strike in strikes]
calls_oi_values = [calls_oi.get(strike, 0) for strike in strikes]

# Debugging: Print lengths of the lists
print(f"Length of strikes: {len(strikes)}")
print(f"Length of puts_oi_values: {len(puts_oi_values)}")
print(f"Length of calls_oi_values: {len(calls_oi_values)}")

fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 8), sharex=True)

# Plot OI for puts
ax1.barh(strikes, puts_oi_values, color='tab:orange')
ax1.set_ylabel('Strike Price')
ax1.set_xlabel('Open Interest (Puts)')
ax1.set_title('Open Interest for Puts')

# Plot OI for calls
ax2.barh(strikes, calls_oi_values, color='tab:blue')
ax2.set_ylabel('Strike Price')
ax2.set_xlabel('Open Interest (Calls)')
ax2.set_title('Open Interest for Calls')

plt.tight_layout()
plt.show()

# Disconnect from IB
ib.disconnect()
