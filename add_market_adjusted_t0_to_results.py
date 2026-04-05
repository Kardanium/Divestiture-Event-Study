import pandas as pd

# Load divestiture event study results
results_file = 'divestiture_event_study_results.xlsx'
results_df = pd.read_excel(results_file, sheet_name='Event_CARs')

# Load price data
price_data = pd.read_csv('Price Data.csv')

# Ensure dates are in datetime format
results_df['announcement_date'] = pd.to_datetime(results_df['announcement_date'])
price_data['date'] = pd.to_datetime(price_data['date'])

# Initialize a list to store market-adjusted returns
mar_t0_values = []

for index, row in results_df.iterrows():
    firm_ticker = row['ticker']
    announcement_date = row['announcement_date']
    next_day = announcement_date + pd.DateOffset(days=1)

    # Filter price data for this firm and date
    firm_data = price_data[price_data['ticker'] == firm_ticker]
    market_data = price_data[price_data['ticker'] == '^FTSE']

    if not firm_data.empty and not market_data.empty:
        # Get price on the next trading day, if available
        firm_price = firm_data[firm_data['date'] >= next_day].iloc[0]['price'] if not firm_data[firm_data['date'] >= next_day].empty else None
        market_price = market_data[market_data['date'] >= next_day].iloc[0]['price'] if not market_data[market_data['date'] >= next_day].empty else None

        # Calculate market-adjusted return
        if firm_price is not None and market_price is not None:
            mar_t0 = (firm_price - market_price) / market_price
            mar_t0_values.append(mar_t0)
        else:
            mar_t0_values.append(None) # No price available
    else:
        mar_t0_values.append(None) # No matching ticker

# Add MAR_t0 values to results DataFrame
results_df['Market_Adjusted_t0'] = mar_t0_values

# Write back to the same Excel file, preserving formatting
with pd.ExcelWriter(results_file, engine='openpyxl', mode='a') as writer:
    results_df.to_excel(writer, sheet_name='Event_CARs', index=False)

