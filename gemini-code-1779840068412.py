import pandas as pd

def extract_recorded_prices(file_path):
    print(f"Reading data from {file_path}...")
    try:
        # Load the consolidated sheet
        df = pd.read_excel(file_path, sheet_name='Consodilated')
    except Exception as e:
        print(f"Error reading file or sheet: {e}")
        return None

    # Check if necessary columns exist
    required_cols = ['Symbol', 'Asset Name', 'Price']
    if not all(col in df.columns for col in required_cols):
        print(f"Error: Excel sheet must contain columns: {required_cols}")
        return None

    # Drop duplicate symbols to get a unique list of assets and their prices
    # Keeping the first occurrence if duplicates exist
    unique_assets = df[required_cols].drop_duplicates(subset=['Symbol'], keep='first')
    
    # Sort by Symbol for readability
    unique_assets = unique_assets.sort_values(by='Symbol').reset_index(drop=True)
    
    return unique_assets

def fetch_live_prices(symbols_list):
    """
    Optional function to fetch real-time live prices from Yahoo Finance.
    Requires the 'yfinance' library: pip install yfinance
    """
    try:
        import yfinance as yf
    except ImportError:
        print("\n[Info] 'yfinance' library not installed. Skipping live price fetching.")
        print("To fetch real-time live prices, run: pip install yfinance")
        return None

    print("\nFetching live market prices from Yahoo Finance...")
    live_data = []
    
    # Filter out common placeholders or internal cash tokens that aren't real stock tickers
    ignore_list = ['CASH', 'SWEEP', 'SPAXX', 'VMRXX', 'VUSXX', 'FDRXX', 'FCASH', 
                   'USDC', 'T-BILL', 'TIPS', 'PNC', 'LIC', 'CKG', 'HSA_CASH', '21J36109']

    for symbol in symbols_list:
        if pd.isna(symbol) or str(symbol).upper() in ignore_list or '-' in str(symbol):
            continue
        
        try:
            ticker = yf.Ticker(str(symbol).strip())
            # Use fast_info or history to get the latest close price safely
            info = ticker.fast_info
            live_price = info.get('last_price', None) or info.get('previous_close', None)
            
            if live_price is not None:
                live_data.append({'Symbol': symbol, 'Live Price': round(live_price, 2)})
        except Exception:
            # Skip tickers that Yahoo Finance doesn't recognize (like specific mutual fund pools)
            continue
            
    return pd.DataFrame(live_data)

if __name__ == "__main__":
    # Define file path to your consolidated pivot file
    FILE_PATH = "outputconsidolatedpivot.xlsx"
    
    # 1. Extract recorded prices from your spreadsheet
    recorded_prices_df = extract_recorded_prices(FILE_PATH)
    
    if recorded_prices_df is not None:
        print("\n--- Unique Assets and Prices Recorded in Excel ---")
        print(recorded_prices_df.to_string(index=False))
        
        # Save the extracted list to a clean CSV file
        recorded_prices_df.to_csv("extracted_asset_prices.csv", index=False)
        print("\nSaved extracted prices to 'extracted_asset_prices.csv'")
        
        # 2. Try fetching live market data for these tickers
        symbols = recorded_prices_df['Symbol'].tolist()
        live_prices_df = fetch_live_prices(symbols)
        
        if live_prices_df is not None and not live_prices_df.empty:
            # Merge both datasets to compare recorded vs live prices
            comparison_df = pd.merge(recorded_prices_df, live_prices_df, on='Symbol', how='left')
            print("\n--- Price Comparison (Recorded vs. Live Market) ---")
            print(comparison_df.to_string(index=False))
            
            # Save the comparison to a CSV file
            comparison_df.to_csv("price_comparison.csv", index=False)
            print("\nSaved comparison spreadsheet to 'price_comparison.csv'")