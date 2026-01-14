""" Script to generate a CSV file with transactions from portfolio.json formatted for TradingView import. """
import json
from datetime import datetime
from pathlib import Path
import pandas as pd


PORTFOLIO_JSON_PATH = Path(__file__).parent.parent / "data" / "portfolio.json"
RESULTS_PATH = Path(__file__).parent.parent / "results" / "tradingview"


def main():
    # Load portfolio data
    with open(PORTFOLIO_JSON_PATH, "r", encoding="utf-8") as f:
        portfolio = json.load(f)
    
    # Collect all transactions
    rows = []
    for account in portfolio.get("accounts", []):
        for transaction in account.get("transactions", []):
            symbol = transaction.get("symbol")
            side = transaction.get("type")

            # Determine Qty and Fill Price based on Side
            if side in ("Buy", "Sell"):
                qty = transaction.get("shares")
                fill_price = transaction.get("price")
            else:  # Dividend
                qty = transaction.get("price")
                fill_price = None
            
            commission = transaction.get("fee")

            # Convert date from DD-MM-YYYY to YYYY-MM-DD
            date_str = transaction.get("date")
            day, month, year = date_str.split('-')
            closing_time = f"{year}-{month}-{day}"

            rows.append({
                "Symbol": symbol,
                "Side": side,
                "Qty": qty,
                "Fill Price": fill_price,
                "Commission": commission,
                "Closing Time": closing_time,
            })

    # Create DataFrame and save to CSV
    df = pd.DataFrame(rows)
    updated_at = portfolio['updated_at']
    updated_at = datetime.strptime(updated_at, "%d-%m-%Y").strftime("%Y-%m-%d")
    output_path = RESULTS_PATH / f"TradingView_Transactions_{updated_at}.csv"
    output_path.mkdir(parents=True, exist_ok=True)
    df.to_csv(output_path, index=False)

    print(f"✅ TradingView transactions CSV file generated successfully!")
    print(f"📁 Location: {output_path}")


if __name__ == "__main__":
    main()
