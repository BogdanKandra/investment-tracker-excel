""" Script to generate a CSV file with transactions from portfolio.json formatted for Yahoo Finance import. """
import json
from datetime import datetime
from pathlib import Path
import pandas as pd


PORTFOLIO_JSON_PATH = Path(__file__).parent.parent / "data" / "portfolio.json"
RESULTS_PATH = Path(__file__).parent.parent / "results" / "yfinance"


def main():
    # Load portfolio data
    with open(PORTFOLIO_JSON_PATH, "r", encoding="utf-8") as f:
        portfolio = json.load(f)
    
    # Collect all transactions
    rows = []
    for account in portfolio.get("accounts", []):
        for transaction in account.get("transactions", []):
            transaction_type = transaction.get("type")
            
            # Skip Dividend transactions
            if transaction_type == "Dividend":
                continue
            
            symbol = transaction.get("symbol")
            
            # Convert date from DD-MM-YYYY to YYYYMMDD
            date_str = transaction.get("date")
            day, month, year = date_str.split('-')
            trade_date = f"{year}{month}{day}"

            # Quantity is negative for Sell transactions
            shares = transaction.get("shares")
            quantity = -shares if transaction_type == "Sell" else shares

            purchase_price = transaction.get("price")
            commission = transaction.get("fee")
            comment = transaction.get("note")

            rows.append({
                "Symbol": symbol,
                "Trade Date": trade_date,
                "Purchase Price": purchase_price,
                "Quantity": quantity,
                "Commission": commission,
                "Comment": comment,
            })

    # Create DataFrame and save to CSV
    df = pd.DataFrame(rows)
    updated_at = portfolio['updated_at']
    updated_at = datetime.strptime(updated_at, "%d-%m-%Y").strftime("%Y-%m-%d")
    output_path = RESULTS_PATH / f"yFinance_Transactions_{updated_at}.csv"
    output_path.mkdir(parents=True, exist_ok=True)
    df.to_csv(output_path, index=False)

    print(f"✅ Yahoo Finance transactions CSV file generated successfully!")
    print(f"📁 Location: {output_path}")


if __name__ == "__main__":
    main()
