"""
Generate Sell Transaction Analysis

This script analyzes profit/loss from sell transactions in the investment portfolio.
Uses FIFO (First In, First Out) accounting method to calculate:
- Realized profit/loss at time of sale
- Unrealized profit/loss if sold at current prices
- Opportunity cost/gain analysis

Output can be saved as either Excel or Markdown format.
"""

import argparse
import json
from collections import deque
from datetime import datetime
from pathlib import Path

import pandas as pd
import yfinance as yf


def analyze_sell_transactions(portfolio_data):
    """
    Analyze sell transactions to compute realized and unrealized profits/losses.
    Uses FIFO (First In, First Out) accounting method.
    """
    # Collect all transactions across all accounts
    all_transactions = []
    for account in portfolio_data['accounts']:
        for transaction in account['transactions']:
            # Add account info to transaction
            transaction_with_account = transaction.copy()
            transaction_with_account['account'] = account['account_name']
            transaction_with_account['account_currency'] = account['currency']
            all_transactions.append(transaction_with_account)
    
    # Sort all transactions by date
    all_transactions.sort(key=lambda x: datetime.strptime(x['date'], "%d-%m-%Y"))
    
    # Group transactions by symbol
    symbol_transactions = {}
    for transaction in all_transactions:
        symbol = transaction['symbol']
        if symbol not in symbol_transactions:
            symbol_transactions[symbol] = []
        symbol_transactions[symbol].append(transaction)
    
    # Analyze each symbol's transactions using FIFO
    sell_analysis_results = []
    
    for symbol, transactions in symbol_transactions.items():
        # FIFO queue to track buy lots: (shares, price, date, account)
        buy_queue = deque()
        
        for transaction in transactions:
            trans_type = transaction['type'].lower()
            shares = transaction['shares']
            price = transaction['price']
            date = transaction['date']
            account = transaction['account']
            fee = transaction.get('fee', 0)
            
            if trans_type == 'buy':
                # Add to FIFO queue
                buy_queue.append((shares, price, date, account))
                
            elif trans_type == 'sell':
                # Process sell using FIFO
                remaining_to_sell = shares
                total_cost_basis = 0
                weighted_avg_cost = 0
                
                # Calculate cost basis using FIFO
                temp_queue = buy_queue.copy()
                shares_processed = 0
                
                while remaining_to_sell > 0 and temp_queue:
                    lot_shares, lot_price, lot_date, lot_account = temp_queue.popleft()
                    
                    if lot_shares <= remaining_to_sell:
                        # Use entire lot
                        shares_used = lot_shares
                        remaining_to_sell -= lot_shares
                        # Remove from actual queue
                        if buy_queue:
                            buy_queue.popleft()
                    else:
                        # Use partial lot
                        shares_used = remaining_to_sell
                        remaining_to_sell = 0
                        # Update actual queue
                        if buy_queue:
                            updated_lot = (lot_shares - shares_used, lot_price, lot_date, lot_account)
                            buy_queue[0] = updated_lot
                    
                    total_cost_basis += shares_used * lot_price
                    shares_processed += shares_used
                
                if shares_processed > 0:
                    weighted_avg_cost = total_cost_basis / shares_processed
                
                # Calculate realized profit/loss
                gross_proceeds = shares * price
                net_proceeds = gross_proceeds - fee
                realized_pnl = net_proceeds - total_cost_basis
                realized_pnl_pct = (realized_pnl / total_cost_basis * 100) if total_cost_basis > 0 else 0
                
                # Store sell transaction analysis
                sell_analysis_results.append({
                    'symbol': symbol,
                    'sell_date': date,
                    'account': account,
                    'shares_sold': shares,
                    'sell_price': price,
                    'gross_proceeds': gross_proceeds,
                    'fees': fee,
                    'net_proceeds': net_proceeds,
                    'weighted_avg_cost': weighted_avg_cost,
                    'total_cost_basis': total_cost_basis,
                    'realized_pnl': realized_pnl,
                    'realized_pnl_pct': realized_pnl_pct,
                    'currency': transaction['currency']
                })
    
    return sell_analysis_results


def get_current_prices(symbols):
    """Get current prices for symbols using yfinance"""
    current_prices = {}
    
    for symbol in symbols:
        try:
            ticker = yf.Ticker(symbol)
            hist = ticker.history(period="1d")
            if not hist.empty:
                current_prices[symbol] = hist['Close'].iloc[-1]
                print(f"✅ {symbol}: ${current_prices[symbol]:.2f}")
            else:
                print(f"❌ No data for {symbol}")
                current_prices[symbol] = None
        except Exception as e:
            print(f"❌ Error fetching {symbol}: {e}")
            current_prices[symbol] = None
    
    return current_prices


def calculate_unrealized_pnl(sell_results, current_prices):
    """Calculate unrealized profits/losses if sold at current prices"""
    for result in sell_results:
        symbol = result['symbol']
        current_price = current_prices.get(symbol)
        
        if current_price is not None:
            # Calculate what the profit/loss would be if sold at current price
            shares_sold = result['shares_sold']
            cost_basis = result['total_cost_basis']
            fees = result['fees']
            
            # Current market value
            current_market_value = shares_sold * current_price
            # Net proceeds if sold today (assuming same fee structure)
            current_net_proceeds = current_market_value - fees
            # Unrealized P&L
            unrealized_pnl = current_net_proceeds - cost_basis
            unrealized_pnl_pct = (unrealized_pnl / cost_basis * 100) if cost_basis > 0 else 0
            
            # Add to result
            result['current_price'] = current_price
            result['current_market_value'] = current_market_value
            result['current_net_proceeds'] = current_net_proceeds
            result['unrealized_pnl'] = unrealized_pnl
            result['unrealized_pnl_pct'] = unrealized_pnl_pct
            
            # Calculate opportunity cost/gain
            opportunity_difference = unrealized_pnl - result['realized_pnl']
            result['opportunity_difference'] = opportunity_difference
            
        else:
            # No current price available
            result['current_price'] = None
            result['current_market_value'] = None
            result['current_net_proceeds'] = None
            result['unrealized_pnl'] = None
            result['unrealized_pnl_pct'] = None
            result['opportunity_difference'] = None


def create_analysis_dataframe(sell_results):
    """Create DataFrame with sell transaction analysis"""
    df_sell_analysis = pd.DataFrame(sell_results)
    
    # Reorder columns for better readability
    column_order = [
        'symbol', 'sell_date', 'account', 'currency',
        'shares_sold', 'sell_price', 'weighted_avg_cost',
        'gross_proceeds', 'fees', 'net_proceeds', 'total_cost_basis',
        'realized_pnl', 'realized_pnl_pct',
        'current_price', 'current_market_value', 'current_net_proceeds',
        'unrealized_pnl', 'unrealized_pnl_pct',
        'opportunity_difference'
    ]
    
    # Reorder DataFrame columns
    df_sell_analysis = df_sell_analysis[column_order]
    
    return df_sell_analysis


def generate_summary_statistics(df_sell_analysis):
    """Generate summary statistics from the analysis"""
    summary = {}
    
    # Realized P&L summary
    summary['total_realized_pnl'] = df_sell_analysis['realized_pnl'].sum()
    summary['avg_realized_pnl_pct'] = df_sell_analysis['realized_pnl_pct'].mean()
    
    # Unrealized P&L summary (excluding None values)
    df_with_current_price = df_sell_analysis.dropna(subset=['unrealized_pnl'])
    if len(df_with_current_price) > 0:
        summary['total_unrealized_pnl'] = df_with_current_price['unrealized_pnl'].sum()
        summary['avg_unrealized_pnl_pct'] = df_with_current_price['unrealized_pnl_pct'].mean()
        summary['total_opportunity_diff'] = df_with_current_price['opportunity_difference'].sum()
    else:
        summary['total_unrealized_pnl'] = None
        summary['avg_unrealized_pnl_pct'] = None
        summary['total_opportunity_diff'] = None
    
    # Best and worst performers
    summary['best_realized'] = df_sell_analysis.loc[df_sell_analysis['realized_pnl_pct'].idxmax()]
    summary['worst_realized'] = df_sell_analysis.loc[df_sell_analysis['realized_pnl_pct'].idxmin()]
    
    # Timing analysis
    if len(df_with_current_price) > 0:
        summary['sold_too_early'] = len(df_with_current_price[df_with_current_price['opportunity_difference'] > 0])
        summary['sold_at_right_time'] = len(df_with_current_price[df_with_current_price['opportunity_difference'] < 0])
    else:
        summary['sold_too_early'] = 0
        summary['sold_at_right_time'] = 0
    
    return summary


def generate_symbol_summary(df_sell_analysis):
    """Generate detailed breakdown by symbol"""
    symbol_summary = df_sell_analysis.groupby('symbol').agg({
        'shares_sold': 'sum',
        'gross_proceeds': 'sum',
        'total_cost_basis': 'sum',
        'realized_pnl': 'sum',
        'realized_pnl_pct': 'mean',
        'unrealized_pnl': 'sum',
        'opportunity_difference': 'sum'
    }).round(2)
    
    symbol_summary['realized_pnl_total_pct'] = (
        symbol_summary['realized_pnl'] / symbol_summary['total_cost_basis'] * 100
    ).round(2)
    
    return symbol_summary


def save_as_excel(df_sell_analysis, symbol_summary, output_path):
    """Save analysis results to Excel file"""
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_sell_analysis.to_excel(writer, sheet_name='Sell Transactions', index=False)
        symbol_summary.to_excel(writer, sheet_name='Symbol Summary')
    print(f"\n💾 Excel file saved to: {output_path}")


def save_as_markdown(df_sell_analysis, summary, symbol_summary, output_path):
    """Save analysis results to Markdown file"""
    with open(output_path, 'w') as f:
        # Title
        f.write("# Sell Transaction Analysis\n\n")
        f.write(f"*Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}*\n\n")
        
        # Summary Statistics
        f.write("## 💰 Summary Statistics\n\n")
        f.write(f"- **Total Realized P&L:** ${summary['total_realized_pnl']:,.2f}\n")
        f.write(f"- **Average Realized P&L %:** {summary['avg_realized_pnl_pct']:.2f}%\n")
        
        if summary['total_unrealized_pnl'] is not None:
            f.write(f"- **Total Unrealized P&L (if sold today):** ${summary['total_unrealized_pnl']:,.2f}\n")
            f.write(f"- **Average Unrealized P&L %:** {summary['avg_unrealized_pnl_pct']:.2f}%\n")
            
            if summary['total_opportunity_diff'] > 0:
                f.write(f"- **💡 Opportunity Cost:** ${summary['total_opportunity_diff']:,.2f} (would have made more if held)\n")
            else:
                f.write(f"- **💰 Opportunity Gain:** ${abs(summary['total_opportunity_diff']):,.2f} (made more by selling)\n")
        
        # Best and Worst Performers
        f.write("\n## 📈 Best and Worst Performers\n\n")
        best = summary['best_realized']
        worst = summary['worst_realized']
        f.write(f"- **🏆 Best Realized:** {best['symbol']} (+{best['realized_pnl_pct']:.2f}%)\n")
        f.write(f"- **📉 Worst Realized:** {worst['symbol']} ({worst['realized_pnl_pct']:.2f}%)\n")
        
        # Timing Analysis
        f.write("\n## ⏰ Timing Analysis\n\n")
        f.write(f"- **Sold too early** (would be better if held): {summary['sold_too_early']} transactions\n")
        f.write(f"- **Sold at right time** (better than holding): {summary['sold_at_right_time']} transactions\n")
    
    print(f"\n💾 Markdown file saved to: {output_path}")


def main():
    """Main function to run the sell transaction analysis"""
    # Parse command line arguments
    parser = argparse.ArgumentParser(
        description="Generate sell transaction analysis for investment portfolio"
    )
    parser.add_argument(
        '--format',
        type=str,
        choices=['markdown', 'excel'],
        default='markdown',
        help='Output format for the analysis (default: markdown)'
    )
    
    args = parser.parse_args()
    
    # Set up paths
    PROJECT_PATH = Path(__file__).parent.parent
    PORTFOLIO_JSON_PATH = PROJECT_PATH / "data" / "portfolio.json"
    
    # Check if portfolio file exists
    if not PORTFOLIO_JSON_PATH.exists():
        print(f"❌ Error: Portfolio file not found at {PORTFOLIO_JSON_PATH}")
        return
    
    # Read portfolio JSON
    print(f"📖 Reading portfolio from: {PORTFOLIO_JSON_PATH}")
    with open(PORTFOLIO_JSON_PATH, "r") as f:
        portfolio_json = json.load(f)
    
    # Analyze sell transactions
    print("\n📊 Analyzing sell transactions...")
    sell_results = analyze_sell_transactions(portfolio_json)
    print(f"Found {len(sell_results)} sell transactions to analyze")
    
    # Get current prices
    sell_symbols = list(set([result['symbol'] for result in sell_results]))
    print(f"\n💵 Getting current prices for {len(sell_symbols)} symbols...")
    current_prices = get_current_prices(sell_symbols)
    
    # Calculate unrealized P&L
    print("\n📈 Calculating unrealized profits/losses...")
    calculate_unrealized_pnl(sell_results, current_prices)
    
    # Create DataFrame
    df_sell_analysis = create_analysis_dataframe(sell_results)
    print(f"\n✅ Created analysis DataFrame with {len(df_sell_analysis)} transactions")
    
    # Generate summary statistics
    summary = generate_summary_statistics(df_sell_analysis)
    
    # Generate symbol summary
    symbol_summary = generate_symbol_summary(df_sell_analysis)
    
    # Create output directory if it doesn't exist
    output_dir = PROJECT_PATH / "results" / "analyses"
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Save results based on format
    update_date = datetime.strptime(portfolio_json["updated_at"], "%d-%m-%Y").strftime("%m-%d-%Y")
    
    if args.format == 'excel':
        output_file_name = f"Sell_Transaction_Analysis_{update_date}.xlsx"
        output_path = output_dir / output_file_name
        save_as_excel(df_sell_analysis, symbol_summary, output_path)
    else:  # markdown
        output_file_name = f"Sell_Transaction_Analysis_{update_date}.md"
        output_path = output_dir / output_file_name
        save_as_markdown(df_sell_analysis, summary, symbol_summary, output_path)
    
    # Print summary to console
    print("\n" + "="*80)
    print("💰 SELL TRANSACTION SUMMARY")
    print("="*80)
    print(f"Total Realized P&L: ${summary['total_realized_pnl']:,.2f}")
    print(f"Average Realized P&L %: {summary['avg_realized_pnl_pct']:.2f}%")
    
    if summary['total_unrealized_pnl'] is not None:
        print(f"Total Unrealized P&L (if sold today): ${summary['total_unrealized_pnl']:,.2f}")
        print(f"Average Unrealized P&L %: {summary['avg_unrealized_pnl_pct']:.2f}%")
    
    print("\n✅ Analysis complete!")


if __name__ == "__main__":
    main()
