# Investment Portfolio Management System

A comprehensive portfolio management system that generates Excel templates, performs FIFO-based transaction analysis, and exports data to external platforms (TradingView, Yahoo Finance). The system integrates real portfolio data from JSON files with live market data via yfinance API to create dynamic, data-driven investment analysis perfect for tracking, analyzing, and visualizing your investment portfolio performance.

## 🚀 Quick Start

1. **Install Dependencies**
   ```bash
   pip install -r requirements.txt
   ```

2. **Add Your Portfolio Data**
   - Edit `data/portfolio.json` with your accounts, transactions, and watchlist

3. **Generate Excel Template**
   - Open and run `notebooks/generate_portfolio.ipynb` (recommended)
   - Or run: `python src/generate_investment_portfolio.py`

4. **Analyze Your Portfolio**
   - Explore data with `notebooks/explore_portfolio.ipynb`
   - Export to TradingView or Yahoo Finance using the respective notebooks

## 📊 Investment-Focused Features

### 6 Specialized Investment Sheets

1. **Overview** - Asset allocation analysis with target vs current allocation and rebalancing recommendations
2. **Holdings** - Multi-account portfolio view with real-time prices, currency conversion, and cash positions
3. **Transactions** - Complete transaction history with FIFO cost basis calculations
4. **Dividends** - Dividend income tracking with yield analysis and projections
5. **Analysis** - Performance analysis with monthly returns and benchmark comparisons
6. **Watchlist** - Real-time tracking of potential investments with fundamental metrics

### 🎨 Professional Investment Styling
- Green-themed headers reflecting financial growth
- Color-coded performance indicators (green for gains, red for losses)
- Multi-currency support with proper formatting ($, €, RON)
- Auto-adjusted column widths for readability
- Conditional formatting for quick visual analysis
- Live data indicators showing real-time vs simulated data

### 📈 Investment-Ready Visualizations
Built-in charts and visualizations:
- **Asset Allocation Charts** - Target vs current allocation with sorted pie charts
- **Portfolio Distribution** - Holdings breakdown by symbol and sector
- **Performance Tracking** - Multi-account portfolio analysis with global USD/EUR views
- **Dividend Analysis** - Income projections and yield comparisons
- **Sector Diversification** - Aggregated sector allocation pie charts

### 🔗 Real-Time Data Integration
- **yfinance API**: Live stock prices, company information, and market data
- **Currency Conversion**: Real-time exchange rates for global portfolio views
- **FIFO Accounting**: Proper cost basis calculations for tax reporting
- **Multi-Account Support**: Track holdings across different brokerage accounts

### 🔄 Export & Analysis Tools
- **CSV Exporters**: Generate portfolio files for TradingView and Yahoo Finance
- **Sell Transaction Analysis**: FIFO-based P&L analysis with opportunity cost calculations
- **Multiple Output Formats**: Excel and Markdown formats for reports
- **Interactive Notebooks**: 4 Jupyter notebooks for different workflows
- **Test Mode**: Separate test data file for quick generation, using only two example watchlist items

## Installation & Usage

1. Install required dependencies:
```bash
pip install -r requirements.txt
```

2. Set up your portfolio data:
   - Update `data/portfolio.json` with your actual portfolio transactions
   - Include account information, transactions, and watchlist symbols
   - Use `data/portfolio_test.json` for testing purposes

3. Generate the investment template:
   
   **Option A: Using Jupyter Notebook (Recommended)**
   - Open `notebooks/generate_portfolio.ipynb`
   - Set `TEST_GENERATION = True` to use test data or `False` for real data
   - Run all cells to generate the Excel file
   
   **Option B: Using Command Line**
   ```bash
   cd src
   python generate_investment_portfolio.py
   # Or for test mode:
   python generate_investment_portfolio.py --test
   ```

4. Explore and analyze your data:
   - **Portfolio EDA**: Use `notebooks/explore_portfolio.ipynb` for exploratory data analysis
   - **Sell Transaction Analysis**: Generate profit/loss analysis with FIFO accounting
   - **Compare Performance**: Analyze realized vs unrealized gains

5. Export to external platforms:
   - **TradingView**: Use `notebooks/generate_tradingview_csv.ipynb` to create CSV import file
   - **Yahoo Finance**: Use `notebooks/generate_yfinance_csv.ipynb` to create CSV import file

## Portfolio Data Structure

The system expects a `portfolio.json` file with the following structure:

```json
{
  "updated_at": "22-09-2025",
  "accounts": [
    {
      "account_name": "xtb_usd",
      "cash": 4936.21,
      "currency": "$",
      "transactions": [
        {
          "date": "03-03-2025",
          "type": "Buy",
          "symbol": "POWL",
          "name": "Powell Industries",
          "shares": 5,
          "price": 161.72,
          "currency": "$",
          "fee": 0,
          "note": "Initiation"
        }
      ]
    }
  ],
  "watchlist": [
    {
      "symbol": "IBM",
      "currency": "$",
      "note": "Earnings growth potential"
    }
  ]
}
```

### Key Data Fields:
- **accounts**: Array of brokerage accounts with cash positions and transaction history
- **transactions**: Complete transaction log (Buy, Sell, Dividend) with fees and dates
- **watchlist**: Symbols to track for potential investments
- **Multi-currency**: Support for $, €, RON, and other currencies

## Investment Data Overview

### Overview Sheet (Asset Allocation Analysis)
- **Asset Class Breakdown**: ETF, Romanian Stocks, Crypto, International Stocks, US Stocks, Cash
- **Target vs Current Allocation**: Configurable target percentages with rebalancing recommendations
- **Multi-Currency Support**: Global portfolio values converted to USD
- **Real-Time Calculations**: Current values based on live market data
- **Visual Charts**: Target vs current allocation and asset distribution pie charts

### Holdings Sheet (Multi-Account Portfolio View)
- **Account Sections**: Individual sections for each brokerage account (XTB, Tradeville, etc.)
- **Global Portfolio Views**: Consolidated USD and EUR views with currency conversion
- **Real-Time Data**: Live prices, sectors, and company information via yfinance
- **Cash Positions**: Integrated cash holdings with proper currency formatting
- **Performance Tracking**: Unrealized gains/losses with FIFO-based cost calculations
- **Visual Analytics**: Portfolio distribution and sector allocation pie charts

### Transactions Sheet (Complete Transaction History)
- **FIFO Accounting**: Proper cost basis calculations for tax reporting
- **Multi-Currency**: Support for $, €, RON transactions
- **Transaction Types**: Buy, Sell, Dividend with fee tracking
- **Real Data Integration**: Loads from portfolio.json file
- **Excel Formulas**: Dynamic calculations for total amounts and net proceeds

### Dividends Sheet (Income Tracking)
- **Yield Analysis**: Current dividend yields and projections
- **Cross-Sheet Integration**: Links to Holdings sheet for share counts
- **Income Projections**: YTD received and annual estimates
- **Visual Charts**: Dividend yields and projected income visualizations

### Analysis Sheet (Performance Tracking)
- **Historical Performance**: Monthly portfolio returns with sample data
- **Benchmark Comparison**: Portfolio vs S&P 500 performance
- **Visual Charts**: Line charts for cumulative returns and bar charts for monthly performance
- **Risk Metrics**: Alpha, beta, and volatility analysis

### Watchlist Sheet (Investment Research)
- **Real-Time Data**: Live prices, fundamentals, and company information
- **Comprehensive Metrics**: P/E ratios, dividend yields, market cap, 52-week ranges
- **Multi-Currency Support**: Track international investments
- **Research Notes**: Track investment thesis and analysis

## Investment Use Cases

### For Portfolio Management
- **Multi-Account Tracking**: Consolidate holdings across different brokers
- **Currency Management**: Handle international investments with auto-conversion
- **Performance Analysis**: Real-time unrealized gains/losses with FIFO accounting
- **Cash Position Tracking**: Monitor available cash across all accounts
- **Rebalancing**: Target vs actual allocation with specific recommendations

### For Investment Analysis
- **Real-Time Research**: Live market data for investment decisions
- **Sell Transaction Analysis**: FIFO-based profit/loss calculations with opportunity cost analysis
- **Sector Diversification**: Understand concentration risk across sectors
- **Dividend Planning**: Track income from dividend-paying positions
- **Global Portfolio Views**: USD and EUR consolidated perspectives

### For Tax & Compliance
- **FIFO Cost Basis**: Proper accounting for tax reporting
- **Realized vs Unrealized**: Track gains for tax planning
- **Multi-Currency Reporting**: Handle foreign investment tax implications
- **Transaction History**: Complete audit trail with fees and dates

### For Data Analysis
- **Jupyter Notebook Integration**: Advanced profit/loss analysis with multiple notebooks
- **CSV Export**: Export to TradingView and Yahoo Finance formats
- **Real-Time Data**: yfinance integration for live market information
- **JSON Data Structure**: Easy integration with other financial tools
- **Multiple Output Formats**: Excel and Markdown formats for sell transaction analysis
- **Portfolio EDA**: Explore portfolio structure, symbols, accounts, and transaction patterns

## Command-Line Tools

### 1. Generate Excel Portfolio Template
```bash
python src/generate_investment_portfolio.py [--test]
```
**Options:**
- `--test`: Use portfolio_test.json instead of portfolio.json

**Output:** Excel file with 6 sheets (Overview, Holdings, Transactions, Dividends, Analysis, Watchlist)

### 2. Sell Transaction Analysis
```bash
python src/generate_sell_transaction_analysis.py --format=<format>
```
**Options:**
- `--format=excel`: Generate Excel workbook with detailed P&L analysis
- `--format=markdown`: Generate Markdown format report

**Output:** FIFO-based profit/loss analysis for all sell transactions

### 3. TradingView CSV Export
```bash
python src/csv_generators/tradingview.py
```
**Output:** CSV file formatted for TradingView portfolio import

### 4. Yahoo Finance CSV Export
```bash
python src/csv_generators/yfinance.py
```
**Output:** CSV file formatted for Yahoo Finance portfolio import

## Jupyter Notebooks

All command-line tools are also available through Jupyter notebooks for interactive use:

1. **generate_portfolio.ipynb**: Main workflow for generating Excel template with test mode toggle
2. **explore_portfolio.ipynb**: Portfolio EDA, sell transaction analysis, and yfinance data exploration
3. **generate_tradingview_csv.ipynb**: Interactive TradingView CSV generation
4. **generate_yfinance_csv.ipynb**: Interactive Yahoo Finance CSV generation

## Customization Options

The template can be easily customized for:
- **Additional Accounts**: Add more brokerage accounts to portfolio.json
- **New Asset Classes**: Extend asset class mapping in the code
- **Different Currencies**: Add support for additional currencies
- **Custom Target Allocations**: Modify target percentages in the Overview sheet
- **Extended Watchlist**: Add more symbols for research tracking
- **Alternative Data Sources**: Replace yfinance with other financial APIs
- **Export Formats**: Customize CSV generators for other platforms
- **Analysis Formats**: Choose between Excel or Markdown output for sell transaction analysis

## Key Technologies

This project is built with:

- **Python 3.x**: Core programming language
- **openpyxl (3.1.5)**: Excel file generation and formatting
- **pandas (2.3.2)**: Data manipulation and analysis
- **yfinance (0.2.65)**: Real-time market data from Yahoo Finance API
- **requests (2.32.5)**: HTTP requests for currency conversion and API calls
- **Jupyter Notebooks**: Interactive data exploration and workflow automation

All dependencies are managed via `requirements.txt` for easy installation.