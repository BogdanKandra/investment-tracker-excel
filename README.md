# Investment Portfolio Management Excel Template

This project generates a comprehensive Excel template specifically designed for personal investment portfolio management. The template integrates real portfolio data from JSON files and fetches live market data to create dynamic, data-driven investment analysis sheets perfect for tracking, analyzing, and visualizing your investment portfolio performance.

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

## Installation & Usage

1. Install required dependencies:
```bash
pip install -r requirements.txt
```

2. Set up your portfolio data:
   - Update `data/portfolio.json` with your actual portfolio transactions
   - Include account information, transactions, and watchlist symbols

3. Generate the investment template:
```bash
cd src
python generate_investment_portfolio.py
```

4. Explore your data:
   - Use `notebooks/explore_portfolio.ipynb` for profit/loss analysis
   - Analyze sell transactions with FIFO accounting
   - Compare realized vs unrealized gains

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
- **Jupyter Notebook Integration**: Advanced profit/loss analysis
- **CSV Export**: Further analysis in Excel or other tools
- **Real-Time Data**: yfinance integration for live market information
- **JSON Data Structure**: Easy integration with other financial tools

## Customization Options

The template can be easily customized for:
- **Additional Accounts**: Add more brokerage accounts to portfolio.json
- **New Asset Classes**: Extend asset class mapping in the code
- **Different Currencies**: Add support for additional currencies
- **Custom Target Allocations**: Modify target percentages in the Overview sheet
- **Extended Watchlist**: Add more symbols for research tracking
- **Alternative Data Sources**: Replace yfinance with other financial APIs

## Investment Best Practices

1. **Data Maintenance**: Keep portfolio.json updated with latest transactions
2. **Regular Analysis**: Run the Jupyter notebook monthly for profit/loss insights
3. **Currency Monitoring**: Review exchange rate impacts on global portfolio
4. **Rebalancing**: Use Overview sheet recommendations quarterly
5. **Tax Planning**: Leverage FIFO calculations for tax-loss harvesting
6. **Performance Tracking**: Monitor realized vs unrealized gains

## File Structure
```
├── src/
│   └── generate_investment_portfolio.py     # Main template generator
├── data/
│   └── portfolio.json                      # Your portfolio data (update this!)
├── notebooks/
│   └── explore_portfolio.ipynb             # Profit/loss analysis notebook
├── results/                               # Generated Excel files and analysis
├── requirements.txt                       # Python dependencies
└── README.md                             # This documentation
```

## Investment Metrics Included

- **Real-Time Performance**: Live unrealized gains/losses with FIFO cost basis
- **Multi-Currency Analytics**: USD and EUR global portfolio views with live exchange rates
- **Income Analysis**: Dividend tracking with yield calculations and projections
- **Risk Assessment**: Sector diversification and concentration analysis
- **Valuation Metrics**: P/E ratios, market cap, 52-week ranges, dividend yields
- **Opportunity Analysis**: Realized vs unrealized profit/loss comparisons

## Advanced Features

### Real-Time Data Integration
- **yfinance API**: Live stock prices, company info, fundamentals
- **Currency Conversion**: Real-time exchange rates for global portfolios
- **Market Data**: Current prices, 52-week ranges, market cap, P/E ratios

### FIFO Accounting System
- **Tax Compliance**: Proper cost basis calculations for tax reporting
- **Sell Analysis**: Detailed profit/loss tracking with opportunity cost analysis
- **Multi-Account**: Consolidated FIFO calculations across all accounts

### Jupyter Notebook Analytics
- **Profit/Loss Analysis**: Compare realized vs unrealized gains on all sell transactions
- **Opportunity Cost**: Analyze whether you sold too early or at the right time
- **CSV Export**: Export analysis results for further processing

## Support for Investment Decisions

This comprehensive system provides:
- **Data-Driven Decisions**: Real-time market data for informed choices
- **Tax-Efficient Trading**: FIFO tracking for optimal tax planning
- **Global Portfolio Management**: Multi-currency, multi-account consolidation
- **Performance Attribution**: Understand what drives your returns
- **Risk Management**: Sector allocation and concentration monitoring
- **Research Integration**: Watchlist with live fundamentals for new investments