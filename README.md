# Excel Finance MCP

A comprehensive Model Context Protocol (MCP) server for Excel operations focused on finance, accounting, rental property management, and expense tracking.

## Features

### Core Excel Operations
- Create, open, save Excel workbooks
- Read/write worksheet data
- Named ranges and formulas
- Data validation and conditional formatting
- File merging and protection

### Financial Calculations
- **Investment Analysis**: NPV, IRR, MIRR, Payback Period, Profitability Index
- **Loan Analysis**: Amortization schedules, payment calculations
- **Depreciation**: Straight-line, declining balance, MACRS, Section 179
- **Bond Pricing**: Price, duration, yield calculations
- **Cost of Capital**: WACC, CAPM calculations
- **Financial Ratios**: Liquidity, leverage, profitability ratios

### Rental Property Management
- **Rent Roll Generation**: Current tenant listings with lease details
- **Vacancy Analysis**: Physical and economic vacancy rates
- **Lease Management**: Expiration tracking, renewal analysis
- **NOI Calculations**: Net Operating Income analysis
- **Cash Flow Projections**: Multi-year property cash flows
- **Market Analysis**: Comparable rent analysis
- **Cap Rate Calculations**: Investment return metrics

### Expense Tracking
- **Expense Management**: Categorized expense tracking
- **Vendor Management**: Vendor database with payment terms
- **Budget Analysis**: Budget vs actual comparisons
- **Cost Savings**: Identify optimization opportunities
- **1099 Reporting**: Vendor payment reporting
- **Cash Flow Impact**: Pending payment analysis

### Cash Flow Analysis
- **Cash Flow Statements**: Operating, investing, financing activities
- **Forecasting**: Multi-scenario cash flow projections
- **Burn Rate Analysis**: Cash runway calculations
- **Working Capital**: Efficiency analysis
- **Cash Flow at Risk**: Risk metrics (CFaR)
- **Liquidity Analysis**: Cash position assessment

### Financial Reporting
- **Income Statements**: P&L with comparative periods
- **Balance Sheets**: Assets, liabilities, equity statements
- **Trial Balance**: Account balance verification
- **Financial Ratios**: Comprehensive ratio analysis
- **Custom Templates**: Pre-built Excel templates

### Tax Calculations
- **Income Tax**: Federal and state calculations
- **Self-Employment Tax**: Social Security and Medicare
- **Payroll Taxes**: Employer and employee portions
- **Quarterly Estimates**: Estimated tax payments
- **Business Deductions**: Allowable deduction analysis
- **Tax Planning**: Strategy recommendations

### Data Validation
- **Field Validation**: Required, numeric, date, currency formats
- **Financial Statement Validation**: Cross-field validations
- **Excel Data Validation**: Worksheet data verification
- **Data Cleaning**: Automated data cleaning and formatting

## Installation

```bash
npm install
pip install -r requirements.txt
```

## Usage

Start the MCP server:
```bash
npm run build
npm start
```

Or for development:
```bash
npm run dev
```

## Available Tools

### Excel Operations
- `excel_create_workbook` - Create new workbooks
- `excel_open_file` - Open existing files
- `excel_save_file` - Save workbooks
- `excel_read_worksheet` - Read worksheet data
- `excel_write_worksheet` - Write data to worksheets
- `excel_add_worksheet` - Add new worksheets
- `excel_find_replace` - Find and replace text
- `excel_conditional_formatting` - Apply formatting rules

### Financial Calculations
- `calculate_npv` - Net Present Value
- `calculate_irr` - Internal Rate of Return
- `loan_amortization` - Loan payment schedules
- `calculate_bond_price` - Bond valuation
- `depreciation_*` - Various depreciation methods
- `calculate_financial_ratios` - Financial ratio analysis

### Rental Management
- `rental_generate_rent_roll` - Rent roll reports
- `rental_calculate_vacancy_rate` - Vacancy analysis
- `rental_calculate_noi` - Net Operating Income
- `rental_project_cash_flow` - Cash flow projections
- `rental_rent_comparables_analysis` - Market analysis

### Expense Management
- `expense_add` - Add expense entries
- `expense_summary_report` - Expense summaries
- `expense_budget_vs_actual` - Budget comparisons
- `expense_forecast` - Expense forecasting
- `expense_1099_report` - 1099 vendor reporting

### Cash Flow
- `cash_flow_statement` - Cash flow statements
- `cash_flow_forecast` - Future cash flow projections
- `cash_burn_analysis` - Burn rate analysis
- `liquidity_analysis` - Liquidity position

### Reporting
- `generate_income_statement` - P&L statements
- `generate_balance_sheet` - Balance sheets
- `generate_trial_balance` - Trial balance reports
- `financial_ratios_analysis` - Ratio analysis

### Tax Calculations
- `calculate_federal_income_tax` - Federal tax calculations
- `calculate_self_employment_tax` - SE tax calculations
- `estimated_quarterly_taxes` - Quarterly payments
- `business_tax_summary` - Business tax overview

## Architecture

- **TypeScript Server**: MCP server implementation
- **Python Calculations**: Complex financial calculations
- **Excel Integration**: ExcelJS for file operations
- **Data Validation**: Comprehensive input validation
- **Modular Design**: Separate modules for different domains

## Example Usage

```typescript
// Calculate NPV
const npv = await client.callTool("calculate_npv", {
  rate: 0.1,
  cashFlows: [100, 200, 300],
  initialInvestment: 500
});

// Generate rent roll
const rentRoll = await client.callTool("rental_generate_rent_roll", {
  propertyId: "PROP001",
  asOfDate: "2024-01-01"
});

// Create income statement
const incomeStatement = await client.callTool("generate_income_statement", {
  startDate: "2024-01-01",
  endDate: "2024-12-31",
  outputToExcel: true,
  filePath: "./income_statement.xlsx"
});
```