# Usage Guide

## Setup

1. **Install Dependencies**:
```bash
npm install
pip install openpyxl pandas numpy numpy-financial python-dateutil xlsxwriter pydantic
```

2. **Build the Project**:
```bash
npm run build
```

3. **Configure Claude Desktop**:
Add this to your Claude Desktop MCP settings (`claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "excel-finance": {
      "command": "node",
      "args": ["C:\\Users\\razor\\Desktop\\projects\\excel-mcp\\dist\\index.js"],
      "env": {}
    }
  }
}
```

4. **Start Claude Desktop** - The MCP server will auto-connect.

## Available Tools

### Excel Operations
- `excel_create_workbook` - Create new Excel files
- `excel_open_file` - Open existing Excel files
- `excel_save_file` - Save workbooks
- `excel_read_worksheet` - Read data from worksheets
- `excel_write_worksheet` - Write data to worksheets

### Financial Calculations
- `calculate_npv` - Net Present Value analysis
- `calculate_irr` - Internal Rate of Return
- `loan_amortization` - Generate loan payment schedules
- `depreciation_straight_line` - Asset depreciation
- `calculate_financial_ratios` - Financial ratio analysis

### Rental Property Tools
- `rental_generate_rent_roll` - Current tenant reports
- `rental_calculate_noi` - Net Operating Income
- `rental_project_cash_flow` - Investment cash flows
- `rental_calculate_vacancy_rate` - Vacancy analysis

### Expense Management
- `expense_add` - Add expense entries
- `expense_summary_report` - Expense analysis
- `expense_budget_vs_actual` - Budget comparisons
- `expense_1099_report` - Tax reporting

### Financial Reporting
- `generate_income_statement` - P&L statements
- `generate_balance_sheet` - Balance sheets
- `financial_ratios_analysis` - Comprehensive ratios

### Tax Calculations
- `calculate_federal_income_tax` - Federal tax calculations
- `calculate_self_employment_tax` - SE tax
- `estimated_quarterly_taxes` - Quarterly payments

## Example Usage in Claude

**"Calculate the NPV of an investment with 10% discount rate and cash flows of $1000, $1500, $2000"**

**"Generate a rent roll for property PROP001 as of today"**

**"Create an income statement Excel file for Q4 2024 and save it to ./financials.xlsx"**

**"Calculate loan amortization for a $500,000 mortgage at 6.5% for 30 years"**

**"Analyze the cap rate for a rental property with $50,000 NOI and $750,000 value"**

The MCP server provides Claude with powerful Excel and financial calculation capabilities for professional accounting and finance work.