import pandas as pd
import numpy as np
from datetime import datetime, date, timedelta
from dateutil.relativedelta import relativedelta
from typing import Dict, List, Optional, Tuple, Union
from dataclasses import dataclass
from enum import Enum

class StatementType(Enum):
    INCOME_STATEMENT = "Income Statement"
    BALANCE_SHEET = "Balance Sheet"
    CASH_FLOW = "Cash Flow Statement"
    TRIAL_BALANCE = "Trial Balance"
    GENERAL_LEDGER = "General Ledger"

class AccountType(Enum):
    ASSET = "Asset"
    LIABILITY = "Liability"
    EQUITY = "Equity"
    REVENUE = "Revenue"
    EXPENSE = "Expense"
    COGS = "Cost of Goods Sold"

class AccountSubtype(Enum):
    # Assets
    CURRENT_ASSET = "Current Asset"
    FIXED_ASSET = "Fixed Asset"
    INTANGIBLE_ASSET = "Intangible Asset"
    
    # Liabilities
    CURRENT_LIABILITY = "Current Liability"
    LONG_TERM_LIABILITY = "Long-term Liability"
    
    # Equity
    CONTRIBUTED_CAPITAL = "Contributed Capital"
    RETAINED_EARNINGS = "Retained Earnings"
    
    # Revenue
    OPERATING_REVENUE = "Operating Revenue"
    OTHER_REVENUE = "Other Revenue"
    
    # Expenses
    OPERATING_EXPENSE = "Operating Expense"
    ADMINISTRATIVE_EXPENSE = "Administrative Expense"
    SELLING_EXPENSE = "Selling Expense"

@dataclass
class ChartOfAccount:
    account_number: str
    account_name: str
    account_type: AccountType
    account_subtype: AccountSubtype
    parent_account: Optional[str] = None
    is_active: bool = True
    description: Optional[str] = None

@dataclass
class JournalEntry:
    entry_id: str
    date: date
    description: str
    reference: Optional[str] = None
    debits: List[Tuple[str, float]] = None  # (account_number, amount)
    credits: List[Tuple[str, float]] = None
    posted: bool = False
    created_by: Optional[str] = None

@dataclass
class TrialBalanceItem:
    account_number: str
    account_name: str
    account_type: AccountType
    debit_balance: float = 0
    credit_balance: float = 0

class FinancialReportGenerator:
    """Generate comprehensive financial reports"""
    
    def __init__(self):
        self.chart_of_accounts: Dict[str, ChartOfAccount] = {}
        self.journal_entries: List[JournalEntry] = []
        self.account_balances: Dict[str, float] = {}
        
        # Initialize standard chart of accounts
        self._initialize_standard_accounts()
    
    def _initialize_standard_accounts(self):
        """Set up standard chart of accounts"""
        standard_accounts = [
            # Assets
            ChartOfAccount("1000", "Cash and Cash Equivalents", AccountType.ASSET, AccountSubtype.CURRENT_ASSET),
            ChartOfAccount("1100", "Accounts Receivable", AccountType.ASSET, AccountSubtype.CURRENT_ASSET),
            ChartOfAccount("1200", "Inventory", AccountType.ASSET, AccountSubtype.CURRENT_ASSET),
            ChartOfAccount("1300", "Prepaid Expenses", AccountType.ASSET, AccountSubtype.CURRENT_ASSET),
            ChartOfAccount("1500", "Property, Plant & Equipment", AccountType.ASSET, AccountSubtype.FIXED_ASSET),
            ChartOfAccount("1600", "Accumulated Depreciation", AccountType.ASSET, AccountSubtype.FIXED_ASSET),
            
            # Liabilities
            ChartOfAccount("2000", "Accounts Payable", AccountType.LIABILITY, AccountSubtype.CURRENT_LIABILITY),
            ChartOfAccount("2100", "Accrued Liabilities", AccountType.LIABILITY, AccountSubtype.CURRENT_LIABILITY),
            ChartOfAccount("2200", "Short-term Debt", AccountType.LIABILITY, AccountSubtype.CURRENT_LIABILITY),
            ChartOfAccount("2500", "Long-term Debt", AccountType.LIABILITY, AccountSubtype.LONG_TERM_LIABILITY),
            
            # Equity
            ChartOfAccount("3000", "Owner's Equity", AccountType.EQUITY, AccountSubtype.CONTRIBUTED_CAPITAL),
            ChartOfAccount("3500", "Retained Earnings", AccountType.EQUITY, AccountSubtype.RETAINED_EARNINGS),
            
            # Revenue
            ChartOfAccount("4000", "Sales Revenue", AccountType.REVENUE, AccountSubtype.OPERATING_REVENUE),
            ChartOfAccount("4100", "Rental Revenue", AccountType.REVENUE, AccountSubtype.OPERATING_REVENUE),
            ChartOfAccount("4900", "Other Income", AccountType.REVENUE, AccountSubtype.OTHER_REVENUE),
            
            # COGS
            ChartOfAccount("5000", "Cost of Goods Sold", AccountType.COGS, AccountSubtype.OPERATING_EXPENSE),
            
            # Expenses
            ChartOfAccount("6000", "Salaries & Wages", AccountType.EXPENSE, AccountSubtype.OPERATING_EXPENSE),
            ChartOfAccount("6100", "Rent Expense", AccountType.EXPENSE, AccountSubtype.OPERATING_EXPENSE),
            ChartOfAccount("6200", "Utilities", AccountType.EXPENSE, AccountSubtype.OPERATING_EXPENSE),
            ChartOfAccount("6300", "Insurance", AccountType.EXPENSE, AccountSubtype.OPERATING_EXPENSE),
            ChartOfAccount("6400", "Professional Fees", AccountType.EXPENSE, AccountSubtype.ADMINISTRATIVE_EXPENSE),
            ChartOfAccount("6500", "Marketing & Advertising", AccountType.EXPENSE, AccountSubtype.SELLING_EXPENSE),
            ChartOfAccount("6600", "Depreciation Expense", AccountType.EXPENSE, AccountSubtype.OPERATING_EXPENSE),
            ChartOfAccount("6700", "Interest Expense", AccountType.EXPENSE, AccountSubtype.ADMINISTRATIVE_EXPENSE),
        ]
        
        for account in standard_accounts:
            self.chart_of_accounts[account.account_number] = account
    
    def generate_income_statement(self, start_date: date, end_date: date) -> Dict:
        """Generate Income Statement (P&L)"""
        # Get relevant journal entries for the period
        period_entries = [
            entry for entry in self.journal_entries
            if entry.posted and start_date <= entry.date <= end_date
        ]
        
        # Calculate account balances for revenue and expense accounts
        account_totals = {}
        
        for entry in period_entries:
            if entry.debits:
                for account_number, amount in entry.debits:
                    account = self.chart_of_accounts.get(account_number)
                    if account and account.account_type in [AccountType.REVENUE, AccountType.EXPENSE, AccountType.COGS]:
                        if account_number not in account_totals:
                            account_totals[account_number] = 0
                        # Debits decrease revenue, increase expenses
                        if account.account_type == AccountType.REVENUE:
                            account_totals[account_number] -= amount
                        else:
                            account_totals[account_number] += amount
            
            if entry.credits:
                for account_number, amount in entry.credits:
                    account = self.chart_of_accounts.get(account_number)
                    if account and account.account_type in [AccountType.REVENUE, AccountType.EXPENSE, AccountType.COGS]:
                        if account_number not in account_totals:
                            account_totals[account_number] = 0
                        # Credits increase revenue, decrease expenses
                        if account.account_type == AccountType.REVENUE:
                            account_totals[account_number] += amount
                        else:
                            account_totals[account_number] -= amount
        
        # Organize by statement sections
        revenue_accounts = []
        cogs_accounts = []
        expense_accounts = []
        
        for account_number, balance in account_totals.items():
            account = self.chart_of_accounts.get(account_number)
            if not account:
                continue
            
            item = {
                'account_number': account_number,
                'account_name': account.account_name,
                'amount': round(balance, 2)
            }
            
            if account.account_type == AccountType.REVENUE:
                revenue_accounts.append(item)
            elif account.account_type == AccountType.COGS:
                cogs_accounts.append(item)
            elif account.account_type == AccountType.EXPENSE:
                expense_accounts.append(item)
        
        # Calculate totals
        total_revenue = sum(item['amount'] for item in revenue_accounts)
        total_cogs = sum(item['amount'] for item in cogs_accounts)
        gross_profit = total_revenue - total_cogs
        
        total_operating_expenses = sum(
            item['amount'] for item in expense_accounts 
            if self.chart_of_accounts[item['account_number']].account_subtype == AccountSubtype.OPERATING_EXPENSE
        )
        
        total_admin_expenses = sum(
            item['amount'] for item in expense_accounts 
            if self.chart_of_accounts[item['account_number']].account_subtype == AccountSubtype.ADMINISTRATIVE_EXPENSE
        )
        
        total_selling_expenses = sum(
            item['amount'] for item in expense_accounts 
            if self.chart_of_accounts[item['account_number']].account_subtype == AccountSubtype.SELLING_EXPENSE
        )
        
        operating_income = gross_profit - total_operating_expenses - total_admin_expenses - total_selling_expenses
        
        # Other income/expenses
        interest_expense = account_totals.get("6700", 0)
        other_income = account_totals.get("4900", 0)
        
        net_income = operating_income - interest_expense + other_income
        
        return {
            'statement_type': 'Income Statement',
            'period': f"{start_date} to {end_date}",
            'revenue': {
                'line_items': revenue_accounts,
                'total': round(total_revenue, 2)
            },
            'cost_of_goods_sold': {
                'line_items': cogs_accounts,
                'total': round(total_cogs, 2)
            },
            'gross_profit': round(gross_profit, 2),
            'gross_margin_pct': round(gross_profit / total_revenue * 100, 2) if total_revenue > 0 else 0,
            'operating_expenses': {
                'operating': total_operating_expenses,
                'administrative': total_admin_expenses,
                'selling': total_selling_expenses,
                'total': round(total_operating_expenses + total_admin_expenses + total_selling_expenses, 2)
            },
            'operating_income': round(operating_income, 2),
            'operating_margin_pct': round(operating_income / total_revenue * 100, 2) if total_revenue > 0 else 0,
            'other_income_expense': {
                'interest_expense': round(interest_expense, 2),
                'other_income': round(other_income, 2)
            },
            'net_income': round(net_income, 2),
            'net_margin_pct': round(net_income / total_revenue * 100, 2) if total_revenue > 0 else 0
        }
    
    def generate_balance_sheet(self, as_of_date: date) -> Dict:
        """Generate Balance Sheet"""
        # Calculate account balances as of date
        account_balances = self._calculate_account_balances(as_of_date)
        
        # Organize by balance sheet sections
        current_assets = []
        fixed_assets = []
        current_liabilities = []
        long_term_liabilities = []
        equity_accounts = []
        
        for account_number, balance in account_balances.items():
            account = self.chart_of_accounts.get(account_number)
            if not account or balance == 0:
                continue
            
            item = {
                'account_number': account_number,
                'account_name': account.account_name,
                'amount': round(balance, 2)
            }
            
            if account.account_type == AccountType.ASSET:
                if account.account_subtype == AccountSubtype.CURRENT_ASSET:
                    current_assets.append(item)
                else:
                    fixed_assets.append(item)
            elif account.account_type == AccountType.LIABILITY:
                if account.account_subtype == AccountSubtype.CURRENT_LIABILITY:
                    current_liabilities.append(item)
                else:
                    long_term_liabilities.append(item)
            elif account.account_type == AccountType.EQUITY:
                equity_accounts.append(item)
        
        # Calculate totals
        total_current_assets = sum(item['amount'] for item in current_assets)
        total_fixed_assets = sum(item['amount'] for item in fixed_assets)
        total_assets = total_current_assets + total_fixed_assets
        
        total_current_liabilities = sum(item['amount'] for item in current_liabilities)
        total_long_term_liabilities = sum(item['amount'] for item in long_term_liabilities)
        total_liabilities = total_current_liabilities + total_long_term_liabilities
        
        total_equity = sum(item['amount'] for item in equity_accounts)
        
        return {
            'statement_type': 'Balance Sheet',
            'as_of_date': as_of_date.isoformat(),
            'assets': {
                'current_assets': {
                    'line_items': current_assets,
                    'total': round(total_current_assets, 2)
                },
                'fixed_assets': {
                    'line_items': fixed_assets,
                    'total': round(total_fixed_assets, 2)
                },
                'total_assets': round(total_assets, 2)
            },
            'liabilities': {
                'current_liabilities': {
                    'line_items': current_liabilities,
                    'total': round(total_current_liabilities, 2)
                },
                'long_term_liabilities': {
                    'line_items': long_term_liabilities,
                    'total': round(total_long_term_liabilities, 2)
                },
                'total_liabilities': round(total_liabilities, 2)
            },
            'equity': {
                'line_items': equity_accounts,
                'total': round(total_equity, 2)
            },
            'total_liabilities_and_equity': round(total_liabilities + total_equity, 2),
            'balanced': abs(total_assets - (total_liabilities + total_equity)) < 0.01
        }
    
    def _calculate_account_balances(self, as_of_date: date) -> Dict[str, float]:
        """Calculate account balances as of a specific date"""
        balances = {}
        
        # Initialize all accounts with zero balance
        for account_number in self.chart_of_accounts.keys():
            balances[account_number] = 0
        
        # Process journal entries up to the date
        for entry in self.journal_entries:
            if not entry.posted or entry.date > as_of_date:
                continue
            
            if entry.debits:
                for account_number, amount in entry.debits:
                    if account_number not in balances:
                        balances[account_number] = 0
                    balances[account_number] += amount
            
            if entry.credits:
                for account_number, amount in entry.credits:
                    if account_number not in balances:
                        balances[account_number] = 0
                    balances[account_number] -= amount
        
        return balances
    
    def generate_trial_balance(self, as_of_date: date) -> pd.DataFrame:
        """Generate trial balance"""
        account_balances = self._calculate_account_balances(as_of_date)
        
        trial_balance_items = []
        total_debits = 0
        total_credits = 0
        
        for account_number, balance in account_balances.items():
            account = self.chart_of_accounts.get(account_number)
            if not account or balance == 0:
                continue
            
            # Determine normal balance based on account type
            normal_debit_accounts = [AccountType.ASSET, AccountType.EXPENSE, AccountType.COGS]
            
            if account.account_type in normal_debit_accounts:
                debit_balance = max(0, balance)
                credit_balance = max(0, -balance)
            else:
                debit_balance = max(0, -balance)
                credit_balance = max(0, balance)
            
            trial_balance_items.append({
                'Account Number': account_number,
                'Account Name': account.account_name,
                'Account Type': account.account_type.value,
                'Debit': round(debit_balance, 2) if debit_balance > 0 else None,
                'Credit': round(credit_balance, 2) if credit_balance > 0 else None
            })
            
            total_debits += debit_balance
            total_credits += credit_balance
        
        # Add totals row
        trial_balance_items.append({
            'Account Number': '',
            'Account Name': 'TOTALS',
            'Account Type': '',
            'Debit': round(total_debits, 2),
            'Credit': round(total_credits, 2)
        })
        
        df = pd.DataFrame(trial_balance_items)
        df['Balanced'] = abs(total_debits - total_credits) < 0.01
        
        return df
    
    def comparative_income_statement(self, periods: List[Tuple[date, date]]) -> pd.DataFrame:
        """Generate comparative income statement for multiple periods"""
        statements = []
        
        for start_date, end_date in periods:
            stmt = self.generate_income_statement(start_date, end_date)
            period_name = f"{start_date.strftime('%Y-%m')} to {end_date.strftime('%Y-%m')}"
            
            # Flatten the statement for comparison
            flattened = {
                'Period': period_name,
                'Total Revenue': stmt['revenue']['total'],
                'Total COGS': stmt['cost_of_goods_sold']['total'],
                'Gross Profit': stmt['gross_profit'],
                'Gross Margin %': stmt['gross_margin_pct'],
                'Operating Expenses': stmt['operating_expenses']['total'],
                'Operating Income': stmt['operating_income'],
                'Operating Margin %': stmt['operating_margin_pct'],
                'Net Income': stmt['net_income'],
                'Net Margin %': stmt['net_margin_pct']
            }
            statements.append(flattened)
        
        df = pd.DataFrame(statements)
        
        # Add variance columns if we have multiple periods
        if len(statements) > 1:
            for i in range(1, len(statements)):
                period_name = df.iloc[i]['Period']
                prev_period_name = df.iloc[i-1]['Period']
                
                for col in ['Total Revenue', 'Gross Profit', 'Operating Income', 'Net Income']:
                    current_value = df.iloc[i][col]
                    previous_value = df.iloc[i-1][col]
                    
                    if previous_value != 0:
                        variance_pct = ((current_value - previous_value) / abs(previous_value)) * 100
                        df.loc[i, f'{col} Variance %'] = round(variance_pct, 2)
        
        return df
    
    def aged_receivables_report(self, as_of_date: date) -> pd.DataFrame:
        """Generate aged accounts receivable report"""
        # This would typically integrate with customer/invoice data
        # For now, creating a template structure
        aging_buckets = ['Current', '1-30 Days', '31-60 Days', '61-90 Days', '90+ Days']
        
        # Sample data structure - would be populated from actual AR data
        ar_data = []
        
        return pd.DataFrame(ar_data, columns=[
            'Customer', 'Invoice Number', 'Invoice Date', 'Due Date', 
            'Original Amount', 'Outstanding Balance'
        ] + aging_buckets)
    
    def aged_payables_report(self, as_of_date: date) -> pd.DataFrame:
        """Generate aged accounts payable report"""
        # Similar to AR but for payables
        aging_buckets = ['Current', '1-30 Days', '31-60 Days', '61-90 Days', '90+ Days']
        
        ap_data = []
        
        return pd.DataFrame(ap_data, columns=[
            'Vendor', 'Invoice Number', 'Invoice Date', 'Due Date', 
            'Original Amount', 'Outstanding Balance'
        ] + aging_buckets)
    
    def financial_ratios_analysis(self, as_of_date: date) -> Dict:
        """Calculate comprehensive financial ratios"""
        balance_sheet = self.generate_balance_sheet(as_of_date)
        
        # Extract balance sheet items for calculations
        assets = balance_sheet['assets']
        liabilities = balance_sheet['liabilities']
        equity = balance_sheet['equity']
        
        # Get income statement for the year
        year_start = date(as_of_date.year, 1, 1)
        income_statement = self.generate_income_statement(year_start, as_of_date)
        
        # Liquidity Ratios
        current_ratio = (assets['current_assets']['total'] / 
                        liabilities['current_liabilities']['total']) if liabilities['current_liabilities']['total'] > 0 else float('inf')
        
        # Leverage Ratios
        debt_to_equity = (liabilities['total_liabilities'] / 
                         equity['total']) if equity['total'] > 0 else float('inf')
        
        debt_to_assets = (liabilities['total_liabilities'] / 
                         assets['total_assets']) if assets['total_assets'] > 0 else 0
        
        # Profitability Ratios
        roa = (income_statement['net_income'] / 
               assets['total_assets']) * 100 if assets['total_assets'] > 0 else 0
        
        roe = (income_statement['net_income'] / 
               equity['total']) * 100 if equity['total'] > 0 else 0
        
        return {
            'as_of_date': as_of_date.isoformat(),
            'liquidity_ratios': {
                'current_ratio': round(current_ratio, 2),
                'quick_ratio': 'N/A',  # Would need inventory detail
                'cash_ratio': 'N/A'    # Would need cash detail
            },
            'leverage_ratios': {
                'debt_to_equity': round(debt_to_equity, 2),
                'debt_to_assets': round(debt_to_assets, 2),
                'equity_multiplier': round(assets['total_assets'] / equity['total'], 2) if equity['total'] > 0 else float('inf')
            },
            'profitability_ratios': {
                'gross_margin_pct': income_statement['gross_margin_pct'],
                'operating_margin_pct': income_statement['operating_margin_pct'],
                'net_margin_pct': income_statement['net_margin_pct'],
                'return_on_assets_pct': round(roa, 2),
                'return_on_equity_pct': round(roe, 2)
            }
        }

class ReportTemplates:
    """Pre-built report templates for common financial statements"""
    
    @staticmethod
    def create_rental_property_analysis(property_data: Dict) -> pd.DataFrame:
        """Create rental property analysis template"""
        template_data = [
            ['RENTAL PROPERTY ANALYSIS', '', '', ''],
            ['Property:', property_data.get('name', 'N/A'), '', ''],
            ['Address:', property_data.get('address', 'N/A'), '', ''],
            ['', '', '', ''],
            ['INCOME ANALYSIS', '', '', ''],
            ['Gross Rental Income', '', '', property_data.get('gross_rental_income', 0)],
            ['Vacancy Loss', '', '', property_data.get('vacancy_loss', 0)],
            ['Effective Rental Income', '', '', property_data.get('effective_rental_income', 0)],
            ['Other Income', '', '', property_data.get('other_income', 0)],
            ['Total Income', '', '', '=C6+C7+C8'],
            ['', '', '', ''],
            ['OPERATING EXPENSES', '', '', ''],
            ['Property Management', '', '', property_data.get('property_management', 0)],
            ['Maintenance & Repairs', '', '', property_data.get('maintenance', 0)],
            ['Insurance', '', '', property_data.get('insurance', 0)],
            ['Property Taxes', '', '', property_data.get('property_taxes', 0)],
            ['Utilities', '', '', property_data.get('utilities', 0)],
            ['Total Operating Expenses', '', '', '=SUM(D13:D17)'],
            ['', '', '', ''],
            ['NET OPERATING INCOME', '', '', '=D10-D18'],
            ['', '', '', ''],
            ['FINANCIAL ANALYSIS', '', '', ''],
            ['NOI Margin %', '', '', '=D20/D10*100'],
            ['Operating Expense Ratio %', '', '', '=D18/D10*100'],
            ['Cap Rate % (if known)', '', '', property_data.get('cap_rate', 0)],
        ]
        
        return pd.DataFrame(template_data, columns=['Item', 'Description', 'Notes', 'Amount'])
    
    @staticmethod
    def create_cash_flow_template(months: int = 12) -> pd.DataFrame:
        """Create cash flow projection template"""
        headers = ['Category'] + [f'Month {i+1}' for i in range(months)] + ['Total']
        
        template_data = [
            ['CASH INFLOWS'] + [''] * (months + 1),
            ['Revenue'] + [0] * months + ['=SUM(B2:' + f'M2)'],
            ['Other Income'] + [0] * months + ['=SUM(B3:' + f'M3)'],
            ['Total Inflows'] + [''] * months + ['=SUM(B2:B3)'],
            [''] + [''] * (months + 1),
            ['CASH OUTFLOWS'] + [''] * (months + 1),
            ['Operating Expenses'] + [0] * months + ['=SUM(B6:' + f'M6)'],
            ['Capital Expenditures'] + [0] * months + ['=SUM(B7:' + f'M7)'],
            ['Debt Service'] + [0] * months + ['=SUM(B8:' + f'M8)'],
            ['Total Outflows'] + [''] * months + ['=SUM(B6:B8)'],
            [''] + [''] * (months + 1),
            ['NET CASH FLOW'] + [''] * months + ['=B4-B10'],
            ['Beginning Balance'] + [0] + [''] * months + [''],
            ['Ending Balance'] + [''] * months + ['=B12+B11']
        ]
        
        df = pd.DataFrame(template_data, columns=headers)
        return df
    
    @staticmethod
    def create_budget_template(categories: List[str]) -> pd.DataFrame:
        """Create budget vs actual template"""
        template_data = []
        
        # Headers
        template_data.append(['Category', 'Annual Budget', 'YTD Budget', 'YTD Actual', 'Variance', 'Variance %'])
        
        # Revenue section
        template_data.append(['REVENUE', '', '', '', '', ''])
        for cat in ['Sales Revenue', 'Service Revenue', 'Other Revenue']:
            template_data.append([cat, 0, '=B{}/12*MONTH(TODAY())', 0, '=D{}-C{}', '=IF(C{}=0,0,E{}/C{}*100)'])
        template_data.append(['Total Revenue', '=SUM(B3:B5)', '=SUM(C3:C5)', '=SUM(D3:D5)', '=D6-C6', '=IF(C6=0,0,E6/C6*100)'])
        
        # Expense section
        template_data.append(['', '', '', '', '', ''])
        template_data.append(['EXPENSES', '', '', '', '', ''])
        
        for category in categories:
            template_data.append([category, 0, '=B{}/12*MONTH(TODAY())', 0, '=D{}-C{}', '=IF(C{}=0,0,E{}/C{}*100)'])
        
        next_row = len(template_data) + 1
        template_data.append(['Total Expenses', f'=SUM(B9:B{next_row-1})', f'=SUM(C9:C{next_row-1})', 
                             f'=SUM(D9:D{next_row-1})', f'=D{next_row}-C{next_row}', f'=IF(C{next_row}=0,0,E{next_row}/C{next_row}*100)'])
        
        # Net Income
        template_data.append(['', '', '', '', '', ''])
        template_data.append(['NET INCOME', f'=B6-B{next_row}', f'=C6-C{next_row}', f'=D6-D{next_row}', 
                             f'=D{next_row+2}-C{next_row+2}', f'=IF(C{next_row+2}=0,0,E{next_row+2}/C{next_row+2}*100)'])
        
        return pd.DataFrame(template_data)

if __name__ == "__main__":
    # Test financial reporting
    reporter = FinancialReportGenerator()
    
    # Add sample journal entry
    entry = JournalEntry(
        "JE001",
        date.today(),
        "Initial cash investment",
        debits=[("1000", 100000)],
        credits=[("3000", 100000)],
        posted=True
    )
    reporter.journal_entries.append(entry)
    
    # Test balance sheet
    print("Balance Sheet:")
    bs = reporter.generate_balance_sheet(date.today())
    print(json.dumps(bs, indent=2, default=str))