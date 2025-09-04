import pandas as pd
import numpy as np
from datetime import datetime, date, timedelta
from typing import Dict, List, Optional, Tuple, Union
from dataclasses import dataclass, field
from enum import Enum
import json

class ExpenseCategory(Enum):
    # Operating Expenses
    RENT = "Rent/Lease"
    UTILITIES = "Utilities"
    SALARIES = "Salaries & Wages"
    BENEFITS = "Employee Benefits"
    INSURANCE = "Insurance"
    MARKETING = "Marketing & Advertising"
    OFFICE_SUPPLIES = "Office Supplies"
    MAINTENANCE = "Maintenance & Repairs"
    PROFESSIONAL_FEES = "Professional Fees"
    TRAVEL = "Travel & Entertainment"
    
    # Cost of Goods Sold
    MATERIALS = "Raw Materials"
    INVENTORY = "Inventory Purchases"
    FREIGHT = "Freight & Shipping"
    
    # Capital Expenses
    EQUIPMENT = "Equipment"
    PROPERTY = "Property"
    VEHICLES = "Vehicles"
    SOFTWARE = "Software"
    
    # Financial
    INTEREST = "Interest Expense"
    BANK_FEES = "Bank Fees"
    TAXES = "Taxes"
    
    # Other
    DEPRECIATION = "Depreciation"
    AMORTIZATION = "Amortization"
    OTHER = "Other"

class PaymentMethod(Enum):
    CASH = "Cash"
    CHECK = "Check"
    CREDIT_CARD = "Credit Card"
    ACH = "ACH Transfer"
    WIRE = "Wire Transfer"
    PAYPAL = "PayPal"
    OTHER = "Other"

class ApprovalStatus(Enum):
    PENDING = "Pending"
    APPROVED = "Approved"
    REJECTED = "Rejected"
    UNDER_REVIEW = "Under Review"
    PAID = "Paid"

@dataclass
class Vendor:
    vendor_id: str
    name: str
    contact_info: Dict[str, str]
    tax_id: Optional[str] = None
    payment_terms: str = "Net 30"
    preferred_payment_method: PaymentMethod = PaymentMethod.CHECK
    w9_on_file: bool = False
    active: bool = True

@dataclass
class Expense:
    expense_id: str
    date: date
    vendor_id: str
    amount: float
    category: ExpenseCategory
    subcategory: Optional[str] = None
    description: str = ""
    invoice_number: Optional[str] = None
    payment_method: PaymentMethod = PaymentMethod.CHECK
    approval_status: ApprovalStatus = ApprovalStatus.PENDING
    approved_by: Optional[str] = None
    paid_date: Optional[date] = None
    receipt_url: Optional[str] = None
    cost_center: Optional[str] = None
    project_id: Optional[str] = None
    tags: List[str] = field(default_factory=list)
    tax_deductible: bool = True
    recurring: bool = False
    recurring_frequency: Optional[str] = None

@dataclass
class Budget:
    budget_id: str
    name: str
    fiscal_year: int
    period: str  # "annual", "quarterly", "monthly"
    categories: Dict[ExpenseCategory, float]
    cost_centers: Optional[Dict[str, float]] = None
    total_budget: float = 0.0
    
    def __post_init__(self):
        if self.total_budget == 0:
            self.total_budget = sum(self.categories.values())

class ExpenseTracker:
    """Comprehensive expense tracking and management system"""
    
    def __init__(self):
        self.vendors: Dict[str, Vendor] = {}
        self.expenses: List[Expense] = []
        self.budgets: Dict[str, Budget] = {}
        self.approval_rules: Dict[str, Dict] = {}
        self.tax_categories: Dict[ExpenseCategory, float] = {}  # Category to tax rate mapping
    
    def add_expense(self, expense: Expense) -> str:
        """Add a new expense with validation"""
        # Validate vendor exists
        if expense.vendor_id not in self.vendors:
            raise ValueError(f"Vendor {expense.vendor_id} not found")
        
        # Check approval rules
        if self._requires_approval(expense):
            expense.approval_status = ApprovalStatus.PENDING
        else:
            expense.approval_status = ApprovalStatus.APPROVED
        
        self.expenses.append(expense)
        return expense.expense_id
    
    def _requires_approval(self, expense: Expense) -> bool:
        """Check if expense requires approval based on rules"""
        # Example rules - would be configurable
        if expense.amount > 5000:
            return True
        if expense.category == ExpenseCategory.EQUIPMENT and expense.amount > 1000:
            return True
        return False
    
    def get_expense_summary(self, start_date: date, end_date: date, 
                           group_by: str = "category") -> pd.DataFrame:
        """Generate expense summary report"""
        filtered_expenses = [
            e for e in self.expenses 
            if start_date <= e.date <= end_date
        ]
        
        if not filtered_expenses:
            return pd.DataFrame()
        
        df = pd.DataFrame([{
            'Date': e.date,
            'Vendor': self.vendors.get(e.vendor_id, Vendor("", "Unknown", {})).name,
            'Amount': e.amount,
            'Category': e.category.value,
            'Subcategory': e.subcategory or '',
            'Description': e.description,
            'Status': e.approval_status.value,
            'Cost Center': e.cost_center or 'Unassigned',
            'Project': e.project_id or 'None',
            'Tax Deductible': e.tax_deductible
        } for e in filtered_expenses])
        
        if group_by == "category":
            summary = df.groupby('Category').agg({
                'Amount': ['sum', 'mean', 'count'],
                'Tax Deductible': lambda x: (x == True).sum()
            }).round(2)
            summary.columns = ['Total', 'Average', 'Count', 'Tax Deductible Count']
            return summary
        elif group_by == "vendor":
            return df.groupby('Vendor')['Amount'].agg(['sum', 'mean', 'count']).round(2)
        elif group_by == "cost_center":
            return df.groupby('Cost Center')['Amount'].agg(['sum', 'mean', 'count']).round(2)
        elif group_by == "month":
            df['Month'] = pd.to_datetime(df['Date']).dt.to_period('M')
            return df.groupby('Month')['Amount'].agg(['sum', 'mean', 'count']).round(2)
        
        return df
    
    def budget_vs_actual(self, budget_id: str, start_date: date, end_date: date) -> Dict:
        """Compare actual expenses against budget"""
        budget = self.budgets.get(budget_id)
        if not budget:
            return {"error": "Budget not found"}
        
        # Get actual expenses
        filtered_expenses = [
            e for e in self.expenses 
            if start_date <= e.date <= end_date
        ]
        
        # Group by category
        actual_by_category = {}
        for expense in filtered_expenses:
            if expense.category not in actual_by_category:
                actual_by_category[expense.category] = 0
            actual_by_category[expense.category] += expense.amount
        
        # Compare with budget
        comparison = []
        total_budgeted = 0
        total_actual = 0
        
        for category, budgeted_amount in budget.categories.items():
            actual_amount = actual_by_category.get(category, 0)
            variance = budgeted_amount - actual_amount
            variance_pct = (variance / budgeted_amount * 100) if budgeted_amount > 0 else 0
            
            comparison.append({
                'Category': category.value,
                'Budgeted': budgeted_amount,
                'Actual': actual_amount,
                'Variance': variance,
                'Variance %': round(variance_pct, 2),
                'Status': 'Under' if variance > 0 else 'Over'
            })
            
            total_budgeted += budgeted_amount
            total_actual += actual_amount
        
        # Add categories with actual but no budget
        for category, actual_amount in actual_by_category.items():
            if category not in budget.categories:
                comparison.append({
                    'Category': category.value,
                    'Budgeted': 0,
                    'Actual': actual_amount,
                    'Variance': -actual_amount,
                    'Variance %': -100,
                    'Status': 'Unbudgeted'
                })
                total_actual += actual_amount
        
        return {
            'budget_name': budget.name,
            'period': f"{start_date} to {end_date}",
            'comparison': comparison,
            'total_budgeted': round(total_budgeted, 2),
            'total_actual': round(total_actual, 2),
            'total_variance': round(total_budgeted - total_actual, 2),
            'total_variance_pct': round((total_budgeted - total_actual) / total_budgeted * 100, 2) if total_budgeted > 0 else 0
        }
    
    def expense_forecast(self, months_ahead: int = 12) -> pd.DataFrame:
        """Forecast future expenses based on historical data"""
        # Get last 12 months of data
        end_date = date.today()
        start_date = end_date - timedelta(days=365)
        
        historical_expenses = [
            e for e in self.expenses 
            if start_date <= e.date <= end_date
        ]
        
        if not historical_expenses:
            return pd.DataFrame()
        
        # Create monthly aggregations
        df = pd.DataFrame([{
            'Date': e.date,
            'Amount': e.amount,
            'Category': e.category.value,
            'Recurring': e.recurring
        } for e in historical_expenses])
        
        df['Month'] = pd.to_datetime(df['Date']).dt.to_period('M')
        monthly_totals = df.groupby('Month')['Amount'].sum()
        
        # Simple moving average forecast
        forecast_data = []
        last_3_months_avg = monthly_totals.tail(3).mean()
        last_6_months_avg = monthly_totals.tail(6).mean()
        last_12_months_avg = monthly_totals.mean()
        
        for i in range(1, months_ahead + 1):
            forecast_month = end_date + timedelta(days=30*i)
            
            # Weight recent data more heavily
            forecast_amount = (
                last_3_months_avg * 0.5 +
                last_6_months_avg * 0.3 +
                last_12_months_avg * 0.2
            )
            
            # Add seasonal adjustment (simplified)
            month_num = forecast_month.month
            seasonal_factor = 1.0
            if month_num in [11, 12]:  # Higher expenses in Nov/Dec
                seasonal_factor = 1.15
            elif month_num in [7, 8]:  # Lower in summer
                seasonal_factor = 0.9
            
            forecast_data.append({
                'Month': forecast_month.strftime('%Y-%m'),
                'Forecast': round(forecast_amount * seasonal_factor, 2),
                'Low Estimate': round(forecast_amount * seasonal_factor * 0.9, 2),
                'High Estimate': round(forecast_amount * seasonal_factor * 1.1, 2)
            })
        
        return pd.DataFrame(forecast_data)
    
    def identify_cost_savings(self, lookback_months: int = 6) -> List[Dict]:
        """Identify potential cost savings opportunities"""
        end_date = date.today()
        start_date = end_date - timedelta(days=30 * lookback_months)
        
        recent_expenses = [
            e for e in self.expenses 
            if start_date <= e.date <= end_date
        ]
        
        suggestions = []
        
        # 1. Duplicate vendor analysis
        vendor_category_spend = {}
        for expense in recent_expenses:
            vendor = self.vendors.get(expense.vendor_id)
            if vendor:
                key = (expense.category, vendor.vendor_id)
                if key not in vendor_category_spend:
                    vendor_category_spend[key] = {'vendor': vendor.name, 'total': 0, 'count': 0}
                vendor_category_spend[key]['total'] += expense.amount
                vendor_category_spend[key]['count'] += 1
        
        # Find categories with multiple vendors
        category_vendors = {}
        for (category, vendor_id), data in vendor_category_spend.items():
            if category not in category_vendors:
                category_vendors[category] = []
            category_vendors[category].append(data)
        
        for category, vendors in category_vendors.items():
            if len(vendors) > 2:
                total_spend = sum(v['total'] for v in vendors)
                suggestions.append({
                    'type': 'Vendor Consolidation',
                    'category': category.value,
                    'description': f"Consider consolidating {len(vendors)} vendors",
                    'current_spend': round(total_spend, 2),
                    'potential_savings': round(total_spend * 0.1, 2),  # 10% savings estimate
                    'vendors': [v['vendor'] for v in vendors]
                })
        
        # 2. Recurring expense optimization
        recurring_expenses = [e for e in recent_expenses if e.recurring]
        recurring_by_category = {}
        for expense in recurring_expenses:
            if expense.category not in recurring_by_category:
                recurring_by_category[expense.category] = []
            recurring_by_category[expense.category].append(expense)
        
        for category, expenses in recurring_by_category.items():
            if len(expenses) > 5:
                total = sum(e.amount for e in expenses)
                suggestions.append({
                    'type': 'Recurring Expense Review',
                    'category': category.value,
                    'description': f"Review {len(expenses)} recurring expenses",
                    'current_spend': round(total, 2),
                    'potential_savings': round(total * 0.15, 2),  # 15% savings estimate
                    'action': 'Negotiate annual contracts or review necessity'
                })
        
        # 3. Outlier detection
        df = pd.DataFrame([{
            'amount': e.amount,
            'category': e.category.value,
            'vendor': self.vendors.get(e.vendor_id, Vendor("", "Unknown", {})).name
        } for e in recent_expenses])
        
        if not df.empty:
            category_stats = df.groupby('category')['amount'].agg(['mean', 'std', 'count'])
            
            for expense in recent_expenses:
                cat_stats = category_stats.loc[expense.category.value]
                if cat_stats['count'] > 5:  # Need enough data
                    z_score = (expense.amount - cat_stats['mean']) / cat_stats['std'] if cat_stats['std'] > 0 else 0
                    if z_score > 3:  # 3 standard deviations
                        vendor = self.vendors.get(expense.vendor_id, Vendor("", "Unknown", {}))
                        suggestions.append({
                            'type': 'Unusual Expense',
                            'category': expense.category.value,
                            'description': f"Expense significantly above average",
                            'vendor': vendor.name,
                            'amount': expense.amount,
                            'average': round(cat_stats['mean'], 2),
                            'action': 'Review for accuracy or negotiate'
                        })
        
        return suggestions
    
    def generate_1099_report(self, tax_year: int) -> pd.DataFrame:
        """Generate 1099 reporting for vendors"""
        start_date = date(tax_year, 1, 1)
        end_date = date(tax_year, 12, 31)
        
        vendor_payments = {}
        
        for expense in self.expenses:
            if (start_date <= expense.date <= end_date and 
                expense.approval_status == ApprovalStatus.PAID):
                
                if expense.vendor_id not in vendor_payments:
                    vendor_payments[expense.vendor_id] = {
                        'total': 0,
                        'payments': []
                    }
                
                vendor_payments[expense.vendor_id]['total'] += expense.amount
                vendor_payments[expense.vendor_id]['payments'].append({
                    'date': expense.date,
                    'amount': expense.amount,
                    'category': expense.category.value
                })
        
        # Create 1099 report
        report_data = []
        for vendor_id, payment_data in vendor_payments.items():
            vendor = self.vendors.get(vendor_id)
            if vendor and payment_data['total'] >= 600:  # 1099 threshold
                report_data.append({
                    'Vendor Name': vendor.name,
                    'Tax ID': vendor.tax_id or 'Missing',
                    'Total Payments': round(payment_data['total'], 2),
                    'Payment Count': len(payment_data['payments']),
                    'W9 on File': vendor.w9_on_file,
                    'Status': 'Ready' if vendor.tax_id and vendor.w9_on_file else 'Incomplete'
                })
        
        return pd.DataFrame(report_data).sort_values('Total Payments', ascending=False)
    
    def cash_flow_impact(self, days_ahead: int = 30) -> Dict[str, float]:
        """Analyze cash flow impact of pending expenses"""
        today = date.today()
        future_date = today + timedelta(days=days_ahead)
        
        # Pending approvals
        pending_expenses = [
            e for e in self.expenses 
            if e.approval_status in [ApprovalStatus.PENDING, ApprovalStatus.APPROVED]
            and e.paid_date is None
        ]
        
        # Group by expected payment date
        cash_flow = {}
        
        for expense in pending_expenses:
            vendor = self.vendors.get(expense.vendor_id)
            if vendor:
                # Calculate expected payment date based on terms
                terms_days = 30  # Default
                if vendor.payment_terms == "Net 15":
                    terms_days = 15
                elif vendor.payment_terms == "Net 45":
                    terms_days = 45
                elif vendor.payment_terms == "Due on Receipt":
                    terms_days = 0
                
                expected_payment = expense.date + timedelta(days=terms_days)
                
                if expected_payment <= future_date:
                    week_key = expected_payment.strftime('%Y-W%V')
                    if week_key not in cash_flow:
                        cash_flow[week_key] = 0
                    cash_flow[week_key] += expense.amount
        
        return {
            'weekly_outflows': cash_flow,
            'total_pending': sum(e.amount for e in pending_expenses),
            'next_30_days': sum(cash_flow.values()),
            'overdue': sum(e.amount for e in pending_expenses if e.date < today - timedelta(days=30))
        }

class ExpenseAnalytics:
    """Advanced analytics for expense data"""
    
    @staticmethod
    def spending_trends(expenses: List[Expense], periods: int = 12) -> pd.DataFrame:
        """Analyze spending trends over time"""
        if not expenses:
            return pd.DataFrame()
        
        df = pd.DataFrame([{
            'Date': e.date,
            'Amount': e.amount,
            'Category': e.category.value
        } for e in expenses])
        
        df['Month'] = pd.to_datetime(df['Date']).dt.to_period('M')
        
        # Overall trend
        monthly_trend = df.groupby('Month')['Amount'].sum()
        
        # Calculate moving averages
        ma_3 = monthly_trend.rolling(window=3).mean()
        ma_6 = monthly_trend.rolling(window=6).mean()
        
        # Calculate growth rates
        growth_rate = monthly_trend.pct_change()
        
        trend_df = pd.DataFrame({
            'Month': monthly_trend.index,
            'Total': monthly_trend.values,
            '3-Month MA': ma_3.values,
            '6-Month MA': ma_6.values,
            'Growth Rate': growth_rate.values
        })
        
        return trend_df
    
    @staticmethod
    def vendor_analysis(expenses: List[Expense], vendors: Dict[str, Vendor]) -> pd.DataFrame:
        """Analyze vendor spending patterns"""
        vendor_data = {}
        
        for expense in expenses:
            if expense.vendor_id not in vendor_data:
                vendor = vendors.get(expense.vendor_id, Vendor("", "Unknown", {}))
                vendor_data[expense.vendor_id] = {
                    'name': vendor.name,
                    'total_spend': 0,
                    'transaction_count': 0,
                    'categories': set(),
                    'last_payment': None,
                    'payment_methods': set()
                }
            
            vd = vendor_data[expense.vendor_id]
            vd['total_spend'] += expense.amount
            vd['transaction_count'] += 1
            vd['categories'].add(expense.category.value)
            vd['payment_methods'].add(expense.payment_method.value)
            
            if vd['last_payment'] is None or expense.date > vd['last_payment']:
                vd['last_payment'] = expense.date
        
        # Convert to DataFrame
        analysis_data = []
        for vendor_id, data in vendor_data.items():
            analysis_data.append({
                'Vendor': data['name'],
                'Total Spend': data['total_spend'],
                'Transactions': data['transaction_count'],
                'Avg Transaction': data['total_spend'] / data['transaction_count'],
                'Categories': len(data['categories']),
                'Last Payment': data['last_payment'],
                'Days Since Last': (date.today() - data['last_payment']).days if data['last_payment'] else None
            })
        
        return pd.DataFrame(analysis_data).sort_values('Total Spend', ascending=False)

if __name__ == "__main__":
    # Test expense tracking system
    tracker = ExpenseTracker()
    
    # Add sample vendor
    vendor = Vendor("V001", "Office Supplies Co", {"phone": "555-1234"}, "12-3456789")
    tracker.vendors[vendor.vendor_id] = vendor
    
    # Add sample expense
    expense = Expense(
        "E001", 
        date.today(), 
        "V001", 
        150.00, 
        ExpenseCategory.OFFICE_SUPPLIES,
        description="Monthly office supplies"
    )
    tracker.add_expense(expense)
    
    # Test summary
    print("Expense Summary:")
    print(tracker.get_expense_summary(date.today() - timedelta(days=30), date.today()))