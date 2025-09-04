import pandas as pd
import numpy as np
from datetime import datetime, date, timedelta
from dateutil.relativedelta import relativedelta
from typing import Dict, List, Optional, Tuple, Union
from dataclasses import dataclass, field
from enum import Enum
import json

class CashFlowType(Enum):
    OPERATING = "Operating"
    INVESTING = "Investing" 
    FINANCING = "Financing"

class CashFlowDirection(Enum):
    INFLOW = "Inflow"
    OUTFLOW = "Outflow"

@dataclass
class CashFlowItem:
    item_id: str
    date: date
    description: str
    amount: float
    flow_type: CashFlowType
    direction: CashFlowDirection
    category: str
    subcategory: Optional[str] = None
    project_id: Optional[str] = None
    recurring: bool = False
    recurring_frequency: Optional[str] = None
    notes: Optional[str] = None

@dataclass
class CashPosition:
    date: date
    opening_balance: float
    total_inflows: float
    total_outflows: float
    closing_balance: float
    operating_cash_flow: float
    investing_cash_flow: float
    financing_cash_flow: float

class CashFlowAnalyzer:
    """Comprehensive cash flow analysis and forecasting"""
    
    def __init__(self):
        self.cash_flows: List[CashFlowItem] = []
        self.opening_balance: float = 0
        self.bank_accounts: Dict[str, float] = {}
        self.credit_lines: Dict[str, Dict] = {}
    
    def add_cash_flow_item(self, item: CashFlowItem) -> str:
        """Add a cash flow item"""
        self.cash_flows.append(item)
        return item.item_id
    
    def generate_cash_flow_statement(self, start_date: date, end_date: date) -> Dict:
        """Generate formal cash flow statement"""
        period_flows = [
            cf for cf in self.cash_flows 
            if start_date <= cf.date <= end_date
        ]
        
        # Categorize cash flows
        operating_flows = [cf for cf in period_flows if cf.flow_type == CashFlowType.OPERATING]
        investing_flows = [cf for cf in period_flows if cf.flow_type == CashFlowType.INVESTING]
        financing_flows = [cf for cf in period_flows if cf.flow_type == CashFlowType.FINANCING]
        
        # Calculate net flows
        def calculate_net_flow(flows: List[CashFlowItem]) -> float:
            inflows = sum(cf.amount for cf in flows if cf.direction == CashFlowDirection.INFLOW)
            outflows = sum(cf.amount for cf in flows if cf.direction == CashFlowDirection.OUTFLOW)
            return inflows - outflows
        
        operating_cash_flow = calculate_net_flow(operating_flows)
        investing_cash_flow = calculate_net_flow(investing_flows)
        financing_cash_flow = calculate_net_flow(financing_flows)
        
        net_change_in_cash = operating_cash_flow + investing_cash_flow + financing_cash_flow
        
        # Detailed breakdown
        operating_detail = self._categorize_flows(operating_flows)
        investing_detail = self._categorize_flows(investing_flows)
        financing_detail = self._categorize_flows(financing_flows)
        
        return {
            'period': f"{start_date} to {end_date}",
            'operating_activities': {
                'total': round(operating_cash_flow, 2),
                'detail': operating_detail
            },
            'investing_activities': {
                'total': round(investing_cash_flow, 2),
                'detail': investing_detail
            },
            'financing_activities': {
                'total': round(financing_cash_flow, 2),
                'detail': financing_detail
            },
            'net_change_in_cash': round(net_change_in_cash, 2),
            'opening_cash_balance': self.opening_balance,
            'closing_cash_balance': round(self.opening_balance + net_change_in_cash, 2)
        }
    
    def _categorize_flows(self, flows: List[CashFlowItem]) -> Dict[str, Dict]:
        """Categorize and aggregate cash flows"""
        categories = {}
        
        for flow in flows:
            category = flow.category
            if category not in categories:
                categories[category] = {
                    'inflows': 0,
                    'outflows': 0,
                    'net': 0,
                    'items': []
                }
            
            if flow.direction == CashFlowDirection.INFLOW:
                categories[category]['inflows'] += flow.amount
            else:
                categories[category]['outflows'] += flow.amount
            
            categories[category]['net'] = categories[category]['inflows'] - categories[category]['outflows']
            categories[category]['items'].append({
                'date': flow.date.isoformat(),
                'description': flow.description,
                'amount': flow.amount * (1 if flow.direction == CashFlowDirection.INFLOW else -1)
            })
        
        # Round values
        for category_data in categories.values():
            category_data['inflows'] = round(category_data['inflows'], 2)
            category_data['outflows'] = round(category_data['outflows'], 2)
            category_data['net'] = round(category_data['net'], 2)
        
        return categories
    
    def forecast_cash_flow(self, months_ahead: int = 12, 
                          scenarios: Optional[Dict[str, float]] = None) -> pd.DataFrame:
        """Forecast cash flows with scenario analysis"""
        if scenarios is None:
            scenarios = {'base': 1.0, 'conservative': 0.9, 'optimistic': 1.1}
        
        # Analyze historical patterns
        historical_flows = self._get_historical_patterns()
        
        forecasts = []
        current_balance = self.opening_balance + sum(
            cf.amount * (1 if cf.direction == CashFlowDirection.INFLOW else -1) 
            for cf in self.cash_flows
        )
        
        for month_offset in range(1, months_ahead + 1):
            forecast_date = date.today() + relativedelta(months=month_offset)
            
            # Base forecast from historical patterns
            base_operating = historical_flows['monthly_operating_avg']
            base_investing = historical_flows['monthly_investing_avg']
            base_financing = historical_flows['monthly_financing_avg']
            
            # Apply seasonal adjustments
            seasonal_factor = self._get_seasonal_factor(forecast_date.month)
            
            month_data = {
                'Month': forecast_date.strftime('%Y-%m'),
                'Date': forecast_date
            }
            
            for scenario, factor in scenarios.items():
                operating = base_operating * seasonal_factor * factor
                investing = base_investing * factor
                financing = base_financing * factor
                
                net_cash_flow = operating + investing + financing
                current_balance += net_cash_flow
                
                month_data[f'{scenario.title()}_Operating'] = round(operating, 2)
                month_data[f'{scenario.title()}_Investing'] = round(investing, 2)
                month_data[f'{scenario.title()}_Financing'] = round(financing, 2)
                month_data[f'{scenario.title()}_Net'] = round(net_cash_flow, 2)
                month_data[f'{scenario.title()}_Balance'] = round(current_balance, 2)
            
            forecasts.append(month_data)
        
        return pd.DataFrame(forecasts)
    
    def _get_historical_patterns(self) -> Dict[str, float]:
        """Analyze historical cash flow patterns"""
        if not self.cash_flows:
            return {
                'monthly_operating_avg': 0,
                'monthly_investing_avg': 0,
                'monthly_financing_avg': 0
            }
        
        # Group by type and calculate monthly averages
        operating_flows = [cf for cf in self.cash_flows if cf.flow_type == CashFlowType.OPERATING]
        investing_flows = [cf for cf in self.cash_flows if cf.flow_type == CashFlowType.INVESTING]
        financing_flows = [cf for cf in self.cash_flows if cf.flow_type == CashFlowType.FINANCING]
        
        def calculate_monthly_avg(flows: List[CashFlowItem]) -> float:
            if not flows:
                return 0
            
            df = pd.DataFrame([{
                'date': cf.date,
                'amount': cf.amount * (1 if cf.direction == CashFlowDirection.INFLOW else -1)
            } for cf in flows])
            
            df['month'] = pd.to_datetime(df['date']).dt.to_period('M')
            monthly_totals = df.groupby('month')['amount'].sum()
            
            return monthly_totals.mean()
        
        return {
            'monthly_operating_avg': calculate_monthly_avg(operating_flows),
            'monthly_investing_avg': calculate_monthly_avg(investing_flows),
            'monthly_financing_avg': calculate_monthly_avg(financing_flows)
        }
    
    def _get_seasonal_factor(self, month: int) -> float:
        """Get seasonal adjustment factor"""
        seasonal_factors = {
            1: 0.95,   # January - post-holiday slow
            2: 0.95,   # February - short month
            3: 1.05,   # March - Q1 close
            4: 1.0,    # April - normal
            5: 1.0,    # May - normal
            6: 1.05,   # June - Q2 close
            7: 0.9,    # July - summer slow
            8: 0.9,    # August - summer slow
            9: 1.05,   # September - Q3 close
            10: 1.0,   # October - normal
            11: 1.1,   # November - holiday prep
            12: 1.15   # December - holiday/year-end
        }
        return seasonal_factors.get(month, 1.0)
    
    def working_capital_analysis(self, start_date: date, end_date: date) -> Dict:
        """Analyze working capital changes"""
        # Simplified analysis - in practice would integrate with balance sheet data
        operating_flows = [
            cf for cf in self.cash_flows 
            if cf.flow_type == CashFlowType.OPERATING and start_date <= cf.date <= end_date
        ]
        
        revenue_flows = [cf for cf in operating_flows if 'revenue' in cf.category.lower()]
        expense_flows = [cf for cf in operating_flows if cf not in revenue_flows]
        
        total_revenue = sum(cf.amount for cf in revenue_flows if cf.direction == CashFlowDirection.INFLOW)
        total_expenses = sum(cf.amount for cf in expense_flows if cf.direction == CashFlowDirection.OUTFLOW)
        
        # Calculate key metrics
        operating_cash_flow = total_revenue - total_expenses
        cash_conversion_cycle = self._estimate_cash_conversion_cycle(operating_flows)
        
        return {
            'period': f"{start_date} to {end_date}",
            'total_revenue': round(total_revenue, 2),
            'total_operating_expenses': round(total_expenses, 2),
            'operating_cash_flow': round(operating_cash_flow, 2),
            'operating_margin': round(operating_cash_flow / total_revenue * 100, 2) if total_revenue > 0 else 0,
            'cash_conversion_cycle_days': round(cash_conversion_cycle, 1),
            'working_capital_efficiency': 'Good' if cash_conversion_cycle < 30 else 'Needs Improvement'
        }
    
    def _estimate_cash_conversion_cycle(self, flows: List[CashFlowItem]) -> float:
        """Estimate cash conversion cycle from cash flow timing"""
        # Simplified estimation - would need more detailed AR/AP data
        revenue_dates = [cf.date for cf in flows if 'revenue' in cf.category.lower()]
        payment_dates = [cf.date for cf in flows if 'payment' in cf.category.lower()]
        
        if not revenue_dates or not payment_dates:
            return 30  # Default assumption
        
        avg_collection_period = np.mean([(p - r).days for r, p in zip(revenue_dates, payment_dates) if (p - r).days > 0])
        return avg_collection_period if not np.isnan(avg_collection_period) else 30
    
    def cash_burn_analysis(self, months_back: int = 6) -> Dict:
        """Analyze cash burn rate and runway"""
        end_date = date.today()
        start_date = end_date - timedelta(days=30 * months_back)
        
        monthly_flows = {}
        
        # Group flows by month
        for cf in self.cash_flows:
            if start_date <= cf.date <= end_date:
                month_key = cf.date.strftime('%Y-%m')
                if month_key not in monthly_flows:
                    monthly_flows[month_key] = {'inflows': 0, 'outflows': 0}
                
                if cf.direction == CashFlowDirection.INFLOW:
                    monthly_flows[month_key]['inflows'] += cf.amount
                else:
                    monthly_flows[month_key]['outflows'] += cf.amount
        
        # Calculate burn rates
        monthly_burns = []
        for month, flows in monthly_flows.items():
            net_burn = flows['outflows'] - flows['inflows']
            monthly_burns.append({
                'month': month,
                'inflows': flows['inflows'],
                'outflows': flows['outflows'],
                'net_burn': net_burn
            })
        
        if not monthly_burns:
            return {'error': 'No cash flow data found'}
        
        df = pd.DataFrame(monthly_burns)
        average_burn = df['net_burn'].mean()
        current_cash = self.opening_balance + sum(
            cf.amount * (1 if cf.direction == CashFlowDirection.INFLOW else -1) 
            for cf in self.cash_flows
        )
        
        # Calculate runway
        runway_months = current_cash / average_burn if average_burn > 0 else float('inf')
        
        # Trend analysis
        if len(monthly_burns) >= 3:
            recent_3_months = df.tail(3)['net_burn'].mean()
            trend = 'Improving' if recent_3_months < average_burn else 'Worsening'
        else:
            trend = 'Insufficient Data'
        
        return {
            'current_cash_balance': round(current_cash, 2),
            'average_monthly_burn': round(average_burn, 2),
            'runway_months': round(runway_months, 1) if runway_months != float('inf') else 'Infinite',
            'trend': trend,
            'monthly_detail': monthly_burns,
            'burn_rate_volatility': round(df['net_burn'].std(), 2),
            'recommendation': self._get_burn_recommendation(runway_months, trend)
        }
    
    def _get_burn_recommendation(self, runway_months: float, trend: str) -> str:
        """Get recommendations based on burn analysis"""
        if runway_months == float('inf'):
            return "Cash positive - focus on growth investments"
        elif runway_months > 18:
            return "Healthy cash position - monitor trends"
        elif runway_months > 12:
            return "Adequate runway - consider efficiency improvements"
        elif runway_months > 6:
            return "Moderate concern - reduce burn or secure funding"
        else:
            return "Critical - immediate action required"
    
    def scenario_analysis(self, scenarios: Dict[str, Dict[str, float]], 
                         months_ahead: int = 12) -> pd.DataFrame:
        """Run scenario analysis on cash flow projections"""
        base_flows = self._get_historical_patterns()
        results = []
        
        for scenario_name, adjustments in scenarios.items():
            current_balance = self.opening_balance
            scenario_results = []
            
            for month_offset in range(1, months_ahead + 1):
                forecast_date = date.today() + relativedelta(months=month_offset)
                
                # Apply scenario adjustments
                operating_flow = base_flows['monthly_operating'] * adjustments.get('operating_multiplier', 1.0)
                investing_flow = base_flows['monthly_investing'] * adjustments.get('investing_multiplier', 1.0)
                financing_flow = base_flows['monthly_financing'] * adjustments.get('financing_multiplier', 1.0)
                
                # Add one-time adjustments
                one_time_adjustment = adjustments.get('one_time_cash_injection', 0) if month_offset == 1 else 0
                
                net_flow = operating_flow + investing_flow + financing_flow + one_time_adjustment
                current_balance += net_flow
                
                scenario_results.append({
                    'Scenario': scenario_name,
                    'Month': forecast_date.strftime('%Y-%m'),
                    'Operating_Flow': round(operating_flow, 2),
                    'Investing_Flow': round(investing_flow, 2),
                    'Financing_Flow': round(financing_flow, 2),
                    'Net_Flow': round(net_flow, 2),
                    'Cash_Balance': round(current_balance, 2),
                    'Months_Runway': round(current_balance / abs(operating_flow), 1) if operating_flow < 0 else float('inf')
                })
            
            results.extend(scenario_results)
        
        return pd.DataFrame(results)
    
    def _get_historical_patterns(self) -> Dict[str, float]:
        """Get historical cash flow patterns for forecasting"""
        if not self.cash_flows:
            return {
                'monthly_operating': 0,
                'monthly_investing': 0,
                'monthly_financing': 0
            }
        
        # Calculate monthly averages by type
        df = pd.DataFrame([{
            'date': cf.date,
            'flow_type': cf.flow_type.value,
            'amount': cf.amount * (1 if cf.direction == CashFlowDirection.INFLOW else -1)
        } for cf in self.cash_flows])
        
        df['month'] = pd.to_datetime(df['date']).dt.to_period('M')
        
        monthly_by_type = df.groupby(['month', 'flow_type'])['amount'].sum().unstack(fill_value=0)
        
        return {
            'monthly_operating': monthly_by_type.get('Operating', pd.Series([0])).mean(),
            'monthly_investing': monthly_by_type.get('Investing', pd.Series([0])).mean(),
            'monthly_financing': monthly_by_type.get('Financing', pd.Series([0])).mean()
        }
    
    def cash_flow_at_risk(self, confidence_level: float = 0.95) -> Dict:
        """Calculate Cash Flow at Risk (CFaR) - similar to VaR for investments"""
        if len(self.cash_flows) < 30:  # Need sufficient data
            return {'error': 'Insufficient historical data for CFaR calculation'}
        
        # Calculate daily cash flows
        daily_flows = {}
        for cf in self.cash_flows:
            date_key = cf.date.isoformat()
            if date_key not in daily_flows:
                daily_flows[date_key] = 0
            
            flow_amount = cf.amount * (1 if cf.direction == CashFlowDirection.INFLOW else -1)
            daily_flows[date_key] += flow_amount
        
        # Calculate percentiles
        flow_values = list(daily_flows.values())
        var_percentile = (1 - confidence_level) * 100
        
        cash_flow_at_risk = np.percentile(flow_values, var_percentile)
        expected_shortfall = np.mean([f for f in flow_values if f <= cash_flow_at_risk])
        
        return {
            'confidence_level': f"{confidence_level * 100}%",
            'cash_flow_at_risk': round(cash_flow_at_risk, 2),
            'expected_shortfall': round(expected_shortfall, 2),
            'volatility': round(np.std(flow_values), 2),
            'interpretation': f"With {confidence_level * 100}% confidence, daily cash flow will not be worse than ${abs(cash_flow_at_risk):,.2f}"
        }
    
    def liquidity_analysis(self) -> Dict:
        """Analyze liquidity position and requirements"""
        current_cash = sum(self.bank_accounts.values())
        available_credit = sum(
            line['limit'] - line['outstanding'] 
            for line in self.credit_lines.values()
        )
        
        # Calculate minimum cash requirement (simplified)
        monthly_outflows = abs(self._get_historical_patterns()['monthly_operating'])
        minimum_cash_requirement = monthly_outflows * 1.5  # 1.5 months of operating expenses
        
        # Analyze cash sources
        quick_assets = current_cash  # Would include other liquid assets
        total_liquidity = current_cash + available_credit
        
        return {
            'current_cash': round(current_cash, 2),
            'available_credit': round(available_credit, 2),
            'total_liquidity': round(total_liquidity, 2),
            'minimum_cash_requirement': round(minimum_cash_requirement, 2),
            'liquidity_ratio': round(total_liquidity / minimum_cash_requirement, 2) if minimum_cash_requirement > 0 else float('inf'),
            'liquidity_status': self._get_liquidity_status(total_liquidity, minimum_cash_requirement),
            'days_cash_on_hand': round(current_cash / (monthly_outflows / 30), 1) if monthly_outflows > 0 else float('inf')
        }
    
    def _get_liquidity_status(self, total_liquidity: float, minimum_requirement: float) -> str:
        """Determine liquidity status"""
        ratio = total_liquidity / minimum_requirement if minimum_requirement > 0 else float('inf')
        
        if ratio >= 3.0:
            return "Excellent"
        elif ratio >= 2.0:
            return "Good"
        elif ratio >= 1.5:
            return "Adequate"
        elif ratio >= 1.0:
            return "Tight"
        else:
            return "Critical"

if __name__ == "__main__":
    # Test cash flow analysis
    analyzer = CashFlowAnalyzer()
    analyzer.opening_balance = 100000
    
    # Add sample cash flows
    flows = [
        CashFlowItem("CF001", date.today() - timedelta(days=30), "Revenue", 50000, 
                    CashFlowType.OPERATING, CashFlowDirection.INFLOW, "Sales"),
        CashFlowItem("CF002", date.today() - timedelta(days=25), "Rent", 5000, 
                    CashFlowType.OPERATING, CashFlowDirection.OUTFLOW, "Facilities"),
        CashFlowItem("CF003", date.today() - timedelta(days=20), "Salaries", 15000, 
                    CashFlowType.OPERATING, CashFlowDirection.OUTFLOW, "Payroll")
    ]
    
    for flow in flows:
        analyzer.add_cash_flow_item(flow)
    
    # Test cash flow statement
    print("Cash Flow Statement:")
    statement = analyzer.generate_cash_flow_statement(
        date.today() - timedelta(days=30), 
        date.today()
    )
    print(json.dumps(statement, indent=2, default=str))