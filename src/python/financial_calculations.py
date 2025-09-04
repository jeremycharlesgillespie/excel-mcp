import numpy as np
import numpy_financial as npf
from decimal import Decimal, ROUND_HALF_UP
from typing import List, Dict, Union, Optional, Tuple
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
import pandas as pd

class FinancialCalculator:
    """High-level financial calculations for accounting and finance"""
    
    @staticmethod
    def npv(rate: float, cash_flows: List[float], initial_investment: float = 0) -> float:
        """Calculate Net Present Value"""
        flows = [-initial_investment] + cash_flows if initial_investment else cash_flows
        return float(npf.npv(rate, flows))
    
    @staticmethod
    def irr(cash_flows: List[float]) -> float:
        """Calculate Internal Rate of Return"""
        return float(npf.irr(cash_flows))
    
    @staticmethod
    def mirr(cash_flows: List[float], finance_rate: float, reinvest_rate: float) -> float:
        """Calculate Modified Internal Rate of Return"""
        return float(npf.mirr(cash_flows, finance_rate, reinvest_rate))
    
    @staticmethod
    def payback_period(initial_investment: float, cash_flows: List[float]) -> Optional[float]:
        """Calculate payback period in years"""
        cumulative = -initial_investment
        for i, flow in enumerate(cash_flows):
            cumulative += flow
            if cumulative >= 0:
                if flow != 0:
                    return i + (cumulative - flow) / flow
                else:
                    return float(i)
        return None
    
    @staticmethod
    def discounted_payback_period(initial_investment: float, cash_flows: List[float], rate: float) -> Optional[float]:
        """Calculate discounted payback period"""
        cumulative = -initial_investment
        for i, flow in enumerate(cash_flows):
            discounted_flow = flow / ((1 + rate) ** (i + 1))
            cumulative += discounted_flow
            if cumulative >= 0:
                if discounted_flow != 0:
                    return i + (cumulative - discounted_flow) / discounted_flow
                else:
                    return float(i)
        return None
    
    @staticmethod
    def profitability_index(initial_investment: float, cash_flows: List[float], rate: float) -> float:
        """Calculate profitability index (PI)"""
        pv_future_flows = sum(flow / ((1 + rate) ** (i + 1)) for i, flow in enumerate(cash_flows))
        return pv_future_flows / initial_investment
    
    @staticmethod
    def loan_amortization(principal: float, annual_rate: float, years: int, 
                         payment_frequency: str = 'monthly') -> pd.DataFrame:
        """Generate loan amortization schedule"""
        frequencies = {
            'monthly': 12,
            'quarterly': 4,
            'semi-annually': 2,
            'annually': 1
        }
        
        periods_per_year = frequencies.get(payment_frequency, 12)
        n_periods = years * periods_per_year
        period_rate = annual_rate / periods_per_year
        
        payment = npf.pmt(period_rate, n_periods, -principal)
        
        schedule = []
        balance = principal
        
        for period in range(1, n_periods + 1):
            interest = balance * period_rate
            principal_payment = payment - interest
            balance -= principal_payment
            
            schedule.append({
                'Period': period,
                'Payment': round(payment, 2),
                'Principal': round(principal_payment, 2),
                'Interest': round(interest, 2),
                'Balance': round(max(0, balance), 2)
            })
        
        return pd.DataFrame(schedule)
    
    @staticmethod
    def effective_annual_rate(nominal_rate: float, compounding_periods: int) -> float:
        """Calculate effective annual rate (EAR)"""
        return ((1 + nominal_rate / compounding_periods) ** compounding_periods) - 1
    
    @staticmethod
    def future_value(present_value: float, rate: float, periods: int) -> float:
        """Calculate future value"""
        return present_value * ((1 + rate) ** periods)
    
    @staticmethod
    def present_value(future_value: float, rate: float, periods: int) -> float:
        """Calculate present value"""
        return future_value / ((1 + rate) ** periods)
    
    @staticmethod
    def bond_price(face_value: float, coupon_rate: float, yield_rate: float, 
                   years: int, frequency: int = 2) -> float:
        """Calculate bond price"""
        periods = years * frequency
        coupon_payment = (face_value * coupon_rate) / frequency
        period_yield = yield_rate / frequency
        
        pv_coupons = sum(coupon_payment / ((1 + period_yield) ** i) 
                        for i in range(1, periods + 1))
        pv_face = face_value / ((1 + period_yield) ** periods)
        
        return pv_coupons + pv_face
    
    @staticmethod
    def duration(cash_flows: List[Tuple[float, float]], yield_rate: float) -> float:
        """Calculate Macaulay duration"""
        total_pv = 0
        weighted_pv = 0
        
        for time, cash_flow in cash_flows:
            pv = cash_flow / ((1 + yield_rate) ** time)
            total_pv += pv
            weighted_pv += time * pv
        
        return weighted_pv / total_pv if total_pv != 0 else 0
    
    @staticmethod
    def capm(risk_free_rate: float, beta: float, market_return: float) -> float:
        """Calculate expected return using CAPM"""
        return risk_free_rate + beta * (market_return - risk_free_rate)
    
    @staticmethod
    def wacc(equity_value: float, debt_value: float, cost_of_equity: float, 
             cost_of_debt: float, tax_rate: float) -> float:
        """Calculate Weighted Average Cost of Capital"""
        total_value = equity_value + debt_value
        equity_weight = equity_value / total_value
        debt_weight = debt_value / total_value
        
        return (equity_weight * cost_of_equity) + (debt_weight * cost_of_debt * (1 - tax_rate))


class DepreciationCalculator:
    """Calculate various depreciation methods"""
    
    @staticmethod
    def straight_line(cost: float, salvage_value: float, useful_life: int) -> List[float]:
        """Calculate straight-line depreciation"""
        annual_depreciation = (cost - salvage_value) / useful_life
        return [annual_depreciation] * useful_life
    
    @staticmethod
    def declining_balance(cost: float, salvage_value: float, useful_life: int, 
                         rate: float = 2.0) -> List[float]:
        """Calculate declining balance depreciation"""
        depreciation_schedule = []
        book_value = cost
        annual_rate = rate / useful_life
        
        for year in range(useful_life):
            depreciation = book_value * annual_rate
            if book_value - depreciation < salvage_value:
                depreciation = book_value - salvage_value
            depreciation_schedule.append(depreciation)
            book_value -= depreciation
            
            if book_value <= salvage_value:
                break
        
        while len(depreciation_schedule) < useful_life:
            depreciation_schedule.append(0)
            
        return depreciation_schedule
    
    @staticmethod
    def sum_of_years_digits(cost: float, salvage_value: float, useful_life: int) -> List[float]:
        """Calculate sum-of-years-digits depreciation"""
        depreciable_amount = cost - salvage_value
        sum_of_years = sum(range(1, useful_life + 1))
        
        depreciation_schedule = []
        for year in range(useful_life, 0, -1):
            fraction = year / sum_of_years
            depreciation = depreciable_amount * fraction
            depreciation_schedule.append(depreciation)
            
        return depreciation_schedule
    
    @staticmethod
    def units_of_production(cost: float, salvage_value: float, total_units: int, 
                           units_per_period: List[int]) -> List[float]:
        """Calculate units of production depreciation"""
        depreciable_amount = cost - salvage_value
        rate_per_unit = depreciable_amount / total_units
        
        return [units * rate_per_unit for units in units_per_period]
    
    @staticmethod
    def macrs(cost: float, recovery_period: int) -> List[float]:
        """Calculate MACRS depreciation (simplified)"""
        macrs_rates = {
            3: [0.3333, 0.4445, 0.1481, 0.0741],
            5: [0.2000, 0.3200, 0.1920, 0.1152, 0.1152, 0.0576],
            7: [0.1429, 0.2449, 0.1749, 0.1249, 0.0893, 0.0892, 0.0893, 0.0446],
            10: [0.1000, 0.1800, 0.1440, 0.1152, 0.0922, 0.0737, 0.0655, 0.0655, 0.0656, 0.0655, 0.0328]
        }
        
        if recovery_period not in macrs_rates:
            raise ValueError(f"MACRS rates not available for {recovery_period} year period")
        
        rates = macrs_rates[recovery_period]
        return [cost * rate for rate in rates]


class RatioAnalysis:
    """Financial ratio analysis for accounting"""
    
    @staticmethod
    def current_ratio(current_assets: float, current_liabilities: float) -> float:
        """Calculate current ratio"""
        return current_assets / current_liabilities if current_liabilities != 0 else float('inf')
    
    @staticmethod
    def quick_ratio(current_assets: float, inventory: float, current_liabilities: float) -> float:
        """Calculate quick ratio (acid test)"""
        return (current_assets - inventory) / current_liabilities if current_liabilities != 0 else float('inf')
    
    @staticmethod
    def debt_to_equity(total_debt: float, total_equity: float) -> float:
        """Calculate debt-to-equity ratio"""
        return total_debt / total_equity if total_equity != 0 else float('inf')
    
    @staticmethod
    def return_on_assets(net_income: float, total_assets: float) -> float:
        """Calculate ROA"""
        return net_income / total_assets if total_assets != 0 else 0
    
    @staticmethod
    def return_on_equity(net_income: float, shareholders_equity: float) -> float:
        """Calculate ROE"""
        return net_income / shareholders_equity if shareholders_equity != 0 else 0
    
    @staticmethod
    def gross_margin(revenue: float, cogs: float) -> float:
        """Calculate gross margin percentage"""
        return ((revenue - cogs) / revenue * 100) if revenue != 0 else 0
    
    @staticmethod
    def operating_margin(operating_income: float, revenue: float) -> float:
        """Calculate operating margin percentage"""
        return (operating_income / revenue * 100) if revenue != 0 else 0
    
    @staticmethod
    def net_margin(net_income: float, revenue: float) -> float:
        """Calculate net profit margin percentage"""
        return (net_income / revenue * 100) if revenue != 0 else 0
    
    @staticmethod
    def inventory_turnover(cogs: float, average_inventory: float) -> float:
        """Calculate inventory turnover ratio"""
        return cogs / average_inventory if average_inventory != 0 else 0
    
    @staticmethod
    def days_sales_outstanding(accounts_receivable: float, annual_sales: float) -> float:
        """Calculate DSO"""
        return (accounts_receivable / annual_sales * 365) if annual_sales != 0 else 0
    
    @staticmethod
    def asset_turnover(revenue: float, average_total_assets: float) -> float:
        """Calculate asset turnover ratio"""
        return revenue / average_total_assets if average_total_assets != 0 else 0
    
    @staticmethod
    def interest_coverage(ebit: float, interest_expense: float) -> float:
        """Calculate interest coverage ratio"""
        return ebit / interest_expense if interest_expense != 0 else float('inf')


if __name__ == "__main__":
    # Test calculations
    calc = FinancialCalculator()
    print("NPV:", calc.npv(0.1, [100, 200, 300], 500))
    print("IRR:", calc.irr([-1000, 400, 400, 400]))
    
    # Test depreciation
    dep = DepreciationCalculator()
    print("Straight Line:", dep.straight_line(10000, 1000, 5))
    
    # Test ratios
    ratios = RatioAnalysis()
    print("Current Ratio:", ratios.current_ratio(50000, 25000))