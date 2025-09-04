import pandas as pd
import numpy as np
from datetime import datetime, date, timedelta
from typing import Dict, List, Optional, Tuple, Union
from dataclasses import dataclass
from enum import Enum
import json

class TaxEntityType(Enum):
    SOLE_PROPRIETORSHIP = "Sole Proprietorship"
    PARTNERSHIP = "Partnership"
    S_CORP = "S Corporation"
    C_CORP = "C Corporation"
    LLC = "LLC"

class DepreciationMethod(Enum):
    STRAIGHT_LINE = "Straight Line"
    MACRS = "MACRS"
    SECTION_179 = "Section 179"
    BONUS_DEPRECIATION = "Bonus Depreciation"

@dataclass
class TaxableEntity:
    entity_id: str
    name: str
    entity_type: TaxEntityType
    tax_id: str
    fiscal_year_end: str  # MM-DD format
    state: str
    
@dataclass
class DepreciableAsset:
    asset_id: str
    description: str
    placed_in_service_date: date
    cost: float
    useful_life: int
    depreciation_method: DepreciationMethod
    section_179_election: bool = False
    bonus_depreciation_rate: float = 0.0
    accumulated_depreciation: float = 0.0

class TaxCalculator:
    """Tax calculations and reporting for various entity types"""
    
    def __init__(self):
        self.entities: Dict[str, TaxableEntity] = {}
        self.assets: Dict[str, DepreciableAsset] = {}
        
        # 2024 tax brackets (would be updated annually)
        self.federal_brackets_single = [
            (11000, 0.10),
            (44725, 0.12),
            (95375, 0.22),
            (182050, 0.24),
            (231250, 0.32),
            (578125, 0.35),
            (float('inf'), 0.37)
        ]
        
        self.federal_brackets_mfj = [
            (22000, 0.10),
            (89450, 0.12),
            (190750, 0.22),
            (364200, 0.24),
            (462500, 0.32),
            (693750, 0.35),
            (float('inf'), 0.37)
        ]
        
        # Standard deductions for 2024
        self.standard_deductions = {
            'single': 13850,
            'married_filing_jointly': 27700,
            'married_filing_separately': 13850,
            'head_of_household': 20800
        }
    
    def calculate_federal_income_tax(self, taxable_income: float, filing_status: str = 'single') -> Dict[str, float]:
        """Calculate federal income tax"""
        if filing_status == 'married_filing_jointly':
            brackets = self.federal_brackets_mfj
        else:
            brackets = self.federal_brackets_single
        
        tax_owed = 0
        tax_calculation_detail = []
        remaining_income = taxable_income
        
        prev_threshold = 0
        for threshold, rate in brackets:
            if remaining_income <= 0:
                break
                
            taxable_at_rate = min(remaining_income, threshold - prev_threshold)
            tax_at_rate = taxable_at_rate * rate
            tax_owed += tax_at_rate
            
            if taxable_at_rate > 0:
                tax_calculation_detail.append({
                    'income_range': f"${prev_threshold:,.0f} - ${min(threshold, prev_threshold + remaining_income):,.0f}",
                    'rate': f"{rate * 100:.1f}%",
                    'taxable_income': taxable_at_rate,
                    'tax': tax_at_rate
                })
            
            remaining_income -= taxable_at_rate
            prev_threshold = threshold
        
        return {
            'taxable_income': taxable_income,
            'total_tax': round(tax_owed, 2),
            'effective_rate': round(tax_owed / taxable_income * 100, 2) if taxable_income > 0 else 0,
            'marginal_rate': round(next(rate for threshold, rate in brackets if taxable_income <= threshold) * 100, 2),
            'calculation_detail': tax_calculation_detail
        }
    
    def calculate_self_employment_tax(self, net_earnings: float) -> Dict[str, float]:
        """Calculate self-employment tax (Social Security and Medicare)"""
        if net_earnings <= 0:
            return {'se_tax': 0, 'social_security': 0, 'medicare': 0}
        
        # 2024 rates and limits
        ss_wage_base = 160200  # Social Security wage base
        ss_rate = 0.124  # 12.4%
        medicare_rate = 0.029  # 2.9%
        additional_medicare_rate = 0.009  # Additional 0.9% on high earners
        additional_medicare_threshold = 200000
        
        # Calculate SE earnings (92.35% of net earnings)
        se_earnings = net_earnings * 0.9235
        
        # Social Security tax
        ss_taxable = min(se_earnings, ss_wage_base)
        ss_tax = ss_taxable * ss_rate
        
        # Medicare tax
        medicare_tax = se_earnings * medicare_rate
        
        # Additional Medicare tax
        additional_medicare = 0
        if se_earnings > additional_medicare_threshold:
            additional_medicare = (se_earnings - additional_medicare_threshold) * additional_medicare_rate
        
        total_se_tax = ss_tax + medicare_tax + additional_medicare
        
        return {
            'se_earnings': round(se_earnings, 2),
            'social_security_tax': round(ss_tax, 2),
            'medicare_tax': round(medicare_tax, 2),
            'additional_medicare_tax': round(additional_medicare, 2),
            'total_se_tax': round(total_se_tax, 2),
            'deductible_portion': round(total_se_tax * 0.5, 2)  # 50% deductible
        }
    
    def calculate_depreciation_deduction(self, asset_id: str, tax_year: int) -> Dict[str, float]:
        """Calculate depreciation deduction for an asset"""
        asset = self.assets.get(asset_id)
        if not asset:
            return {'error': 'Asset not found'}
        
        years_in_service = tax_year - asset.placed_in_service_date.year
        
        if asset.depreciation_method == DepreciationMethod.SECTION_179:
            # Section 179 immediate expensing (2024 limit: $1,220,000)
            section_179_limit = 1220000
            deduction = min(asset.cost, section_179_limit)
            
            return {
                'asset_id': asset_id,
                'depreciation_method': 'Section 179',
                'annual_deduction': deduction,
                'remaining_basis': 0,
                'total_depreciation': deduction
            }
        
        elif asset.depreciation_method == DepreciationMethod.MACRS:
            # MACRS depreciation
            from financial_calculations import DepreciationCalculator
            
            macrs_schedule = DepreciationCalculator.macrs(asset.cost, asset.useful_life)
            
            if years_in_service < len(macrs_schedule):
                annual_deduction = macrs_schedule[years_in_service]
                total_depreciation = sum(macrs_schedule[:years_in_service + 1])
            else:
                annual_deduction = 0
                total_depreciation = sum(macrs_schedule)
            
            return {
                'asset_id': asset_id,
                'depreciation_method': 'MACRS',
                'annual_deduction': round(annual_deduction, 2),
                'remaining_basis': round(asset.cost - total_depreciation, 2),
                'total_depreciation': round(total_depreciation, 2)
            }
        
        elif asset.depreciation_method == DepreciationMethod.STRAIGHT_LINE:
            annual_deduction = asset.cost / asset.useful_life
            total_depreciation = annual_deduction * min(years_in_service + 1, asset.useful_life)
            
            return {
                'asset_id': asset_id,
                'depreciation_method': 'Straight Line',
                'annual_deduction': round(annual_deduction, 2),
                'remaining_basis': round(asset.cost - total_depreciation, 2),
                'total_depreciation': round(total_depreciation, 2)
            }
        
        return {'error': 'Unknown depreciation method'}
    
    def calculate_business_deductions(self, expenses: List[Dict], entity_type: TaxEntityType) -> Dict:
        """Calculate allowable business deductions"""
        deductible_expenses = {}
        non_deductible = []
        
        # Categorize expenses and apply limits
        for expense in expenses:
            category = expense.get('category', '')
            amount = expense.get('amount', 0)
            description = expense.get('description', '')
            
            if category in ['Meals & Entertainment']:
                # 50% limit on meals
                deductible_amount = amount * 0.5
                deductible_expenses[category] = deductible_expenses.get(category, 0) + deductible_amount
                non_deductible.append({
                    'expense': description,
                    'total': amount,
                    'deductible': deductible_amount,
                    'reason': '50% limit on business meals'
                })
            elif category in ['Travel']:
                # Check for personal vs business travel
                deductible_expenses[category] = deductible_expenses.get(category, 0) + amount
            elif category in ['Home Office']:
                # Simplified home office deduction
                if entity_type == TaxEntityType.SOLE_PROPRIETORSHIP:
                    # Can use simplified method or actual expense method
                    deductible_expenses[category] = deductible_expenses.get(category, 0) + amount
                else:
                    # Different rules for other entities
                    deductible_expenses[category] = deductible_expenses.get(category, 0) + amount
            else:
                deductible_expenses[category] = deductible_expenses.get(category, 0) + amount
        
        total_deductions = sum(deductible_expenses.values())
        
        return {
            'deductible_expenses': deductible_expenses,
            'total_deductions': round(total_deductions, 2),
            'non_deductible_items': non_deductible,
            'entity_type': entity_type.value
        }
    
    def estimated_quarterly_taxes(self, annual_income: float, filing_status: str = 'single',
                                self_employed: bool = False) -> Dict:
        """Calculate estimated quarterly tax payments"""
        # Calculate income tax
        adjusted_income = annual_income
        
        if self_employed:
            # Deduct SE tax deduction
            se_tax_info = self.calculate_self_employment_tax(annual_income)
            adjusted_income -= se_tax_info['deductible_portion']
        
        # Apply standard deduction
        standard_deduction = self.standard_deductions.get(filing_status, 13850)
        taxable_income = max(0, adjusted_income - standard_deduction)
        
        # Calculate federal income tax
        federal_tax_info = self.calculate_federal_income_tax(taxable_income, filing_status)
        federal_tax = federal_tax_info['total_tax']
        
        # Add SE tax if applicable
        se_tax = 0
        if self_employed:
            se_tax = self.calculate_self_employment_tax(annual_income)['total_se_tax']
        
        total_tax = federal_tax + se_tax
        
        # Safe harbor rule (pay 100% of prior year tax or 90% of current year)
        safe_harbor_90 = total_tax * 0.9
        quarterly_payment_90 = safe_harbor_90 / 4
        
        return {
            'annual_income': annual_income,
            'adjusted_income': round(adjusted_income, 2),
            'taxable_income': round(taxable_income, 2),
            'federal_income_tax': round(federal_tax, 2),
            'self_employment_tax': round(se_tax, 2),
            'total_annual_tax': round(total_tax, 2),
            'quarterly_payment_90_percent': round(quarterly_payment_90, 2),
            'due_dates': [
                f"{date.today().year}-01-15",  # Q4 prior year
                f"{date.today().year}-04-15",  # Q1
                f"{date.today().year}-06-15",  # Q2
                f"{date.today().year}-09-15"   # Q3
            ]
        }
    
    def calculate_state_taxes(self, state: str, taxable_income: float, 
                            filing_status: str = 'single') -> Dict:
        """Calculate state income tax (simplified)"""
        # This would need to be expanded for all states
        state_rates = {
            'CA': {
                'single': [(10099, 0.01), (23942, 0.02), (37788, 0.04), 
                          (52455, 0.06), (66295, 0.08), (338639, 0.093),
                          (406364, 0.103), (677278, 0.113), (float('inf'), 0.123)],
                'married_filing_jointly': [(20198, 0.01), (47884, 0.02), (75576, 0.04),
                                         (104910, 0.06), (132590, 0.08), (677278, 0.093),
                                         (812728, 0.103), (1354556, 0.113), (float('inf'), 0.123)]
            },
            'NY': {
                'single': [(8500, 0.04), (11700, 0.045), (13900, 0.0525),
                          (21400, 0.059), (80650, 0.0645), (215400, 0.0665),
                          (1077550, 0.0685), (float('inf'), 0.0882)],
                'married_filing_jointly': [(17150, 0.04), (23600, 0.045), (27900, 0.0525),
                                         (43000, 0.059), (161550, 0.0645), (323200, 0.0665),
                                         (2155350, 0.0685), (float('inf'), 0.0882)]
            },
            'TX': {
                'single': [(float('inf'), 0.0)],  # No state income tax
                'married_filing_jointly': [(float('inf'), 0.0)]
            },
            'FL': {
                'single': [(float('inf'), 0.0)],  # No state income tax
                'married_filing_jointly': [(float('inf'), 0.0)]
            }
        }
        
        if state not in state_rates:
            return {'error': f'State {state} tax rates not available'}
        
        brackets = state_rates[state].get(filing_status, state_rates[state]['single'])
        
        tax_owed = 0
        remaining_income = taxable_income
        prev_threshold = 0
        
        for threshold, rate in brackets:
            if remaining_income <= 0:
                break
            
            taxable_at_rate = min(remaining_income, threshold - prev_threshold)
            tax_owed += taxable_at_rate * rate
            remaining_income -= taxable_at_rate
            prev_threshold = threshold
        
        return {
            'state': state,
            'taxable_income': taxable_income,
            'state_tax': round(tax_owed, 2),
            'effective_rate': round(tax_owed / taxable_income * 100, 2) if taxable_income > 0 else 0
        }
    
    def calculate_payroll_taxes(self, wages: float, year: int = 2024) -> Dict:
        """Calculate payroll taxes (employer and employee portions)"""
        # 2024 rates and limits
        ss_wage_base = 160200
        ss_rate = 0.062  # 6.2%
        medicare_rate = 0.0145  # 1.45%
        additional_medicare_rate = 0.009  # 0.9%
        additional_medicare_threshold = 200000
        futa_wage_base = 7000
        futa_rate = 0.006  # 0.6%
        
        # Employee taxes
        ss_employee = min(wages, ss_wage_base) * ss_rate
        medicare_employee = wages * medicare_rate
        additional_medicare_employee = max(0, wages - additional_medicare_threshold) * additional_medicare_rate
        
        total_employee = ss_employee + medicare_employee + additional_medicare_employee
        
        # Employer taxes (matches employee for SS and Medicare, plus FUTA)
        ss_employer = ss_employee  # Same as employee
        medicare_employer = medicare_employee  # Same as employee
        futa_employer = min(wages, futa_wage_base) * futa_rate
        
        total_employer = ss_employer + medicare_employer + futa_employer
        
        return {
            'wages': wages,
            'employee_taxes': {
                'social_security': round(ss_employee, 2),
                'medicare': round(medicare_employee, 2),
                'additional_medicare': round(additional_medicare_employee, 2),
                'total': round(total_employee, 2)
            },
            'employer_taxes': {
                'social_security': round(ss_employer, 2),
                'medicare': round(medicare_employer, 2),
                'futa': round(futa_employer, 2),
                'total': round(total_employer, 2)
            },
            'total_payroll_cost': round(wages + total_employer, 2)
        }
    
    def business_tax_summary(self, entity_id: str, tax_year: int, 
                           financial_data: Dict) -> Dict:
        """Generate comprehensive business tax summary"""
        entity = self.entities.get(entity_id)
        if not entity:
            return {'error': 'Entity not found'}
        
        revenue = financial_data.get('revenue', 0)
        expenses = financial_data.get('expenses', 0)
        
        # Calculate depreciation
        total_depreciation = 0
        asset_depreciation = []
        
        for asset_id, asset in self.assets.items():
            if asset.placed_in_service_date.year <= tax_year:
                dep_info = self.calculate_depreciation_deduction(asset_id, tax_year)
                if 'annual_deduction' in dep_info:
                    total_depreciation += dep_info['annual_deduction']
                    asset_depreciation.append(dep_info)
        
        # Net income calculation
        net_income = revenue - expenses - total_depreciation
        
        # Tax calculations based on entity type
        tax_calculation = {}
        
        if entity.entity_type == TaxEntityType.SOLE_PROPRIETORSHIP:
            # Schedule C income
            se_tax_info = self.calculate_self_employment_tax(net_income)
            adjusted_income = net_income - se_tax_info['deductible_portion']
            income_tax_info = self.calculate_federal_income_tax(adjusted_income)
            
            tax_calculation = {
                'schedule_c_income': round(net_income, 2),
                'se_tax': se_tax_info,
                'income_tax': income_tax_info,
                'total_federal_tax': round(se_tax_info['total_se_tax'] + income_tax_info['total_tax'], 2)
            }
            
        elif entity.entity_type == TaxEntityType.C_CORP:
            # Corporate tax rate (21% flat rate for 2024)
            corporate_tax = net_income * 0.21
            
            tax_calculation = {
                'taxable_income': round(net_income, 2),
                'corporate_tax_rate': '21%',
                'corporate_tax': round(corporate_tax, 2),
                'after_tax_income': round(net_income - corporate_tax, 2)
            }
        
        elif entity.entity_type == TaxEntityType.S_CORP:
            # Pass-through entity
            tax_calculation = {
                'pass_through_income': round(net_income, 2),
                'note': 'Income passes through to owners - no entity-level tax'
            }
        
        return {
            'entity': entity.name,
            'entity_type': entity.entity_type.value,
            'tax_year': tax_year,
            'financial_summary': {
                'revenue': revenue,
                'expenses': expenses,
                'depreciation': round(total_depreciation, 2),
                'net_income': round(net_income, 2)
            },
            'depreciation_detail': asset_depreciation,
            'tax_calculation': tax_calculation
        }
    
    def tax_planning_strategies(self, current_income: float, projected_income: float,
                              entity_type: TaxEntityType) -> List[Dict]:
        """Suggest tax planning strategies"""
        strategies = []
        
        current_bracket = self._get_tax_bracket(current_income)
        projected_bracket = self._get_tax_bracket(projected_income)
        
        if projected_bracket > current_bracket:
            strategies.append({
                'strategy': 'Accelerate Deductions',
                'description': 'Consider accelerating deductible expenses into current year',
                'potential_savings': (projected_bracket - current_bracket) * min(10000, projected_income - current_income),
                'implementation': 'Prepay expenses, equipment purchases, retirement contributions'
            })
        
        if entity_type == TaxEntityType.SOLE_PROPRIETORSHIP and current_income > 50000:
            strategies.append({
                'strategy': 'Consider Entity Election',
                'description': 'Evaluate S-Corp election to reduce self-employment taxes',
                'potential_savings': current_income * 0.153 * 0.4,  # Rough estimate
                'implementation': 'File Form 2553, establish reasonable salary'
            })
        
        if current_income > 100000:
            strategies.append({
                'strategy': 'Equipment Purchases',
                'description': 'Consider Section 179 or bonus depreciation for equipment',
                'potential_savings': current_bracket * 50000,  # Assuming $50k equipment purchase
                'implementation': 'Purchase and place in service before year-end'
            })
        
        return strategies
    
    def _get_tax_bracket(self, income: float) -> float:
        """Get marginal tax rate for income level"""
        for threshold, rate in self.federal_brackets_single:
            if income <= threshold:
                return rate
        return 0.37  # Top bracket
    
    def generate_tax_projection(self, entity_id: str, scenarios: Dict[str, Dict]) -> pd.DataFrame:
        """Generate tax projections under different scenarios"""
        entity = self.entities.get(entity_id)
        if not entity:
            return pd.DataFrame()
        
        projections = []
        
        for scenario_name, scenario_data in scenarios.items():
            revenue = scenario_data.get('revenue', 0)
            expenses = scenario_data.get('expenses', 0)
            net_income = revenue - expenses
            
            if entity.entity_type == TaxEntityType.SOLE_PROPRIETORSHIP:
                se_tax = self.calculate_self_employment_tax(net_income)['total_se_tax']
                income_tax = self.calculate_federal_income_tax(net_income - se_tax * 0.5)['total_tax']
                total_tax = se_tax + income_tax
            elif entity.entity_type == TaxEntityType.C_CORP:
                total_tax = net_income * 0.21
            else:
                total_tax = 0  # Pass-through entities
            
            projections.append({
                'Scenario': scenario_name,
                'Revenue': revenue,
                'Expenses': expenses,
                'Net Income': round(net_income, 2),
                'Total Tax': round(total_tax, 2),
                'After-tax Income': round(net_income - total_tax, 2),
                'Effective Tax Rate': round(total_tax / net_income * 100, 2) if net_income > 0 else 0
            })
        
        return pd.DataFrame(projections)

if __name__ == "__main__":
    # Test tax calculations
    calc = TaxCalculator()
    
    # Test federal income tax
    tax_info = calc.calculate_federal_income_tax(75000, 'single')
    print("Federal Income Tax:")
    print(json.dumps(tax_info, indent=2))
    
    # Test SE tax
    se_info = calc.calculate_self_employment_tax(50000)
    print("\nSelf-Employment Tax:")
    print(json.dumps(se_info, indent=2))