import pandas as pd
import numpy as np
from datetime import datetime, date, timedelta
from dateutil.relativedelta import relativedelta
from typing import Dict, List, Optional, Tuple, Union
from dataclasses import dataclass
from enum import Enum

class LeaseStatus(Enum):
    ACTIVE = "active"
    PENDING = "pending"
    EXPIRED = "expired"
    TERMINATED = "terminated"
    MONTH_TO_MONTH = "month_to_month"

@dataclass
class Tenant:
    tenant_id: str
    name: str
    contact_info: Dict[str, str]
    credit_score: Optional[int] = None
    employment_info: Optional[Dict[str, str]] = None

@dataclass
class Lease:
    lease_id: str
    unit_id: str
    tenant_id: str
    start_date: date
    end_date: date
    monthly_rent: float
    security_deposit: float
    status: LeaseStatus
    rent_escalation_rate: float = 0.0
    rent_escalation_frequency: str = "annually"
    additional_charges: Dict[str, float] = None

@dataclass
class Unit:
    unit_id: str
    property_id: str
    unit_number: str
    square_feet: int
    bedrooms: int
    bathrooms: float
    unit_type: str
    amenities: List[str]
    market_rent: float

@dataclass
class Property:
    property_id: str
    name: str
    address: str
    property_type: str
    total_units: int
    year_built: int
    amenities: List[str]

class RentalPropertyManager:
    """Comprehensive rental property management system"""
    
    def __init__(self):
        self.properties: Dict[str, Property] = {}
        self.units: Dict[str, Unit] = {}
        self.leases: Dict[str, Lease] = {}
        self.tenants: Dict[str, Tenant] = {}
        self.payments: List[Dict] = []
        self.maintenance_requests: List[Dict] = []
    
    def calculate_rent_roll(self, property_id: str, as_of_date: date = None) -> pd.DataFrame:
        """Generate rent roll report for a property"""
        if as_of_date is None:
            as_of_date = date.today()
        
        rent_roll_data = []
        
        for unit_id, unit in self.units.items():
            if unit.property_id != property_id:
                continue
            
            active_lease = None
            for lease in self.leases.values():
                if (lease.unit_id == unit_id and 
                    lease.start_date <= as_of_date <= lease.end_date and
                    lease.status == LeaseStatus.ACTIVE):
                    active_lease = lease
                    break
            
            if active_lease:
                tenant = self.tenants.get(active_lease.tenant_id)
                current_rent = self.calculate_current_rent(active_lease, as_of_date)
                
                rent_roll_data.append({
                    'Unit': unit.unit_number,
                    'Tenant': tenant.name if tenant else 'Unknown',
                    'Lease Start': active_lease.start_date,
                    'Lease End': active_lease.end_date,
                    'Monthly Rent': current_rent,
                    'Security Deposit': active_lease.security_deposit,
                    'Status': 'Occupied',
                    'Days Remaining': (active_lease.end_date - as_of_date).days
                })
            else:
                rent_roll_data.append({
                    'Unit': unit.unit_number,
                    'Tenant': 'VACANT',
                    'Lease Start': None,
                    'Lease End': None,
                    'Monthly Rent': unit.market_rent,
                    'Security Deposit': 0,
                    'Status': 'Vacant',
                    'Days Remaining': None
                })
        
        return pd.DataFrame(rent_roll_data)
    
    def calculate_current_rent(self, lease: Lease, as_of_date: date) -> float:
        """Calculate current rent considering escalations"""
        if lease.rent_escalation_rate == 0:
            return lease.monthly_rent
        
        months_elapsed = (as_of_date.year - lease.start_date.year) * 12 + \
                        (as_of_date.month - lease.start_date.month)
        
        if lease.rent_escalation_frequency == "annually":
            years_elapsed = months_elapsed // 12
            return lease.monthly_rent * ((1 + lease.rent_escalation_rate) ** years_elapsed)
        elif lease.rent_escalation_frequency == "semi-annually":
            periods_elapsed = months_elapsed // 6
            return lease.monthly_rent * ((1 + lease.rent_escalation_rate) ** periods_elapsed)
        
        return lease.monthly_rent
    
    def calculate_vacancy_rate(self, property_id: str, start_date: date, end_date: date) -> Dict[str, float]:
        """Calculate vacancy rate and loss"""
        total_units = len([u for u in self.units.values() if u.property_id == property_id])
        days_in_period = (end_date - start_date).days + 1
        
        total_unit_days = total_units * days_in_period
        vacant_unit_days = 0
        potential_rent = 0
        actual_rent = 0
        
        for unit in self.units.values():
            if unit.property_id != property_id:
                continue
            
            unit_vacant_days = 0
            current_date = start_date
            
            while current_date <= end_date:
                lease_found = False
                for lease in self.leases.values():
                    if (lease.unit_id == unit.unit_id and 
                        lease.start_date <= current_date <= lease.end_date and
                        lease.status == LeaseStatus.ACTIVE):
                        lease_found = True
                        actual_rent += self.calculate_current_rent(lease, current_date) / 30
                        break
                
                if not lease_found:
                    unit_vacant_days += 1
                    
                potential_rent += unit.market_rent / 30
                current_date += timedelta(days=1)
            
            vacant_unit_days += unit_vacant_days
        
        vacancy_rate = (vacant_unit_days / total_unit_days) * 100 if total_unit_days > 0 else 0
        economic_vacancy = ((potential_rent - actual_rent) / potential_rent * 100) if potential_rent > 0 else 0
        
        return {
            'physical_vacancy_rate': round(vacancy_rate, 2),
            'economic_vacancy_rate': round(economic_vacancy, 2),
            'vacant_unit_days': vacant_unit_days,
            'total_unit_days': total_unit_days,
            'potential_rent': round(potential_rent, 2),
            'actual_rent': round(actual_rent, 2),
            'vacancy_loss': round(potential_rent - actual_rent, 2)
        }
    
    def generate_lease_expiration_report(self, months_ahead: int = 3) -> pd.DataFrame:
        """Generate report of upcoming lease expirations"""
        cutoff_date = date.today() + relativedelta(months=months_ahead)
        expiring_leases = []
        
        for lease in self.leases.values():
            if (lease.status == LeaseStatus.ACTIVE and 
                date.today() <= lease.end_date <= cutoff_date):
                
                unit = self.units.get(lease.unit_id)
                tenant = self.tenants.get(lease.tenant_id)
                
                expiring_leases.append({
                    'Unit': unit.unit_number if unit else 'Unknown',
                    'Tenant': tenant.name if tenant else 'Unknown',
                    'Lease End': lease.end_date,
                    'Days Until Expiry': (lease.end_date - date.today()).days,
                    'Current Rent': self.calculate_current_rent(lease, date.today()),
                    'Market Rent': unit.market_rent if unit else 0,
                    'Variance': ((unit.market_rent - self.calculate_current_rent(lease, date.today())) 
                                / self.calculate_current_rent(lease, date.today()) * 100) if unit else 0
                })
        
        df = pd.DataFrame(expiring_leases)
        if not df.empty:
            df = df.sort_values('Days Until Expiry')
        return df
    
    def calculate_noi(self, property_id: str, year: int) -> Dict[str, float]:
        """Calculate Net Operating Income for a property"""
        start_date = date(year, 1, 1)
        end_date = date(year, 12, 31)
        
        # Revenue calculations
        rental_income = 0
        other_income = 0
        
        for unit in self.units.values():
            if unit.property_id != property_id:
                continue
            
            for month in range(1, 13):
                month_start = date(year, month, 1)
                month_end = date(year, month, 28)  # Simplified
                
                for lease in self.leases.values():
                    if (lease.unit_id == unit.unit_id and 
                        lease.start_date <= month_end and lease.end_date >= month_start and
                        lease.status == LeaseStatus.ACTIVE):
                        
                        rental_income += self.calculate_current_rent(lease, month_start)
                        
                        if lease.additional_charges:
                            other_income += sum(lease.additional_charges.values())
        
        vacancy_data = self.calculate_vacancy_rate(property_id, start_date, end_date)
        vacancy_loss = vacancy_data['vacancy_loss']
        
        # Operating expenses (simplified - would be tracked separately)
        operating_expenses = {
            'Property Management': rental_income * 0.08,
            'Maintenance & Repairs': rental_income * 0.10,
            'Insurance': rental_income * 0.05,
            'Property Taxes': rental_income * 0.15,
            'Utilities': rental_income * 0.03,
            'Administrative': rental_income * 0.02,
            'Marketing': vacancy_loss * 0.20
        }
        
        gross_rental_income = rental_income
        effective_rental_income = rental_income - vacancy_loss
        total_revenue = effective_rental_income + other_income
        total_expenses = sum(operating_expenses.values())
        noi = total_revenue - total_expenses
        
        return {
            'Gross Rental Income': round(gross_rental_income, 2),
            'Vacancy Loss': round(vacancy_loss, 2),
            'Effective Rental Income': round(effective_rental_income, 2),
            'Other Income': round(other_income, 2),
            'Total Revenue': round(total_revenue, 2),
            'Operating Expenses': operating_expenses,
            'Total Operating Expenses': round(total_expenses, 2),
            'Net Operating Income': round(noi, 2),
            'Operating Expense Ratio': round(total_expenses / total_revenue * 100, 2) if total_revenue > 0 else 0,
            'NOI Margin': round(noi / total_revenue * 100, 2) if total_revenue > 0 else 0
        }
    
    def calculate_cap_rate(self, property_id: str, property_value: float, year: int) -> float:
        """Calculate capitalization rate"""
        noi_data = self.calculate_noi(property_id, year)
        noi = noi_data['Net Operating Income']
        return (noi / property_value * 100) if property_value > 0 else 0
    
    def project_cash_flow(self, property_id: str, years: int, 
                         initial_investment: float, loan_amount: float = 0,
                         loan_rate: float = 0, loan_term_years: int = 0) -> pd.DataFrame:
        """Project property cash flows"""
        projections = []
        current_year = date.today().year
        
        # Calculate loan payment if applicable
        if loan_amount > 0 and loan_rate > 0 and loan_term_years > 0:
            monthly_rate = loan_rate / 12
            n_payments = loan_term_years * 12
            monthly_payment = loan_amount * (monthly_rate * (1 + monthly_rate)**n_payments) / \
                            ((1 + monthly_rate)**n_payments - 1)
            annual_debt_service = monthly_payment * 12
        else:
            annual_debt_service = 0
        
        for year_offset in range(years):
            year = current_year + year_offset
            noi_data = self.calculate_noi(property_id, year)
            
            # Apply growth assumptions
            if year_offset > 0:
                noi = noi_data['Net Operating Income'] * (1.03 ** year_offset)  # 3% growth
            else:
                noi = noi_data['Net Operating Income']
            
            # Calculate cash flow
            before_tax_cash_flow = noi - annual_debt_service
            
            # Simplified tax calculation
            depreciation = initial_investment / 27.5  # Residential depreciation
            taxable_income = noi - annual_debt_service - depreciation
            taxes = max(0, taxable_income * 0.25)  # 25% tax rate
            
            after_tax_cash_flow = before_tax_cash_flow - taxes
            
            projections.append({
                'Year': year,
                'NOI': round(noi, 2),
                'Debt Service': round(annual_debt_service, 2),
                'Before Tax Cash Flow': round(before_tax_cash_flow, 2),
                'Depreciation': round(depreciation, 2),
                'Taxes': round(taxes, 2),
                'After Tax Cash Flow': round(after_tax_cash_flow, 2)
            })
        
        return pd.DataFrame(projections)
    
    def analyze_rent_comparables(self, unit_id: str, radius_properties: List[str]) -> Dict:
        """Analyze comparable rents for market analysis"""
        target_unit = self.units.get(unit_id)
        if not target_unit:
            return {}
        
        comparables = []
        
        for prop_id in radius_properties:
            for unit in self.units.values():
                if unit.property_id != prop_id:
                    continue
                
                # Find similar units (same bedroom count, similar size)
                if (unit.bedrooms == target_unit.bedrooms and
                    abs(unit.square_feet - target_unit.square_feet) / target_unit.square_feet < 0.15):
                    
                    # Get current rent if leased
                    current_rent = unit.market_rent
                    for lease in self.leases.values():
                        if (lease.unit_id == unit.unit_id and 
                            lease.status == LeaseStatus.ACTIVE and
                            lease.start_date <= date.today() <= lease.end_date):
                            current_rent = self.calculate_current_rent(lease, date.today())
                            break
                    
                    comparables.append({
                        'property': self.properties.get(prop_id).name if prop_id in self.properties else 'Unknown',
                        'unit': unit.unit_number,
                        'sq_ft': unit.square_feet,
                        'rent': current_rent,
                        'rent_per_sq_ft': current_rent / unit.square_feet
                    })
        
        if not comparables:
            return {'error': 'No comparable units found'}
        
        df = pd.DataFrame(comparables)
        
        return {
            'target_unit': target_unit.unit_number,
            'target_sq_ft': target_unit.square_feet,
            'current_market_rent': target_unit.market_rent,
            'comparables': comparables,
            'average_rent': round(df['rent'].mean(), 2),
            'median_rent': round(df['rent'].median(), 2),
            'average_rent_per_sq_ft': round(df['rent_per_sq_ft'].mean(), 2),
            'suggested_rent': round(df['rent_per_sq_ft'].median() * target_unit.square_feet, 2),
            'rent_variance': round(((df['rent_per_sq_ft'].median() * target_unit.square_feet) - 
                                   target_unit.market_rent) / target_unit.market_rent * 100, 2)
        }

if __name__ == "__main__":
    # Test the rental management system
    rpm = RentalPropertyManager()
    
    # Add sample property
    prop = Property("P001", "Sunset Apartments", "123 Main St", "Multifamily", 10, 2010, ["Pool", "Gym"])
    rpm.properties[prop.property_id] = prop
    
    # Add sample unit
    unit = Unit("U001", "P001", "101", 850, 2, 1.0, "Apartment", ["Balcony"], 1500)
    rpm.units[unit.unit_id] = unit
    
    # Add sample tenant and lease
    tenant = Tenant("T001", "John Doe", {"phone": "555-1234", "email": "john@example.com"})
    rpm.tenants[tenant.tenant_id] = tenant
    
    lease = Lease("L001", "U001", "T001", date(2024, 1, 1), date(2025, 12, 31), 
                  1500, 3000, LeaseStatus.ACTIVE, 0.03)
    rpm.leases[lease.lease_id] = lease
    
    # Test rent roll
    print("Rent Roll:")
    print(rpm.calculate_rent_roll("P001"))
    
    # Test NOI
    print("\nNOI Analysis:")
    print(rpm.calculate_noi("P001", 2024))