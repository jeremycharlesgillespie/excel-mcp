import pandas as pd
import numpy as np
from datetime import datetime, date
from typing import Dict, List, Optional, Any, Union
from dataclasses import dataclass
from enum import Enum
import re
from decimal import Decimal, InvalidOperation

class ValidationRule(Enum):
    REQUIRED = "required"
    NUMERIC = "numeric"
    POSITIVE = "positive"
    NEGATIVE = "negative"
    PERCENTAGE = "percentage"
    CURRENCY = "currency"
    DATE = "date"
    EMAIL = "email"
    PHONE = "phone"
    TAX_ID = "tax_id"
    RANGE = "range"
    LENGTH = "length"
    REGEX = "regex"
    CUSTOM = "custom"

@dataclass
class ValidationResult:
    is_valid: bool
    errors: List[str]
    warnings: List[str]
    cleaned_value: Any = None

@dataclass
class ValidationRuleConfig:
    rule_type: ValidationRule
    parameters: Dict[str, Any] = None
    error_message: Optional[str] = None
    warning_message: Optional[str] = None

class FinancialDataValidator:
    """Comprehensive validation for financial data inputs"""
    
    def __init__(self):
        self.currency_symbols = ['$', '€', '£', '¥', '₹']
        self.phone_patterns = [
            r'^\+?1?[-.\s]?\(?([0-9]{3})\)?[-.\s]?([0-9]{3})[-.\s]?([0-9]{4})$',  # US
            r'^\+?([0-9]{1,4})[-.\s]?([0-9]{3,4})[-.\s]?([0-9]{3,4})[-.\s]?([0-9]{3,4})$'  # International
        ]
    
    def validate_field(self, value: Any, rules: List[ValidationRuleConfig]) -> ValidationResult:
        """Validate a single field against multiple rules"""
        errors = []
        warnings = []
        cleaned_value = value
        
        for rule_config in rules:
            result = self._apply_rule(value, rule_config)
            errors.extend(result.errors)
            warnings.extend(result.warnings)
            if result.cleaned_value is not None:
                cleaned_value = result.cleaned_value
        
        return ValidationResult(
            is_valid=len(errors) == 0,
            errors=errors,
            warnings=warnings,
            cleaned_value=cleaned_value
        )
    
    def _apply_rule(self, value: Any, rule_config: ValidationRuleConfig) -> ValidationResult:
        """Apply a single validation rule"""
        rule = rule_config.rule_type
        params = rule_config.parameters or {}
        
        if rule == ValidationRule.REQUIRED:
            if value is None or value == '' or (isinstance(value, str) and value.strip() == ''):
                return ValidationResult(False, [rule_config.error_message or "Field is required"], [])
        
        # Skip other validations if value is empty and not required
        if value is None or value == '':
            return ValidationResult(True, [], [])
        
        if rule == ValidationRule.NUMERIC:
            try:
                if isinstance(value, str):
                    # Clean currency symbols and commas
                    cleaned = re.sub(r'[$,€£¥₹\s]', '', value.strip())
                    float_val = float(cleaned)
                    return ValidationResult(True, [], [], float_val)
                elif isinstance(value, (int, float)):
                    return ValidationResult(True, [], [], float(value))
                else:
                    return ValidationResult(False, [rule_config.error_message or "Value must be numeric"], [])
            except ValueError:
                return ValidationResult(False, [rule_config.error_message or "Invalid numeric value"], [])
        
        elif rule == ValidationRule.POSITIVE:
            numeric_result = self._apply_rule(value, ValidationRuleConfig(ValidationRule.NUMERIC))
            if not numeric_result.is_valid:
                return numeric_result
            if numeric_result.cleaned_value <= 0:
                return ValidationResult(False, [rule_config.error_message or "Value must be positive"], [])
        
        elif rule == ValidationRule.NEGATIVE:
            numeric_result = self._apply_rule(value, ValidationRuleConfig(ValidationRule.NEGATIVE))
            if not numeric_result.is_valid:
                return numeric_result
            if numeric_result.cleaned_value >= 0:
                return ValidationResult(False, [rule_config.error_message or "Value must be negative"], [])
        
        elif rule == ValidationRule.PERCENTAGE:
            numeric_result = self._apply_rule(value, ValidationRuleConfig(ValidationRule.NUMERIC))
            if not numeric_result.is_valid:
                return numeric_result
            
            val = numeric_result.cleaned_value
            if isinstance(value, str) and '%' in value:
                val = val / 100  # Convert percentage to decimal
            
            if not (0 <= val <= 1):
                return ValidationResult(False, [], [rule_config.warning_message or "Percentage should be between 0% and 100%"])
            
            return ValidationResult(True, [], [], val)
        
        elif rule == ValidationRule.CURRENCY:
            if isinstance(value, str):
                # Extract currency amount
                currency_pattern = r'[^\d.,]*([\d,]+\.?\d*)'
                match = re.search(currency_pattern, value)
                if match:
                    amount_str = match.group(1).replace(',', '')
                    try:
                        amount = float(amount_str)
                        return ValidationResult(True, [], [], amount)
                    except ValueError:
                        pass
                return ValidationResult(False, [rule_config.error_message or "Invalid currency format"], [])
        
        elif rule == ValidationRule.DATE:
            if isinstance(value, date):
                return ValidationResult(True, [], [], value)
            elif isinstance(value, str):
                date_formats = ['%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y', '%Y-%m-%d %H:%M:%S']
                for fmt in date_formats:
                    try:
                        parsed_date = datetime.strptime(value, fmt).date()
                        return ValidationResult(True, [], [], parsed_date)
                    except ValueError:
                        continue
                return ValidationResult(False, [rule_config.error_message or "Invalid date format"], [])
        
        elif rule == ValidationRule.EMAIL:
            email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
            if isinstance(value, str) and re.match(email_pattern, value):
                return ValidationResult(True, [], [], value.lower())
            else:
                return ValidationResult(False, [rule_config.error_message or "Invalid email format"], [])
        
        elif rule == ValidationRule.PHONE:
            if isinstance(value, str):
                for pattern in self.phone_patterns:
                    if re.match(pattern, value):
                        # Clean phone number
                        cleaned = re.sub(r'[^\d+]', '', value)
                        return ValidationResult(True, [], [], cleaned)
                return ValidationResult(False, [rule_config.error_message or "Invalid phone format"], [])
        
        elif rule == ValidationRule.TAX_ID:
            if isinstance(value, str):
                # EIN format: XX-XXXXXXX
                ein_pattern = r'^\d{2}-\d{7}$'
                # SSN format: XXX-XX-XXXX
                ssn_pattern = r'^\d{3}-\d{2}-\d{4}$'
                
                cleaned = re.sub(r'[^\d-]', '', value)
                
                if re.match(ein_pattern, cleaned) or re.match(ssn_pattern, cleaned):
                    return ValidationResult(True, [], [], cleaned)
                else:
                    return ValidationResult(False, [rule_config.error_message or "Invalid Tax ID format"], [])
        
        elif rule == ValidationRule.RANGE:
            numeric_result = self._apply_rule(value, ValidationRuleConfig(ValidationRule.NUMERIC))
            if not numeric_result.is_valid:
                return numeric_result
            
            min_val = params.get('min')
            max_val = params.get('max')
            val = numeric_result.cleaned_value
            
            if min_val is not None and val < min_val:
                return ValidationResult(False, [f"Value must be at least {min_val}"], [])
            if max_val is not None and val > max_val:
                return ValidationResult(False, [f"Value must not exceed {max_val}"], [])
        
        elif rule == ValidationRule.LENGTH:
            if isinstance(value, str):
                min_len = params.get('min', 0)
                max_len = params.get('max', float('inf'))
                
                if len(value) < min_len:
                    return ValidationResult(False, [f"Minimum length is {min_len}"], [])
                if len(value) > max_len:
                    return ValidationResult(False, [f"Maximum length is {max_len}"], [])
        
        elif rule == ValidationRule.REGEX:
            pattern = params.get('pattern')
            if pattern and isinstance(value, str):
                if not re.match(pattern, value):
                    return ValidationResult(False, [rule_config.error_message or "Value doesn't match required pattern"], [])
        
        return ValidationResult(True, [], [], cleaned_value)
    
    def validate_financial_statement(self, statement_data: Dict[str, Any]) -> Dict[str, ValidationResult]:
        """Validate an entire financial statement"""
        validation_rules = {
            'revenue': [
                ValidationRuleConfig(ValidationRule.REQUIRED),
                ValidationRuleConfig(ValidationRule.POSITIVE)
            ],
            'expenses': [
                ValidationRuleConfig(ValidationRule.REQUIRED),
                ValidationRuleConfig(ValidationRule.POSITIVE)
            ],
            'assets': [
                ValidationRuleConfig(ValidationRule.POSITIVE)
            ],
            'liabilities': [
                ValidationRuleConfig(ValidationRule.POSITIVE)
            ],
            'equity': [
                ValidationRuleConfig(ValidationRule.NUMERIC)
            ]
        }
        
        results = {}
        for field, rules in validation_rules.items():
            if field in statement_data:
                results[field] = self.validate_field(statement_data[field], rules)
        
        # Cross-field validations
        if 'assets' in results and 'liabilities' in results and 'equity' in results:
            assets_val = results['assets'].cleaned_value or 0
            liabilities_val = results['liabilities'].cleaned_value or 0
            equity_val = results['equity'].cleaned_value or 0
            
            if abs(assets_val - (liabilities_val + equity_val)) > 0.01:
                results['balance_check'] = ValidationResult(
                    False, 
                    ["Balance sheet doesn't balance: Assets ≠ Liabilities + Equity"], 
                    []
                )
            else:
                results['balance_check'] = ValidationResult(True, [], [])
        
        return results
    
    def validate_rental_data(self, rental_data: Dict[str, Any]) -> Dict[str, ValidationResult]:
        """Validate rental property data"""
        validation_rules = {
            'monthly_rent': [
                ValidationRuleConfig(ValidationRule.REQUIRED),
                ValidationRuleConfig(ValidationRule.CURRENCY),
                ValidationRuleConfig(ValidationRule.POSITIVE)
            ],
            'security_deposit': [
                ValidationRuleConfig(ValidationRule.CURRENCY),
                ValidationRuleConfig(ValidationRule.POSITIVE)
            ],
            'lease_start_date': [
                ValidationRuleConfig(ValidationRule.REQUIRED),
                ValidationRuleConfig(ValidationRule.DATE)
            ],
            'lease_end_date': [
                ValidationRuleConfig(ValidationRule.REQUIRED),
                ValidationRuleConfig(ValidationRule.DATE)
            ],
            'square_feet': [
                ValidationRuleConfig(ValidationRule.POSITIVE),
                ValidationRuleConfig(ValidationRule.RANGE, {'min': 100, 'max': 50000})
            ],
            'bedrooms': [
                ValidationRuleConfig(ValidationRule.RANGE, {'min': 0, 'max': 20})
            ],
            'bathrooms': [
                ValidationRuleConfig(ValidationRule.RANGE, {'min': 0, 'max': 20})
            ]
        }
        
        results = {}
        for field, rules in validation_rules.items():
            if field in rental_data:
                results[field] = self.validate_field(rental_data[field], rules)
        
        # Date range validation
        if ('lease_start_date' in results and 'lease_end_date' in results and 
            results['lease_start_date'].is_valid and results['lease_end_date'].is_valid):
            
            start_date = results['lease_start_date'].cleaned_value
            end_date = results['lease_end_date'].cleaned_value
            
            if start_date >= end_date:
                results['date_range'] = ValidationResult(
                    False, 
                    ["Lease start date must be before end date"], 
                    []
                )
            else:
                # Check for reasonable lease term
                lease_days = (end_date - start_date).days
                if lease_days < 30:
                    results['date_range'] = ValidationResult(
                        True, 
                        [], 
                        ["Very short lease term - please verify"]
                    )
                elif lease_days > 365 * 5:
                    results['date_range'] = ValidationResult(
                        True, 
                        [], 
                        ["Very long lease term - please verify"]
                    )
                else:
                    results['date_range'] = ValidationResult(True, [], [])
        
        return results
    
    def validate_expense_data(self, expense_data: Dict[str, Any]) -> Dict[str, ValidationResult]:
        """Validate expense entry data"""
        validation_rules = {
            'amount': [
                ValidationRuleConfig(ValidationRule.REQUIRED),
                ValidationRuleConfig(ValidationRule.CURRENCY),
                ValidationRuleConfig(ValidationRule.POSITIVE)
            ],
            'date': [
                ValidationRuleConfig(ValidationRule.REQUIRED),
                ValidationRuleConfig(ValidationRule.DATE)
            ],
            'vendor_id': [
                ValidationRuleConfig(ValidationRule.REQUIRED),
                ValidationRuleConfig(ValidationRule.LENGTH, {'min': 1, 'max': 50})
            ],
            'category': [
                ValidationRuleConfig(ValidationRule.REQUIRED)
            ],
            'invoice_number': [
                ValidationRuleConfig(ValidationRule.LENGTH, {'max': 50})
            ],
            'description': [
                ValidationRuleConfig(ValidationRule.LENGTH, {'max': 500})
            ]
        }
        
        # Validate against known expense categories
        valid_categories = [
            "Rent/Lease", "Utilities", "Salaries & Wages", "Employee Benefits", 
            "Insurance", "Marketing & Advertising", "Office Supplies", 
            "Maintenance & Repairs", "Professional Fees", "Travel & Entertainment",
            "Raw Materials", "Inventory Purchases", "Freight & Shipping",
            "Equipment", "Property", "Vehicles", "Software",
            "Interest Expense", "Bank Fees", "Taxes",
            "Depreciation", "Amortization", "Other"
        ]
        
        results = {}
        for field, rules in validation_rules.items():
            if field in expense_data:
                results[field] = self.validate_field(expense_data[field], rules)
        
        # Category validation
        if 'category' in expense_data:
            if expense_data['category'] not in valid_categories:
                results['category'] = ValidationResult(
                    False,
                    [f"Invalid category. Must be one of: {', '.join(valid_categories)}"],
                    []
                )
        
        # Amount reasonableness check
        if 'amount' in results and results['amount'].is_valid:
            amount = results['amount'].cleaned_value
            if amount > 100000:
                results['amount_check'] = ValidationResult(
                    True,
                    [],
                    ["Large expense amount - please verify accuracy"]
                )
        
        return results
    
    def validate_cash_flow_data(self, cash_flow_data: Dict[str, Any]) -> Dict[str, ValidationResult]:
        """Validate cash flow entry data"""
        validation_rules = {
            'amount': [
                ValidationRuleConfig(ValidationRule.REQUIRED),
                ValidationRuleConfig(ValidationRule.CURRENCY),
                ValidationRuleConfig(ValidationRule.POSITIVE)
            ],
            'date': [
                ValidationRuleConfig(ValidationRule.REQUIRED),
                ValidationRuleConfig(ValidationRule.DATE)
            ],
            'flow_type': [
                ValidationRuleConfig(ValidationRule.REQUIRED)
            ],
            'direction': [
                ValidationRuleConfig(ValidationRule.REQUIRED)
            ],
            'description': [
                ValidationRuleConfig(ValidationRule.REQUIRED),
                ValidationRuleConfig(ValidationRule.LENGTH, {'min': 5, 'max': 200})
            ]
        }
        
        valid_flow_types = ['Operating', 'Investing', 'Financing']
        valid_directions = ['Inflow', 'Outflow']
        
        results = {}
        for field, rules in validation_rules.items():
            if field in cash_flow_data:
                results[field] = self.validate_field(cash_flow_data[field], rules)
        
        # Enum validations
        if 'flow_type' in cash_flow_data:
            if cash_flow_data['flow_type'] not in valid_flow_types:
                results['flow_type'] = ValidationResult(
                    False,
                    [f"Flow type must be one of: {', '.join(valid_flow_types)}"],
                    []
                )
        
        if 'direction' in cash_flow_data:
            if cash_flow_data['direction'] not in valid_directions:
                results['direction'] = ValidationResult(
                    False,
                    [f"Direction must be one of: {', '.join(valid_directions)}"],
                    []
                )
        
        return results
    
    def validate_excel_data(self, worksheet_data: List[List[Any]], 
                          expected_columns: List[str],
                          column_rules: Dict[str, List[ValidationRuleConfig]]) -> Dict[str, Any]:
        """Validate Excel worksheet data"""
        if not worksheet_data:
            return {'error': 'No data provided'}
        
        # Check headers
        headers = worksheet_data[0] if worksheet_data else []
        header_issues = []
        
        for expected_col in expected_columns:
            if expected_col not in headers:
                header_issues.append(f"Missing required column: {expected_col}")
        
        if header_issues:
            return {'header_errors': header_issues}
        
        # Validate data rows
        validation_results = []
        column_indices = {col: headers.index(col) for col in expected_columns if col in headers}
        
        for row_idx, row in enumerate(worksheet_data[1:], start=2):  # Skip header row
            row_results = {'row': row_idx, 'errors': [], 'warnings': []}
            
            for col_name, col_idx in column_indices.items():
                if col_idx < len(row):
                    cell_value = row[col_idx]
                    rules = column_rules.get(col_name, [])
                    
                    if rules:
                        validation = self.validate_field(cell_value, rules)
                        if not validation.is_valid:
                            row_results['errors'].extend([f"{col_name}: {err}" for err in validation.errors])
                        if validation.warnings:
                            row_results['warnings'].extend([f"{col_name}: {warn}" for warn in validation.warnings])
            
            validation_results.append(row_results)
        
        # Summary
        total_errors = sum(len(row['errors']) for row in validation_results)
        total_warnings = sum(len(row['warnings']) for row in validation_results)
        error_rows = [row['row'] for row in validation_results if row['errors']]
        
        return {
            'total_rows': len(validation_results),
            'total_errors': total_errors,
            'total_warnings': total_warnings,
            'error_rows': error_rows,
            'validation_details': validation_results,
            'is_valid': total_errors == 0
        }
    
    def clean_financial_dataset(self, df: pd.DataFrame, 
                               column_rules: Dict[str, List[ValidationRuleConfig]]) -> Tuple[pd.DataFrame, Dict]:
        """Clean and validate an entire financial dataset"""
        cleaned_df = df.copy()
        cleaning_report = {
            'original_rows': len(df),
            'columns_processed': [],
            'cleaning_actions': [],
            'errors': [],
            'warnings': []
        }
        
        for column, rules in column_rules.items():
            if column not in df.columns:
                cleaning_report['errors'].append(f"Column '{column}' not found in dataset")
                continue
            
            cleaning_report['columns_processed'].append(column)
            
            # Apply cleaning rules to each value in the column
            cleaned_values = []
            column_errors = []
            column_warnings = []
            
            for idx, value in enumerate(df[column]):
                validation = self.validate_field(value, rules)
                
                if validation.is_valid:
                    cleaned_values.append(validation.cleaned_value)
                else:
                    # Keep original value but log error
                    cleaned_values.append(value)
                    column_errors.append(f"Row {idx + 1}: {', '.join(validation.errors)}")
                
                column_warnings.extend([f"Row {idx + 1}: {warn}" for warn in validation.warnings])
            
            cleaned_df[column] = cleaned_values
            
            if column_errors:
                cleaning_report['errors'].extend([f"{column} - {err}" for err in column_errors])
            if column_warnings:
                cleaning_report['warnings'].extend([f"{column} - {warn}" for warn in column_warnings])
        
        # Remove rows with critical errors if specified
        if cleaning_report['errors']:
            cleaning_report['cleaning_actions'].append(f"Dataset contains {len(cleaning_report['errors'])} validation errors")
        
        cleaning_report['final_rows'] = len(cleaned_df)
        cleaning_report['success'] = len(cleaning_report['errors']) == 0
        
        return cleaned_df, cleaning_report

def validate_loan_parameters(principal: float, annual_rate: float, years: int) -> ValidationResult:
    """Validate loan calculation parameters"""
    validator = FinancialDataValidator()
    errors = []
    warnings = []
    
    # Validate principal
    principal_result = validator.validate_field(principal, [
        ValidationRuleConfig(ValidationRule.REQUIRED),
        ValidationRuleConfig(ValidationRule.POSITIVE)
    ])
    errors.extend(principal_result.errors)
    
    # Validate rate
    if annual_rate <= 0 or annual_rate > 1:
        if annual_rate > 1:
            warnings.append("Interest rate appears to be in percentage form - should be decimal")
        else:
            errors.append("Interest rate must be positive")
    
    if annual_rate > 0.5:
        warnings.append("Very high interest rate - please verify")
    
    # Validate years
    if years <= 0:
        errors.append("Loan term must be positive")
    elif years > 50:
        warnings.append("Very long loan term - please verify")
    
    return ValidationResult(
        is_valid=len(errors) == 0,
        errors=errors,
        warnings=warnings
    )

def validate_npv_parameters(rate: float, cash_flows: List[float]) -> ValidationResult:
    """Validate NPV calculation parameters"""
    validator = FinancialDataValidator()
    errors = []
    warnings = []
    
    # Validate rate
    if not isinstance(rate, (int, float)):
        errors.append("Discount rate must be numeric")
    elif rate < -1 or rate > 1:
        warnings.append("Unusual discount rate - should typically be between 0% and 50%")
    
    # Validate cash flows
    if not cash_flows:
        errors.append("Cash flows cannot be empty")
    elif len(cash_flows) < 2:
        warnings.append("NPV typically requires multiple periods")
    
    # Check for all-zero cash flows
    if all(cf == 0 for cf in cash_flows):
        errors.append("All cash flows cannot be zero")
    
    return ValidationResult(
        is_valid=len(errors) == 0,
        errors=errors,
        warnings=warnings
    )

if __name__ == "__main__":
    # Test validation
    validator = FinancialDataValidator()
    
    # Test currency validation
    result = validator.validate_field("$1,234.56", [
        ValidationRuleConfig(ValidationRule.CURRENCY)
    ])
    print("Currency validation:", result)
    
    # Test financial statement validation
    statement = {
        'revenue': 100000,
        'expenses': 80000,
        'assets': 150000,
        'liabilities': 50000,
        'equity': 100000
    }
    
    validation = validator.validate_financial_statement(statement)
    print("\nStatement validation:")
    for field, result in validation.items():
        print(f"{field}: Valid={result.is_valid}, Errors={result.errors}")