import pandas as pd
import numpy as np
from datetime import datetime, date, timedelta
from typing import Dict, List, Optional, Any
import json

def generate_cash_flow_statement(start_date: str, end_date: str) -> Dict[str, Any]:
    """Generate cash flow statement for period"""
    # This would integrate with the CashFlowAnalyzer
    from cash_flow_analysis import CashFlowAnalyzer
    
    analyzer = CashFlowAnalyzer()
    # In practice, would load data from storage
    
    start_dt = datetime.strptime(start_date, '%Y-%m-%d').date()
    end_dt = datetime.strptime(end_date, '%Y-%m-%d').date()
    
    return analyzer.generate_cash_flow_statement(start_dt, end_dt)

def forecast_cash_flow(months_ahead: int = 12, scenarios: Optional[Dict[str, float]] = None) -> List[Dict]:
    """Forecast cash flows with scenario analysis"""
    from cash_flow_analysis import CashFlowAnalyzer
    
    analyzer = CashFlowAnalyzer()
    result = analyzer.forecast_cash_flow(months_ahead, scenarios)
    
    return result.to_dict('records') if hasattr(result, 'to_dict') else []

def cash_burn_analysis(months_back: int = 6) -> Dict[str, Any]:
    """Analyze cash burn rate and runway"""
    from cash_flow_analysis import CashFlowAnalyzer
    
    analyzer = CashFlowAnalyzer()
    return analyzer.cash_burn_analysis(months_back)

def working_capital_analysis(start_date: str, end_date: str) -> Dict[str, Any]:
    """Analyze working capital changes"""
    from cash_flow_analysis import CashFlowAnalyzer
    
    analyzer = CashFlowAnalyzer()
    start_dt = datetime.strptime(start_date, '%Y-%m-%d').date()
    end_dt = datetime.strptime(end_date, '%Y-%m-%d').date()
    
    return analyzer.working_capital_analysis(start_dt, end_dt)

def cash_flow_at_risk(confidence_level: float = 0.95) -> Dict[str, Any]:
    """Calculate Cash Flow at Risk (CFaR)"""
    from cash_flow_analysis import CashFlowAnalyzer
    
    analyzer = CashFlowAnalyzer()
    return analyzer.cash_flow_at_risk(confidence_level)

def liquidity_analysis() -> Dict[str, Any]:
    """Analyze liquidity position"""
    from cash_flow_analysis import CashFlowAnalyzer
    
    analyzer = CashFlowAnalyzer()
    return analyzer.liquidity_analysis()

if __name__ == "__main__":
    # Test cash flow tools
    print("Cash Flow Statement:")
    stmt = generate_cash_flow_statement('2024-01-01', '2024-12-31')
    print(json.dumps(stmt, indent=2, default=str))