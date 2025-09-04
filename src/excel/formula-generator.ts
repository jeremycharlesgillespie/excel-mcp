import { CellValue } from './excel-manager.js';

export interface FormulaInfo {
  formula: string;
  description: string;
  parameters: string[];
  accountingStandard?: string;
  referenceUrl?: string;
  example?: string;
  validation?: string;
}

export class FormulaGenerator {
  
  static generateNPVFormula(rate: number, cashFlowRange: string, initialInvestment?: number): FormulaInfo {
    const formula = initialInvestment 
      ? `=NPV(${rate},${cashFlowRange})-${initialInvestment}`
      : `=NPV(${rate},${cashFlowRange})`;
    
    return {
      formula,
      description: "Net Present Value - calculates present value of future cash flows discounted at specified rate",
      parameters: [`Rate: ${rate}`, `Cash Flows: ${cashFlowRange}`, initialInvestment ? `Initial Investment: ${initialInvestment}` : ''].filter(Boolean),
      accountingStandard: "GAAP/IFRS - Capital Investment Analysis",
      referenceUrl: "https://www.fasb.org/page/PageContent?pageId=/standards/concepts-statements/concept-7.html",
      example: "NPV = CF1/(1+r)^1 + CF2/(1+r)^2 + ... - Initial Investment",
      validation: "Rate should be positive decimal (e.g., 0.1 for 10%)"
    };
  }

  static generateLoanPaymentFormula(principal: number, rate: number, periods: number): FormulaInfo {
    const monthlyRate = rate / 12;
    const formula = `=PMT(${monthlyRate},${periods},-${principal})`;
    
    return {
      formula,
      description: "Monthly loan payment calculation using standard amortization formula",
      parameters: [`Principal: ${principal}`, `Annual Rate: ${rate}`, `Periods: ${periods} months`],
      accountingStandard: "GAAP - ASC 835 Interest",
      referenceUrl: "https://www.fasb.org/page/PageContent?pageId=/asc/835/835-20.html",
      example: "PMT = P * [r(1+r)^n] / [(1+r)^n - 1]",
      validation: "Principal > 0, Rate > 0, Periods > 0"
    };
  }

  static generateDepreciationFormula(method: 'straight-line' | 'declining-balance' | 'sum-of-years', 
                                   cost: number, salvage: number, life: number, year?: number): FormulaInfo {
    let formula: string;
    let description: string;
    let example: string;
    let referenceUrl: string;

    switch (method) {
      case 'straight-line':
        formula = `=(${cost}-${salvage})/${life}`;
        description = "Straight-line depreciation - equal annual depreciation over asset's useful life";
        example = "Annual Depreciation = (Cost - Salvage Value) / Useful Life";
        referenceUrl = "https://www.fasb.org/page/PageContent?pageId=/asc/360/360-10-35.html";
        break;
        
      case 'declining-balance':
        const rate = 2 / life; // Double declining balance
        formula = year ? `=${cost}*${rate}*(1-${rate})^(${year}-1)` : `=${cost}*${rate}`;
        description = "Declining balance depreciation - accelerated depreciation method";
        example = "Annual Depreciation = Book Value * Depreciation Rate";
        referenceUrl = "https://www.fasb.org/page/PageContent?pageId=/asc/360/360-10-35.html";
        break;
        
      case 'sum-of-years':
        const sumOfYears = life * (life + 1) / 2;
        formula = year ? `=(${cost}-${salvage})*((${life}-${year}+1)/${sumOfYears})` : `=(${cost}-${salvage})*${life}/${sumOfYears}`;
        description = "Sum-of-years-digits depreciation - accelerated method based on fraction of remaining life";
        example = "Annual Depreciation = (Cost - Salvage) * (Remaining Life / Sum of Years)";
        referenceUrl = "https://www.fasb.org/page/PageContent?pageId=/asc/360/360-10-35.html";
        break;
        
      default:
        throw new Error(`Unknown depreciation method: ${method}`);
    }

    return {
      formula,
      description,
      parameters: [`Cost: ${cost}`, `Salvage: ${salvage}`, `Life: ${life} years`, year ? `Year: ${year}` : ''].filter(Boolean),
      accountingStandard: "GAAP - ASC 360 Property, Plant, and Equipment",
      referenceUrl,
      example,
      validation: "Cost > Salvage ≥ 0, Life > 0, Year ≤ Life"
    };
  }

  static generateCurrentRatioFormula(currentAssetsRange: string, currentLiabilitiesRange: string): FormulaInfo {
    return {
      formula: `=${currentAssetsRange}/${currentLiabilitiesRange}`,
      description: "Current Ratio - measures company's ability to pay short-term obligations with current assets",
      parameters: [`Current Assets: ${currentAssetsRange}`, `Current Liabilities: ${currentLiabilitiesRange}`],
      accountingStandard: "GAAP - Liquidity Analysis",
      referenceUrl: "https://www.fasb.org/page/PageContent?pageId=/asc/210/210-10.html",
      example: "Current Ratio = Current Assets / Current Liabilities",
      validation: "Ratio > 1.0 indicates good liquidity, > 2.0 may indicate excess cash"
    };
  }

  static generateROAFormula(netIncomeCell: string, totalAssetsCell: string): FormulaInfo {
    return {
      formula: `=${netIncomeCell}/${totalAssetsCell}*100`,
      description: "Return on Assets - measures how efficiently company uses assets to generate profit",
      parameters: [`Net Income: ${netIncomeCell}`, `Total Assets: ${totalAssetsCell}`],
      accountingStandard: "GAAP - Performance Analysis",
      referenceUrl: "https://www.sec.gov/files/aqfs.pdf",
      example: "ROA = Net Income / Total Assets",
      validation: "Higher ROA indicates better asset utilization efficiency"
    };
  }

  static generateROEFormula(netIncomeCell: string, shareholdersEquityCell: string): FormulaInfo {
    return {
      formula: `=${netIncomeCell}/${shareholdersEquityCell}*100`,
      description: "Return on Equity - measures return generated on shareholders' equity investment",
      parameters: [`Net Income: ${netIncomeCell}`, `Shareholders Equity: ${shareholdersEquityCell}`],
      accountingStandard: "GAAP - Performance Analysis",
      referenceUrl: "https://www.sec.gov/files/aqfs.pdf",
      example: "ROE = Net Income / Shareholders' Equity",
      validation: "ROE > Cost of Equity indicates value creation for shareholders"
    };
  }

  static generateDebtToEquityFormula(totalDebtCell: string, totalEquityCell: string): FormulaInfo {
    return {
      formula: `=${totalDebtCell}/${totalEquityCell}`,
      description: "Debt-to-Equity Ratio - measures financial leverage and capital structure",
      parameters: [`Total Debt: ${totalDebtCell}`, `Total Equity: ${totalEquityCell}`],
      accountingStandard: "GAAP - ASC 470 Debt",
      referenceUrl: "https://www.fasb.org/page/PageContent?pageId=/asc/470/470-10.html",
      example: "D/E = Total Debt / Total Equity",
      validation: "Lower ratio indicates less financial risk, industry benchmarks vary"
    };
  }

  static generateCapRateFormula(noiCell: string, propertyValueCell: string): FormulaInfo {
    return {
      formula: `=${noiCell}/${propertyValueCell}*100`,
      description: "Capitalization Rate - measures return on real estate investment based on NOI",
      parameters: [`Net Operating Income: ${noiCell}`, `Property Value: ${propertyValueCell}`],
      accountingStandard: "Real Estate Investment Analysis",
      referenceUrl: "https://www.appraisalinstitute.org/education/",
      example: "Cap Rate = Net Operating Income / Property Value",
      validation: "Typical cap rates range 4-12% depending on property type and location"
    };
  }

  static generateNOIFormula(grossIncomeRange: string, operatingExpensesRange: string, vacancyLossCell?: string): FormulaInfo {
    const formula = vacancyLossCell 
      ? `=SUM(${grossIncomeRange})-${vacancyLossCell}-SUM(${operatingExpensesRange})`
      : `=SUM(${grossIncomeRange})-SUM(${operatingExpensesRange})`;

    return {
      formula,
      description: "Net Operating Income - rental income minus operating expenses, excluding financing costs",
      parameters: [
        `Gross Income: ${grossIncomeRange}`, 
        `Operating Expenses: ${operatingExpensesRange}`,
        vacancyLossCell ? `Vacancy Loss: ${vacancyLossCell}` : ''
      ].filter(Boolean),
      accountingStandard: "Real Estate Investment Analysis - GAAP ASC 360",
      referenceUrl: "https://www.fasb.org/page/PageContent?pageId=/asc/360/360-20.html",
      example: "NOI = Gross Rental Income - Vacancy Loss - Operating Expenses",
      validation: "NOI should exclude financing costs, depreciation, and capital expenditures"
    };
  }

  static generateCashOnCashFormula(annualCashFlowCell: string, totalInvestmentCell: string): FormulaInfo {
    return {
      formula: `=${annualCashFlowCell}/${totalInvestmentCell}*100`,
      description: "Cash-on-Cash Return - measures annual cash flow relative to cash invested",
      parameters: [`Annual Cash Flow: ${annualCashFlowCell}`, `Total Cash Investment: ${totalInvestmentCell}`],
      accountingStandard: "Real Estate Investment Analysis",
      referenceUrl: "https://www.reit.com/investing/measuring-reit-performance",
      example: "Cash-on-Cash = Annual Cash Flow / Total Cash Invested",
      validation: "Measures return on actual cash invested, excludes financing leverage effects"
    };
  }

  static generateDSCRFormula(noiCell: string, debtServiceCell: string): FormulaInfo {
    return {
      formula: `=${noiCell}/${debtServiceCell}`,
      description: "Debt Service Coverage Ratio - measures ability to service debt from operating income",
      parameters: [`Net Operating Income: ${noiCell}`, `Annual Debt Service: ${debtServiceCell}`],
      accountingStandard: "Banking/Lending Standards",
      referenceUrl: "https://www.occ.gov/publications-and-resources/publications/commercial-real-estate/pub-cre-concentration-risk-mgmt.pdf",
      example: "DSCR = Net Operating Income / Annual Debt Service",
      validation: "DSCR > 1.25 typically required by lenders, > 1.5 preferred"
    };
  }

  static generateWorkingCapitalFormula(currentAssetsRange: string, currentLiabilitiesRange: string): FormulaInfo {
    return {
      formula: `=SUM(${currentAssetsRange})-SUM(${currentLiabilitiesRange})`,
      description: "Working Capital - measures short-term liquidity and operational efficiency",
      parameters: [`Current Assets: ${currentAssetsRange}`, `Current Liabilities: ${currentLiabilitiesRange}`],
      accountingStandard: "GAAP - ASC 210 Balance Sheet",
      referenceUrl: "https://www.fasb.org/page/PageContent?pageId=/asc/210/210-10.html",
      example: "Working Capital = Current Assets - Current Liabilities",
      validation: "Positive working capital indicates ability to meet short-term obligations"
    };
  }

  static generateGrossMarginFormula(revenueCell: string, cogsCell: string): FormulaInfo {
    return {
      formula: `=(${revenueCell}-${cogsCell})/${revenueCell}*100`,
      description: "Gross Margin Percentage - measures profitability after direct costs",
      parameters: [`Revenue: ${revenueCell}`, `Cost of Goods Sold: ${cogsCell}`],
      accountingStandard: "GAAP - Revenue Recognition ASC 606",
      referenceUrl: "https://www.fasb.org/page/PageContent?pageId=/asc/606/606-10.html",
      example: "Gross Margin = (Revenue - COGS) / Revenue",
      validation: "Higher margins indicate better pricing power and cost control"
    };
  }

  static generateOperatingMarginFormula(operatingIncomeCell: string, revenueCell: string): FormulaInfo {
    return {
      formula: `=${operatingIncomeCell}/${revenueCell}*100`,
      description: "Operating Margin - measures operational efficiency excluding financing and tax effects",
      parameters: [`Operating Income: ${operatingIncomeCell}`, `Revenue: ${revenueCell}`],
      accountingStandard: "GAAP - Income Statement Presentation",
      referenceUrl: "https://www.fasb.org/page/PageContent?pageId=/asc/220/220-10.html",
      example: "Operating Margin = Operating Income / Revenue",
      validation: "Consistent margins indicate stable operations, declining margins signal inefficiencies"
    };
  }

  static generateInventoryTurnoverFormula(cogsCell: string, avgInventoryCell: string): FormulaInfo {
    return {
      formula: `=${cogsCell}/${avgInventoryCell}`,
      description: "Inventory Turnover - measures how efficiently inventory is converted to sales",
      parameters: [`Cost of Goods Sold: ${cogsCell}`, `Average Inventory: ${avgInventoryCell}`],
      accountingStandard: "GAAP - ASC 330 Inventory",
      referenceUrl: "https://www.fasb.org/page/PageContent?pageId=/asc/330/330-10.html",
      example: "Inventory Turnover = COGS / Average Inventory",
      validation: "Higher turnover indicates efficient inventory management"
    };
  }

  static generateDaysSalesOutstandingFormula(accountsReceivableCell: string, annualSalesCell: string): FormulaInfo {
    return {
      formula: `=${accountsReceivableCell}/${annualSalesCell}*365`,
      description: "Days Sales Outstanding - average collection period for receivables",
      parameters: [`Accounts Receivable: ${accountsReceivableCell}`, `Annual Sales: ${annualSalesCell}`],
      accountingStandard: "GAAP - ASC 310 Receivables",
      referenceUrl: "https://www.fasb.org/page/PageContent?pageId=/asc/310/310-10.html",
      example: "DSO = (Accounts Receivable / Annual Sales) * 365",
      validation: "Lower DSO indicates faster collection, compare to payment terms"
    };
  }

  static generateBreakEvenFormula(fixedCostsCell: string, pricePer: number, variableCostPer: number): FormulaInfo {
    return {
      formula: `=${fixedCostsCell}/(${pricePer}-${variableCostPer})`,
      description: "Break-even Analysis - units needed to cover all costs",
      parameters: [`Fixed Costs: ${fixedCostsCell}`, `Price per Unit: ${pricePer}`, `Variable Cost per Unit: ${variableCostPer}`],
      accountingStandard: "Management Accounting - Cost-Volume-Profit Analysis",
      referenceUrl: "https://www.imanet.org/insights-and-trends/management-accounting/cost-management",
      example: "Break-even Units = Fixed Costs / (Price per Unit - Variable Cost per Unit)",
      validation: "Price per unit must exceed variable cost per unit"
    };
  }

  static generateWACCFormula(equityValueCell: string, debtValueCell: string, costOfEquityCell: string, 
                           costOfDebtCell: string, taxRateCell: string): FormulaInfo {
    const totalValueFormula = `(${equityValueCell}+${debtValueCell})`;
    const formula = `=(${equityValueCell}/${totalValueFormula})*${costOfEquityCell}+(${debtValueCell}/${totalValueFormula})*${costOfDebtCell}*(1-${taxRateCell})`;
    
    return {
      formula,
      description: "Weighted Average Cost of Capital - blended cost of equity and debt financing",
      parameters: [
        `Equity Value: ${equityValueCell}`, 
        `Debt Value: ${debtValueCell}`,
        `Cost of Equity: ${costOfEquityCell}`,
        `Cost of Debt: ${costOfDebtCell}`,
        `Tax Rate: ${taxRateCell}`
      ],
      accountingStandard: "Corporate Finance - Capital Structure Analysis",
      referenceUrl: "https://www.sec.gov/files/aqfs.pdf",
      example: "WACC = (E/V * Cost of Equity) + (D/V * Cost of Debt * (1 - Tax Rate))",
      validation: "Used for DCF analysis and investment evaluation"
    };
  }

  static generateEBITDAFormula(netIncomeCell: string, taxesCell: string, interestCell: string, 
                             depreciationCell: string, amortizationCell: string): FormulaInfo {
    return {
      formula: `=${netIncomeCell}+${taxesCell}+${interestCell}+${depreciationCell}+${amortizationCell}`,
      description: "EBITDA - Earnings before Interest, Taxes, Depreciation, and Amortization",
      parameters: [
        `Net Income: ${netIncomeCell}`,
        `Taxes: ${taxesCell}`, 
        `Interest: ${interestCell}`,
        `Depreciation: ${depreciationCell}`,
        `Amortization: ${amortizationCell}`
      ],
      accountingStandard: "Non-GAAP Measure - SEC Regulation G",
      referenceUrl: "https://www.sec.gov/corpfin/guidance/non-gaap-interp",
      example: "EBITDA = Net Income + Taxes + Interest + Depreciation + Amortization",
      validation: "Non-GAAP measure - must reconcile to GAAP net income per SEC requirements"
    };
  }

  static generateQuickRatioFormula(currentAssetsCell: string, inventoryCell: string, currentLiabilitiesCell: string): FormulaInfo {
    return {
      formula: `=(${currentAssetsCell}-${inventoryCell})/${currentLiabilitiesCell}`,
      description: "Quick Ratio (Acid Test) - measures immediate liquidity excluding inventory",
      parameters: [
        `Current Assets: ${currentAssetsCell}`,
        `Inventory: ${inventoryCell}`,
        `Current Liabilities: ${currentLiabilitiesCell}`
      ],
      accountingStandard: "GAAP - Liquidity Analysis",
      referenceUrl: "https://www.fasb.org/page/PageContent?pageId=/asc/210/210-10.html",
      example: "Quick Ratio = (Current Assets - Inventory) / Current Liabilities",
      validation: "Ratio > 1.0 indicates good short-term liquidity without relying on inventory sales"
    };
  }

  static generatePriceToEarningsFormula(stockPriceCell: string, epsCell: string): FormulaInfo {
    return {
      formula: `=${stockPriceCell}/${epsCell}`,
      description: "Price-to-Earnings Ratio - valuation metric comparing stock price to earnings per share",
      parameters: [`Stock Price: ${stockPriceCell}`, `Earnings per Share: ${epsCell}`],
      accountingStandard: "SEC - Financial Reporting Requirements",
      referenceUrl: "https://www.sec.gov/files/aqfs.pdf",
      example: "P/E Ratio = Stock Price / Earnings per Share",
      validation: "Compare to industry averages and historical P/E ranges"
    };
  }

  static generateFreeeCashFlowFormula(operatingCashFlowCell: string, capexCell: string): FormulaInfo {
    return {
      formula: `=${operatingCashFlowCell}-${capexCell}`,
      description: "Free Cash Flow - cash generated after necessary capital expenditures",
      parameters: [`Operating Cash Flow: ${operatingCashFlowCell}`, `Capital Expenditures: ${capexCell}`],
      accountingStandard: "GAAP - ASC 230 Statement of Cash Flows",
      referenceUrl: "https://www.fasb.org/page/PageContent?pageId=/asc/230/230-10.html",
      example: "Free Cash Flow = Operating Cash Flow - Capital Expenditures",
      validation: "Positive FCF indicates ability to generate cash after growth investments"
    };
  }

  static createFormulaDocumentationSheet(): Array<Array<string | CellValue>> {
    return [
      ['FINANCIAL FORMULA REFERENCE', '', '', '', '', ''],
      ['Generated by Excel Finance MCP', '', '', '', '', ''],
      ['', '', '', '', '', ''],
      ['Formula Name', 'Excel Formula', 'Description', 'Accounting Standard', 'Reference', 'Validation Notes'],
      
      // Investment Analysis
      ['NPV (Net Present Value)', '=NPV(rate,cash_flows)-initial_investment', 
       'Present value of future cash flows discounted at specified rate',
       'GAAP Capital Investment Analysis', 
       'https://www.fasb.org/page/PageContent?pageId=/standards/concepts-statements/concept-7.html',
       'Rate should be WACC or required return'],
       
      ['IRR (Internal Rate of Return)', '=IRR(cash_flows)', 
       'Discount rate that makes NPV equal to zero',
       'GAAP Capital Investment Analysis',
       'https://www.fasb.org/page/PageContent?pageId=/standards/concepts-statements/concept-7.html', 
       'Compare to WACC for investment decisions'],
       
      ['Loan Payment (PMT)', '=PMT(rate/12,periods,-principal)',
       'Monthly payment for loan amortization',
       'GAAP ASC 835 Interest',
       'https://www.fasb.org/page/PageContent?pageId=/asc/835/835-20.html',
       'Use monthly rate (annual rate / 12)'],
       
      // Depreciation
      ['Straight-Line Depreciation', '=(cost-salvage)/life',
       'Equal annual depreciation over useful life',
       'GAAP ASC 360 Property, Plant & Equipment',
       'https://www.fasb.org/page/PageContent?pageId=/asc/360/360-10-35.html',
       'Most common method, required cost > salvage'],
       
      ['Double Declining Balance', '=cost*rate*(1-rate)^(year-1)',
       'Accelerated depreciation method',
       'GAAP ASC 360 Property, Plant & Equipment',
       'https://www.fasb.org/page/PageContent?pageId=/asc/360/360-10-35.html',
       'Rate = 2/useful life, cannot depreciate below salvage'],
       
      // Liquidity Ratios
      ['Current Ratio', '=current_assets/current_liabilities',
       'Short-term liquidity measurement',
       'GAAP ASC 210 Balance Sheet',
       'https://www.fasb.org/page/PageContent?pageId=/asc/210/210-10.html',
       'Ratio > 1.0 good, > 2.0 may indicate excess cash'],
       
      ['Quick Ratio (Acid Test)', '=(current_assets-inventory)/current_liabilities',
       'Immediate liquidity excluding inventory',
       'GAAP ASC 210 Balance Sheet', 
       'https://www.fasb.org/page/PageContent?pageId=/asc/210/210-10.html',
       'More conservative than current ratio'],
       
      ['Cash Ratio', '=cash_and_equivalents/current_liabilities',
       'Most conservative liquidity measure',
       'GAAP ASC 305 Cash and Cash Equivalents',
       'https://www.fasb.org/page/PageContent?pageId=/asc/305/305-10.html',
       'Only most liquid assets considered'],
       
      // Profitability Ratios  
      ['Gross Margin %', '=(revenue-cogs)/revenue*100',
       'Profitability after direct costs',
       'GAAP Revenue Recognition ASC 606',
       'https://www.fasb.org/page/PageContent?pageId=/asc/606/606-10.html',
       'Industry comparison important'],
       
      ['Operating Margin %', '=operating_income/revenue*100', 
       'Operational efficiency measure',
       'GAAP ASC 220 Income Statement',
       'https://www.fasb.org/page/PageContent?pageId=/asc/220/220-10.html',
       'Excludes financing and tax effects'],
       
      ['Net Margin %', '=net_income/revenue*100',
       'Bottom-line profitability measure', 
       'GAAP ASC 220 Income Statement',
       'https://www.fasb.org/page/PageContent?pageId=/asc/220/220-10.html',
       'Final profitability after all expenses'],
       
      ['Return on Assets (ROA) %', '=net_income/total_assets*100',
       'Asset utilization efficiency',
       'SEC Financial Analysis Requirements',
       'https://www.sec.gov/files/aqfs.pdf',
       'Compare to industry and cost of capital'],
       
      ['Return on Equity (ROE) %', '=net_income/shareholders_equity*100',
       'Return to equity investors',
       'SEC Financial Analysis Requirements', 
       'https://www.sec.gov/files/aqfs.pdf',
       'Should exceed cost of equity'],
       
      // Leverage Ratios
      ['Debt-to-Equity', '=total_debt/total_equity',
       'Financial leverage measurement',
       'GAAP ASC 470 Debt',
       'https://www.fasb.org/page/PageContent?pageId=/asc/470/470-10.html',
       'Lower ratios indicate less financial risk'],
       
      ['Debt-to-Assets', '=total_debt/total_assets',
       'Proportion of assets financed by debt',
       'GAAP ASC 470 Debt',
       'https://www.fasb.org/page/PageContent?pageId=/asc/470/470-10.html',
       'Should align with industry standards'],
       
      ['Interest Coverage', '=ebit/interest_expense',
       'Ability to service interest payments',
       'GAAP ASC 470 Debt',
       'https://www.fasb.org/page/PageContent?pageId=/asc/470/470-10.html',
       'Higher ratios indicate lower default risk'],
       
      // Real Estate Specific
      ['Cap Rate %', '=noi/property_value*100',
       'Real estate investment return measure',
       'Real Estate Investment Analysis',
       'https://www.appraisalinstitute.org/education/',
       'Compare to market cap rates for similar properties'],
       
      ['Cash-on-Cash Return %', '=annual_cash_flow/cash_invested*100',
       'Return on actual cash invested',
       'Real Estate Investment Analysis',
       'https://www.reit.com/investing/measuring-reit-performance',
       'Excludes financing leverage effects'],
       
      ['Debt Service Coverage (DSCR)', '=noi/annual_debt_service',
       'Ability to service debt from property income',
       'Banking/Lending Standards',
       'https://www.occ.gov/publications-and-resources/publications/commercial-real-estate/',
       'DSCR > 1.25 typically required by lenders'],
       
      ['Gross Rent Multiplier', '=property_value/gross_annual_rent',
       'Property valuation relative to rent',
       'Real Estate Investment Analysis',
       'https://www.appraisalinstitute.org/education/',
       'Lower GRM may indicate better value'],
       
      // Cash Flow Analysis
      ['Operating Cash Flow Margin', '=operating_cash_flow/revenue*100',
       'Cash generation from operations',
       'GAAP ASC 230 Statement of Cash Flows',
       'https://www.fasb.org/page/PageContent?pageId=/asc/230/230-10.html',
       'Higher margins indicate quality earnings'],
       
      ['Free Cash Flow', '=operating_cash_flow-capital_expenditures',
       'Cash available after necessary investments',
       'GAAP ASC 230 Statement of Cash Flows',
       'https://www.fasb.org/page/PageContent?pageId=/asc/230/230-10.html',
       'Positive FCF indicates financial flexibility'],
       
      ['Cash Conversion Cycle', '=dso+days_inventory_outstanding-days_payable_outstanding',
       'Time to convert investments to cash',
       'Working Capital Management',
       'https://www.imanet.org/insights-and-trends/management-accounting/',
       'Shorter cycles indicate efficient working capital management'],
       
      // Working Capital
      ['Working Capital', '=current_assets-current_liabilities',
       'Short-term operational liquidity',
       'GAAP ASC 210 Balance Sheet',
       'https://www.fasb.org/page/PageContent?pageId=/asc/210/210-10.html',
       'Positive indicates ability to meet obligations'],
       
      ['Working Capital Ratio', '=current_assets/current_liabilities',
       'Same as current ratio',
       'GAAP ASC 210 Balance Sheet',
       'https://www.fasb.org/page/PageContent?pageId=/asc/210/210-10.html',
       'Benchmark: 1.2 to 2.0 typically healthy'],
       
      // Market Valuation
      ['Price-to-Book', '=market_cap/book_value',
       'Market valuation vs accounting book value',
       'SEC Market Analysis',
       'https://www.sec.gov/files/aqfs.pdf',
       'P/B > 1 indicates market premium to book value'],
       
      ['Enterprise Value', '=market_cap+total_debt-cash',
       'Total company valuation',
       'SEC Market Analysis',
       'https://www.sec.gov/files/aqfs.pdf',
       'Used for acquisition analysis'],
       
      ['', '', '', '', '', ''],
      ['IMPORTANT NOTES:', '', '', '', '', ''],
      ['1. All formulas follow GAAP/IFRS standards where applicable', '', '', '', '', ''],
      ['2. Non-GAAP measures must be reconciled to GAAP equivalents', '', '', '', '', ''],
      ['3. Industry benchmarks should be used for ratio analysis', '', '', '', '', ''],
      ['4. Formulas assume consistent accounting periods and methods', '', '', '', '', ''],
      ['5. Professional judgment required for interpretation', '', '', '', '', ''],
      ['', '', '', '', '', ''],
      ['For additional guidance, consult:', '', '', '', '', ''],
      ['- FASB Accounting Standards Codification (ASC)', '', '', '', '', ''],
      ['- SEC Financial Reporting Guidelines', '', '', '', '', ''],
      ['- AICPA Professional Standards', '', '', '', '', ''],
      ['- Industry-specific accounting guidance', '', '', '', '', '']
    ];
  }
}