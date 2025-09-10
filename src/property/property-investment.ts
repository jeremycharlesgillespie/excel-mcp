import { CellValue } from '../excel/excel-manager.js';

export interface PropertyInvestment {
  propertyId: string;
  propertyName: string;
  acquisitionPrice: number;
  acquisitionDate: string;
  propertyType: 'residential' | 'commercial' | 'retail' | 'office' | 'industrial' | 'mixed_use';
  squareFeet: number;
  units: number;
  currentMarketValue: number;
  financingDetails: {
    loanAmount: number;
    interestRate: number;
    loanTermYears: number;
    monthlyPayment: number;
    remainingBalance: number;
  };
  operatingMetrics: {
    grossRentalIncome: number;
    operatingExpenses: number;
    netOperatingIncome: number;
    capRate: number;
    occupancyRate: number;
  };
  taxInformation: {
    annualPropertyTaxes: number;
    depreciation: number;
    lastAppraisalValue: number;
    assessedValue: number;
  };
}

export interface InvestmentAnalysis {
  propertyId: string;
  analysisDate: string;
  returnMetrics: {
    currentCapRate: number;
    cashOnCashReturn: number;
    totalReturn: number; // Including appreciation
    internalRateOfReturn: number;
    cashFlow: number; // Monthly
    equityPosition: number;
    totalEquityGain: number; // Since acquisition
  };
  riskMetrics: {
    loanToValueRatio: number;
    debtServiceCoverageRatio: number;
    vacancyRisk: number;
    marketRisk: 'low' | 'medium' | 'high';
    interestRateRisk: number;
    concentrationRisk: number;
  };
  marketAnalysis: {
    appreciationRate: number; // Annual
    rentGrowthRate: number;
    marketCapRate: number;
    comparableValues: number[];
    marketTrend: 'improving' | 'stable' | 'declining';
  };
  cashFlowProjections: {
    year1: number;
    year2: number;
    year3: number;
    year5: number;
    year10: number;
    termininalValue: number;
  };
}

export interface PortfolioAnalysis {
  totalProperties: number;
  totalValue: number;
  totalEquity: number;
  totalDebt: number;
  portfolioMetrics: {
    weightedCapRate: number;
    portfolioCashFlow: number;
    totalNOI: number;
    occupancyRate: number;
    diversification: {
      byType: { [type: string]: number };
      byGeography: { [location: string]: number };
      byVintage: { [era: string]: number };
    };
  };
  riskProfile: {
    portfolioRisk: 'conservative' | 'moderate' | 'aggressive';
    concentrationRisk: number;
    leverageRatio: number;
    interestRateExposure: number;
  };
  performanceMetrics: {
    portfolioIRR: number;
    totalReturnYTD: number;
    bestPerformer: string;
    worstPerformer: string;
    averageHoldingPeriod: number; // months
  };
}

export class PropertyInvestmentAnalyzer {

  private calculateReturnMetrics(property: PropertyInvestment, analysisDate: Date): InvestmentAnalysis['returnMetrics'] {
    const acquisitionDate = new Date(property.acquisitionDate);
    const holdingPeriodYears = (analysisDate.getTime() - acquisitionDate.getTime()) / (365.25 * 24 * 60 * 60 * 1000);
    
    // Current cap rate
    const currentCapRate = property.operatingMetrics.netOperatingIncome / property.currentMarketValue;
    
    // Cash on cash return (annual NOI - debt service / initial equity)
    const annualDebtService = property.financingDetails.monthlyPayment * 12;
    const initialEquity = property.acquisitionPrice - property.financingDetails.loanAmount;
    const annualCashFlow = property.operatingMetrics.netOperatingIncome - annualDebtService;
    const cashOnCashReturn = annualCashFlow / initialEquity;
    
    // Total return including appreciation
    const totalAppreciation = property.currentMarketValue - property.acquisitionPrice;
    const totalCashFlow = annualCashFlow * holdingPeriodYears;
    const totalReturn = (totalAppreciation + totalCashFlow) / property.acquisitionPrice;
    
    // IRR calculation (simplified)
    const annualizedReturn = Math.pow(1 + totalReturn, 1 / holdingPeriodYears) - 1;
    
    // Current equity position
    const currentEquity = property.currentMarketValue - property.financingDetails.remainingBalance;
    const equityGain = currentEquity - initialEquity;
    
    return {
      currentCapRate,
      cashOnCashReturn,
      totalReturn,
      internalRateOfReturn: annualizedReturn,
      cashFlow: annualCashFlow / 12, // Monthly
      equityPosition: currentEquity,
      totalEquityGain: equityGain
    };
  }

  private calculateRiskMetrics(property: PropertyInvestment): InvestmentAnalysis['riskMetrics'] {
    // Loan to Value ratio
    const loanToValueRatio = property.financingDetails.remainingBalance / property.currentMarketValue;
    
    // Debt Service Coverage Ratio
    const annualDebtService = property.financingDetails.monthlyPayment * 12;
    const debtServiceCoverageRatio = property.operatingMetrics.netOperatingIncome / annualDebtService;
    
    // Vacancy risk based on market and property type
    let vacancyRisk = 0.05; // Base 5%
    if (property.propertyType === 'retail') vacancyRisk = 0.12;
    else if (property.propertyType === 'office') vacancyRisk = 0.08;
    else if (property.propertyType === 'industrial') vacancyRisk = 0.04;
    else if (property.propertyType === 'residential') vacancyRisk = 0.05;
    
    // Market risk assessment
    let marketRisk: 'low' | 'medium' | 'high' = 'medium';
    if (debtServiceCoverageRatio > 1.5 && loanToValueRatio < 0.7) marketRisk = 'low';
    else if (debtServiceCoverageRatio < 1.2 || loanToValueRatio > 0.8) marketRisk = 'high';
    
    // Interest rate risk (impact of 1% rate increase)
    const interestRateRisk = (property.financingDetails.remainingBalance * 0.01) / property.operatingMetrics.netOperatingIncome;
    
    return {
      loanToValueRatio,
      debtServiceCoverageRatio,
      vacancyRisk,
      marketRisk,
      interestRateRisk,
      concentrationRisk: 0 // Would calculate based on portfolio context
    };
  }

  private analyzeMarketConditions(property: PropertyInvestment): InvestmentAnalysis['marketAnalysis'] {
    // Market analysis based on property type and current conditions
    let appreciationRate = 0.04; // Base 4% appreciation
    let rentGrowthRate = 0.03; // Base 3% rent growth
    let marketCapRate = 0.06; // Market average
    
    // Adjust by property type
    switch (property.propertyType) {
      case 'industrial':
        appreciationRate = 0.06;
        rentGrowthRate = 0.05;
        marketCapRate = 0.055;
        break;
      case 'residential':
        appreciationRate = 0.05;
        rentGrowthRate = 0.04;
        marketCapRate = 0.05;
        break;
      case 'office':
        appreciationRate = 0.02;
        rentGrowthRate = 0.02;
        marketCapRate = 0.07;
        break;
      case 'retail':
        appreciationRate = 0.01;
        rentGrowthRate = 0.01;
        marketCapRate = 0.08;
        break;
    }
    
    // Generate comparable values (would be from real market data)
    const baseValue = property.currentMarketValue;
    const comparableValues = [
      baseValue * 0.95,
      baseValue * 0.98,
      baseValue * 1.02,
      baseValue * 1.05,
      baseValue * 1.08
    ];
    
    const marketTrend: 'improving' | 'stable' | 'declining' = appreciationRate > 0.035 ? 'improving' : 
                                                            appreciationRate < 0.02 ? 'declining' : 'stable';
    
    return {
      appreciationRate,
      rentGrowthRate,
      marketCapRate,
      comparableValues,
      marketTrend
    };
  }

  private projectCashFlows(property: PropertyInvestment, marketAnalysis: InvestmentAnalysis['marketAnalysis']): InvestmentAnalysis['cashFlowProjections'] {
    const baseNOI = property.operatingMetrics.netOperatingIncome;
    const annualDebtService = property.financingDetails.monthlyPayment * 12;
    const rentGrowthRate = marketAnalysis.rentGrowthRate;
    
    // Project NOI growth and calculate cash flows
    const year1NOI = baseNOI * (1 + rentGrowthRate);
    const year2NOI = year1NOI * (1 + rentGrowthRate);
    const year3NOI = year2NOI * (1 + rentGrowthRate);
    const year5NOI = baseNOI * Math.pow(1 + rentGrowthRate, 5);
    const year10NOI = baseNOI * Math.pow(1 + rentGrowthRate, 10);
    
    // Calculate terminal value (year 10 NOI / terminal cap rate)
    const terminalCapRate = marketAnalysis.marketCapRate + 0.005; // Slight cap rate expansion
    const terminalValue = year10NOI / terminalCapRate;
    
    return {
      year1: year1NOI - annualDebtService,
      year2: year2NOI - annualDebtService,
      year3: year3NOI - annualDebtService,
      year5: year5NOI - annualDebtService,
      year10: year10NOI - annualDebtService,
      termininalValue: terminalValue
    };
  }

  analyzePropertyInvestment(property: PropertyInvestment, analysisDate: Date = new Date()): InvestmentAnalysis {
    const returnMetrics = this.calculateReturnMetrics(property, analysisDate);
    const riskMetrics = this.calculateRiskMetrics(property);
    const marketAnalysis = this.analyzeMarketConditions(property);
    const cashFlowProjections = this.projectCashFlows(property, marketAnalysis);
    
    return {
      propertyId: property.propertyId,
      analysisDate: analysisDate.toISOString(),
      returnMetrics,
      riskMetrics,
      marketAnalysis,
      cashFlowProjections
    };
  }

  analyzePortfolio(properties: PropertyInvestment[], analysisDate: Date = new Date()): PortfolioAnalysis {
    const totalProperties = properties.length;
    const totalValue = properties.reduce((sum, prop) => sum + prop.currentMarketValue, 0);
    const totalDebt = properties.reduce((sum, prop) => sum + prop.financingDetails.remainingBalance, 0);
    const totalEquity = totalValue - totalDebt;
    
    // Portfolio metrics
    const totalNOI = properties.reduce((sum, prop) => sum + prop.operatingMetrics.netOperatingIncome, 0);
    const weightedCapRate = totalNOI / totalValue;
    const portfolioCashFlow = properties.reduce((sum, prop) => {
      const annualDebtService = prop.financingDetails.monthlyPayment * 12;
      return sum + (prop.operatingMetrics.netOperatingIncome - annualDebtService);
    }, 0);
    
    const totalUnits = properties.reduce((sum, prop) => sum + prop.units, 0);
    const occupiedUnits = properties.reduce((sum, prop) => sum + (prop.units * prop.operatingMetrics.occupancyRate), 0);
    const occupancyRate = occupiedUnits / totalUnits;
    
    // Diversification analysis
    const byType: { [type: string]: number } = {};
    const byGeography: { [location: string]: number } = { 'Portfolio': 100 }; // Simplified
    const byVintage: { [era: string]: number } = {};
    
    properties.forEach(prop => {
      byType[prop.propertyType] = (byType[prop.propertyType] || 0) + prop.currentMarketValue;
      
      // Vintage analysis based on acquisition date
      const acqYear = new Date(prop.acquisitionDate).getFullYear();
      const vintage = acqYear < 2015 ? 'Pre-2015' : acqYear < 2020 ? '2015-2019' : '2020+';
      byVintage[vintage] = (byVintage[vintage] || 0) + prop.currentMarketValue;
    });
    
    // Convert to percentages
    Object.keys(byType).forEach(type => {
      byType[type] = (byType[type] / totalValue * 100);
    });
    Object.keys(byVintage).forEach(vintage => {
      byVintage[vintage] = (byVintage[vintage] / totalValue * 100);
    });
    
    // Risk profile assessment
    const leverageRatio = totalDebt / totalValue;
    let portfolioRisk: 'conservative' | 'moderate' | 'aggressive' = 'moderate';
    
    if (leverageRatio < 0.5 && weightedCapRate > 0.06) portfolioRisk = 'conservative';
    else if (leverageRatio > 0.75 || weightedCapRate < 0.05) portfolioRisk = 'aggressive';
    
    // Performance analysis
    const analyses = properties.map(prop => this.analyzePropertyInvestment(prop, analysisDate));
    const portfolioIRR = analyses.reduce((sum, analysis) => sum + analysis.returnMetrics.internalRateOfReturn, 0) / analyses.length;
    
    let bestPerformer = properties[0].propertyId;
    let worstPerformer = properties[0].propertyId;
    let bestReturn = analyses[0].returnMetrics.totalReturn;
    let worstReturn = analyses[0].returnMetrics.totalReturn;
    
    analyses.forEach((analysis, index) => {
      if (analysis.returnMetrics.totalReturn > bestReturn) {
        bestReturn = analysis.returnMetrics.totalReturn;
        bestPerformer = properties[index].propertyId;
      }
      if (analysis.returnMetrics.totalReturn < worstReturn) {
        worstReturn = analysis.returnMetrics.totalReturn;
        worstPerformer = properties[index].propertyId;
      }
    });
    
    // Calculate average holding period
    const totalHoldingMonths = properties.reduce((sum, prop) => {
      const acqDate = new Date(prop.acquisitionDate);
      const holdingMonths = (analysisDate.getTime() - acqDate.getTime()) / (30.44 * 24 * 60 * 60 * 1000);
      return sum + holdingMonths;
    }, 0);
    const averageHoldingPeriod = totalHoldingMonths / totalProperties;
    
    return {
      totalProperties,
      totalValue,
      totalEquity,
      totalDebt,
      portfolioMetrics: {
        weightedCapRate,
        portfolioCashFlow,
        totalNOI,
        occupancyRate,
        diversification: {
          byType,
          byGeography,
          byVintage
        }
      },
      riskProfile: {
        portfolioRisk,
        concentrationRisk: Math.max(...Object.values(byType)) / 100, // Highest concentration
        leverageRatio,
        interestRateExposure: totalDebt / totalNOI // Debt per dollar of NOI
      },
      performanceMetrics: {
        portfolioIRR,
        totalReturnYTD: portfolioIRR, // Simplified
        bestPerformer,
        worstPerformer,
        averageHoldingPeriod
      }
    };
  }

  generateInvestmentAnalysisWorksheet(
    property: PropertyInvestment, 
    analysis: InvestmentAnalysis
  ): Array<Array<CellValue | string | number>> {
    const worksheet: Array<Array<CellValue | string | number>> = [];
    
    // Header
    worksheet.push(['PROPERTY INVESTMENT ANALYSIS', '', '', '', '', '', '', '', '', '']);
    worksheet.push([`Property: ${property.propertyName} (${property.propertyId})`, '', '', '', '', '', '', '', '', '']);
    worksheet.push([`Analysis Date: ${new Date(analysis.analysisDate).toLocaleDateString()}`, '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Real Estate Investment Intelligence Platform', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Executive Summary
    worksheet.push(['📊 EXECUTIVE SUMMARY', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Current Market Value', (property.currentMarketValue / 1000000).toFixed(2) + 'M', 'market_value', 'Current estimated market value']);
    worksheet.push(['Current Equity Position', (analysis.returnMetrics.equityPosition / 1000000).toFixed(2) + 'M', 'equity', 'Owner equity in property']);
    worksheet.push(['Cap Rate', (analysis.returnMetrics.currentCapRate * 100).toFixed(2) + '%', 'return', 'Net Operating Income / Market Value']);
    worksheet.push(['Cash on Cash Return', (analysis.returnMetrics.cashOnCashReturn * 100).toFixed(2) + '%', 'return', 'Annual cash flow / Initial equity']);
    worksheet.push(['Total Return Since Acquisition', (analysis.returnMetrics.totalReturn * 100).toFixed(2) + '%', 'return', 'Total return including appreciation']);
    worksheet.push(['Monthly Cash Flow', (analysis.returnMetrics.cashFlow / 1000).toFixed(1) + 'K', 'cash_flow', 'Monthly net cash flow after debt service']);
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Property Details
    worksheet.push(['🏢 PROPERTY DETAILS', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Property Type', property.propertyType.toUpperCase(), '', 'Commercial property classification']);
    worksheet.push(['Square Feet', property.squareFeet.toLocaleString(), 'sq_ft', 'Total rentable square footage']);
    worksheet.push(['Number of Units', property.units, 'units', 'Total rental units/spaces']);
    worksheet.push(['Acquisition Price', (property.acquisitionPrice / 1000000).toFixed(2) + 'M', 'cost', 'Original purchase price']);
    worksheet.push(['Acquisition Date', property.acquisitionDate, 'date', 'Date of property acquisition']);
    worksheet.push(['Current Occupancy', (property.operatingMetrics.occupancyRate * 100).toFixed(1) + '%', 'occupancy', 'Current occupancy rate']);
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Financial Performance
    worksheet.push(['💰 FINANCIAL PERFORMANCE', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Gross Rental Income', (property.operatingMetrics.grossRentalIncome / 1000).toFixed(0) + 'K', 'annual', 'Total annual rental income']);
    worksheet.push(['Operating Expenses', (property.operatingMetrics.operatingExpenses / 1000).toFixed(0) + 'K', 'annual', 'Annual operating expenses']);
    worksheet.push(['Net Operating Income', (property.operatingMetrics.netOperatingIncome / 1000).toFixed(0) + 'K', 'annual', 'NOI = Gross Income - Operating Expenses']);
    worksheet.push(['Debt Service (Annual)', (property.financingDetails.monthlyPayment * 12 / 1000).toFixed(0) + 'K', 'annual', 'Annual mortgage payments']);
    worksheet.push(['Cash Flow After Debt Service', (analysis.returnMetrics.cashFlow * 12 / 1000).toFixed(0) + 'K', 'annual', 'Net cash flow available to owner']);
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Financing Details
    worksheet.push(['🏦 FINANCING DETAILS', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Original Loan Amount', (property.financingDetails.loanAmount / 1000000).toFixed(2) + 'M', 'debt', 'Original mortgage amount']);
    worksheet.push(['Current Balance', (property.financingDetails.remainingBalance / 1000000).toFixed(2) + 'M', 'debt', 'Current outstanding loan balance']);
    worksheet.push(['Interest Rate', (property.financingDetails.interestRate * 100).toFixed(3) + '%', 'rate', 'Current mortgage interest rate']);
    worksheet.push(['Loan Term', property.financingDetails.loanTermYears + ' years', 'term', 'Original loan term in years']);
    worksheet.push(['Monthly Payment', (property.financingDetails.monthlyPayment / 1000).toFixed(1) + 'K', 'monthly', 'Principal and interest payment']);
    worksheet.push(['Current LTV Ratio', (analysis.riskMetrics.loanToValueRatio * 100).toFixed(1) + '%', 'leverage', 'Loan balance / Current market value']);
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Return Analysis
    worksheet.push(['📈 RETURN ANALYSIS', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Internal Rate of Return', (analysis.returnMetrics.internalRateOfReturn * 100).toFixed(2) + '%', 'return', 'Annualized total return including cash flow and appreciation']);
    worksheet.push(['Cash on Cash Return', (analysis.returnMetrics.cashOnCashReturn * 100).toFixed(2) + '%', 'return', 'Annual cash flow return on initial equity']);
    worksheet.push(['Capitalization Rate', (analysis.returnMetrics.currentCapRate * 100).toFixed(2) + '%', 'return', 'NOI / Current market value']);
    worksheet.push(['Total Equity Gain', (analysis.returnMetrics.totalEquityGain / 1000000).toFixed(2) + 'M', 'equity', 'Total equity appreciation since acquisition']);
    worksheet.push(['Market Cap Rate', (analysis.marketAnalysis.marketCapRate * 100).toFixed(2) + '%', 'market', 'Average market cap rate for comparable properties']);
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Risk Assessment
    worksheet.push(['⚠️ RISK ASSESSMENT', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Market Risk Level', analysis.riskMetrics.marketRisk.toUpperCase(), 'risk', 'Overall market risk assessment']);
    worksheet.push(['Debt Service Coverage', analysis.riskMetrics.debtServiceCoverageRatio.toFixed(2) + 'x', 'coverage', 'NOI / Annual debt service (>1.25x preferred)']);
    worksheet.push(['Vacancy Risk', (analysis.riskMetrics.vacancyRisk * 100).toFixed(1) + '%', 'risk', 'Estimated vacancy risk based on market conditions']);
    worksheet.push(['Interest Rate Risk', (analysis.riskMetrics.interestRateRisk * 100).toFixed(1) + '%', 'risk', 'Impact of 1% interest rate increase on returns']);
    worksheet.push(['Market Trend', analysis.marketAnalysis.marketTrend.toUpperCase(), 'trend', 'Current market direction']);
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Cash Flow Projections
    worksheet.push(['💵 CASH FLOW PROJECTIONS', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Year 1 Cash Flow', (analysis.cashFlowProjections.year1 / 1000).toFixed(0) + 'K', 'projection', 'Projected cash flow year 1']);
    worksheet.push(['Year 2 Cash Flow', (analysis.cashFlowProjections.year2 / 1000).toFixed(0) + 'K', 'projection', 'Projected cash flow year 2']);
    worksheet.push(['Year 3 Cash Flow', (analysis.cashFlowProjections.year3 / 1000).toFixed(0) + 'K', 'projection', 'Projected cash flow year 3']);
    worksheet.push(['Year 5 Cash Flow', (analysis.cashFlowProjections.year5 / 1000).toFixed(0) + 'K', 'projection', 'Projected cash flow year 5']);
    worksheet.push(['Year 10 Cash Flow', (analysis.cashFlowProjections.year10 / 1000).toFixed(0) + 'K', 'projection', 'Projected cash flow year 10']);
    worksheet.push(['Terminal Value (Year 10)', (analysis.cashFlowProjections.termininalValue / 1000000).toFixed(1) + 'M', 'valuation', 'Estimated sale value in year 10']);
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Market Analysis
    worksheet.push(['🏙️ MARKET ANALYSIS', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Expected Appreciation Rate', (analysis.marketAnalysis.appreciationRate * 100).toFixed(2) + '%/yr', 'growth', 'Annual property value appreciation expectation']);
    worksheet.push(['Rent Growth Rate', (analysis.marketAnalysis.rentGrowthRate * 100).toFixed(2) + '%/yr', 'growth', 'Annual rental income growth expectation']);
    worksheet.push(['Market Comparable Low', (Math.min(...analysis.marketAnalysis.comparableValues) / 1000000).toFixed(2) + 'M', 'comparable', 'Lowest comparable sale value']);
    worksheet.push(['Market Comparable High', (Math.max(...analysis.marketAnalysis.comparableValues) / 1000000).toFixed(2) + 'M', 'comparable', 'Highest comparable sale value']);
    worksheet.push(['Market Comparable Average', (analysis.marketAnalysis.comparableValues.reduce((a, b) => a + b) / analysis.marketAnalysis.comparableValues.length / 1000000).toFixed(2) + 'M', 'comparable', 'Average comparable sale value']);
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Tax Information  
    worksheet.push(['📋 TAX INFORMATION', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Annual Property Taxes', (property.taxInformation.annualPropertyTaxes / 1000).toFixed(0) + 'K', 'tax', 'Annual property tax expense']);
    worksheet.push(['Annual Depreciation', (property.taxInformation.depreciation / 1000).toFixed(0) + 'K', 'tax', 'Annual depreciation deduction']);
    worksheet.push(['Last Appraisal Value', (property.taxInformation.lastAppraisalValue / 1000000).toFixed(2) + 'M', 'appraisal', 'Most recent professional appraisal']);
    worksheet.push(['Assessed Value', (property.taxInformation.assessedValue / 1000000).toFixed(2) + 'M', 'assessment', 'Current tax assessed value']);
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Investment Recommendations
    worksheet.push(['💡 INVESTMENT RECOMMENDATIONS', '', '', '', '', '', '', '', '', '']);
    
    let recommendation = 'HOLD';
    let reasoning = 'Property performance meets expectations';
    
    if (analysis.returnMetrics.internalRateOfReturn > 0.12 && analysis.riskMetrics.marketRisk === 'low') {
      recommendation = 'STRONG HOLD';
      reasoning = 'Excellent returns with low risk profile';
    } else if (analysis.returnMetrics.internalRateOfReturn < 0.06 || analysis.riskMetrics.marketRisk === 'high') {
      recommendation = 'CONSIDER SALE';
      reasoning = 'Below-market returns or high risk concerns';
    } else if (analysis.marketAnalysis.marketTrend === 'improving' && analysis.returnMetrics.currentCapRate > analysis.marketAnalysis.marketCapRate) {
      recommendation = 'HOLD/IMPROVE';
      reasoning = 'Market improving, property outperforming';
    }
    
    worksheet.push(['Investment Recommendation', recommendation, 'recommendation', reasoning]);
    worksheet.push(['Key Reasoning', reasoning, 'rationale', 'Primary factors supporting recommendation']);
    
    // Action items based on analysis
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['📋 ACTION ITEMS', '', '', '', '', '', '', '', '', '']);
    
    if (analysis.riskMetrics.debtServiceCoverageRatio < 1.25) {
      worksheet.push(['• Improve debt service coverage through rent increases or expense reduction', '', '', '', '', '', '', '', '', '']);
    }
    if (analysis.riskMetrics.loanToValueRatio > 0.8) {
      worksheet.push(['• Consider debt paydown to improve leverage ratio', '', '', '', '', '', '', '', '', '']);
    }
    if (property.operatingMetrics.occupancyRate < 0.9) {
      worksheet.push(['• Focus on improving occupancy through leasing and tenant retention', '', '', '', '', '', '', '', '', '']);
    }
    if (analysis.returnMetrics.currentCapRate < analysis.marketAnalysis.marketCapRate - 0.005) {
      worksheet.push(['• Evaluate opportunities for rent increases to market levels', '', '', '', '', '', '', '', '', '']);
    }
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Professional Notes
    worksheet.push(['📝 ANALYSIS METHODOLOGY & DISCLAIMERS', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- Market values based on comparable sales and income approach', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- Cash flow projections assume current market rent growth trends', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- Risk assessments based on market conditions and property fundamentals', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- Tax implications not included in return calculations', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- Consult real estate and tax professionals for investment decisions', '', '', '', '', '', '', '', '', '']);
    
    return worksheet;
  }

  generatePortfolioAnalysisWorksheet(
    properties: PropertyInvestment[], 
    portfolioAnalysis: PortfolioAnalysis
  ): Array<Array<CellValue | string | number>> {
    const worksheet: Array<Array<CellValue | string | number>> = [];
    
    // Header
    worksheet.push(['REAL ESTATE PORTFOLIO ANALYSIS', '', '', '', '', '', '', '', '', '']);
    worksheet.push([`Portfolio Analysis Date: ${new Date().toLocaleDateString()}`, '', '', '', '', '', '', '', '', '']);
    worksheet.push([`Total Properties: ${portfolioAnalysis.totalProperties}`, '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Real Estate Investment Intelligence Platform', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Portfolio Summary
    worksheet.push(['📊 PORTFOLIO SUMMARY', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Total Portfolio Value', (portfolioAnalysis.totalValue / 1000000).toFixed(1) + 'M', 'value', 'Combined market value of all properties']);
    worksheet.push(['Total Equity', (portfolioAnalysis.totalEquity / 1000000).toFixed(1) + 'M', 'equity', 'Combined owner equity across portfolio']);
    worksheet.push(['Total Debt', (portfolioAnalysis.totalDebt / 1000000).toFixed(1) + 'M', 'debt', 'Combined outstanding loan balances']);
    worksheet.push(['Portfolio Leverage', (portfolioAnalysis.riskProfile.leverageRatio * 100).toFixed(1) + '%', 'leverage', 'Total debt / Total value']);
    worksheet.push(['Weighted Cap Rate', (portfolioAnalysis.portfolioMetrics.weightedCapRate * 100).toFixed(2) + '%', 'return', 'Portfolio-wide capitalization rate']);
    worksheet.push(['Annual Cash Flow', (portfolioAnalysis.portfolioMetrics.portfolioCashFlow / 1000).toFixed(0) + 'K', 'cash_flow', 'Total annual cash flow after debt service']);
    worksheet.push(['Portfolio Occupancy', (portfolioAnalysis.portfolioMetrics.occupancyRate * 100).toFixed(1) + '%', 'occupancy', 'Weighted average occupancy rate']);
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Performance Metrics
    worksheet.push(['📈 PERFORMANCE METRICS', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Portfolio IRR', (portfolioAnalysis.performanceMetrics.portfolioIRR * 100).toFixed(2) + '%', 'return', 'Weighted average internal rate of return']);
    worksheet.push(['Total NOI', (portfolioAnalysis.portfolioMetrics.totalNOI / 1000).toFixed(0) + 'K', 'income', 'Combined net operating income']);
    worksheet.push(['Best Performing Property', portfolioAnalysis.performanceMetrics.bestPerformer, 'performance', 'Highest total return property']);
    worksheet.push(['Worst Performing Property', portfolioAnalysis.performanceMetrics.worstPerformer, 'performance', 'Lowest total return property']);
    worksheet.push(['Average Holding Period', (portfolioAnalysis.performanceMetrics.averageHoldingPeriod / 12).toFixed(1) + ' years', 'duration', 'Average years since acquisition']);
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Diversification Analysis
    worksheet.push(['🎯 DIVERSIFICATION ANALYSIS', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Property Type Breakdown:', '', '', '', '', '', '', '', '', '']);
    
    Object.entries(portfolioAnalysis.portfolioMetrics.diversification.byType).forEach(([type, percentage]) => {
      worksheet.push([`  ${type.toUpperCase()}`, percentage.toFixed(1) + '%', 'allocation', 'Percentage of total portfolio value']);
    });
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Vintage Analysis:', '', '', '', '', '', '', '', '', '']);
    
    Object.entries(portfolioAnalysis.portfolioMetrics.diversification.byVintage).forEach(([vintage, percentage]) => {
      worksheet.push([`  ${vintage}`, percentage.toFixed(1) + '%', 'allocation', 'Percentage by acquisition period']);
    });
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Risk Assessment
    worksheet.push(['⚠️ PORTFOLIO RISK ASSESSMENT', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Portfolio Risk Profile', portfolioAnalysis.riskProfile.portfolioRisk.toUpperCase(), 'risk', 'Overall portfolio risk classification']);
    worksheet.push(['Concentration Risk', (portfolioAnalysis.riskProfile.concentrationRisk * 100).toFixed(1) + '%', 'risk', 'Highest single property type concentration']);
    worksheet.push(['Interest Rate Exposure', portfolioAnalysis.riskProfile.interestRateExposure.toFixed(2) + 'x', 'exposure', 'Total debt per dollar of NOI']);
    
    let riskRecommendation = 'Portfolio risk within acceptable parameters';
    if (portfolioAnalysis.riskProfile.concentrationRisk > 0.5) {
      riskRecommendation = 'HIGH CONCENTRATION RISK - Consider diversification';
    } else if (portfolioAnalysis.riskProfile.leverageRatio > 0.75) {
      riskRecommendation = 'HIGH LEVERAGE RISK - Consider debt reduction';
    } else if (portfolioAnalysis.riskProfile.portfolioRisk === 'aggressive') {
      riskRecommendation = 'AGGRESSIVE PROFILE - Monitor market conditions closely';
    }
    
    worksheet.push(['Risk Assessment', riskRecommendation, 'assessment', 'Overall risk evaluation and recommendations']);
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Individual Property Performance
    worksheet.push(['🏢 INDIVIDUAL PROPERTY PERFORMANCE', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Property ID', 'Type', 'Value', 'Cap Rate', 'Cash Flow', 'LTV', 'Performance', '', '', '']);
    
    properties.forEach(property => {
      const analysis = this.analyzePropertyInvestment(property);
      let performance = 'MEETS EXPECTATIONS';
      
      if (analysis.returnMetrics.internalRateOfReturn > 0.12) performance = 'OUTPERFORMING';
      else if (analysis.returnMetrics.internalRateOfReturn < 0.06) performance = 'UNDERPERFORMING';
      
      worksheet.push([
        property.propertyId,
        property.propertyType.toUpperCase(),
        (property.currentMarketValue / 1000000).toFixed(1) + 'M',
        (analysis.returnMetrics.currentCapRate * 100).toFixed(2) + '%',
        (analysis.returnMetrics.cashFlow / 1000).toFixed(0) + 'K/mo',
        (analysis.riskMetrics.loanToValueRatio * 100).toFixed(1) + '%',
        performance,
        '', '', ''
      ]);
    });
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Portfolio Strategy Recommendations
    worksheet.push(['💡 PORTFOLIO STRATEGY RECOMMENDATIONS', '', '', '', '', '', '', '', '', '']);
    
    const recommendations = [];
    
    // Concentration analysis
    if (portfolioAnalysis.riskProfile.concentrationRisk > 0.4) {
      recommendations.push('Diversify property types - current concentration above 40%');
    }
    
    // Performance analysis
    if (portfolioAnalysis.performanceMetrics.portfolioIRR < 0.08) {
      recommendations.push('Portfolio IRR below 8% - consider asset optimization or disposition');
    }
    
    // Leverage analysis
    if (portfolioAnalysis.riskProfile.leverageRatio > 0.7) {
      recommendations.push('High leverage ratio - consider debt paydown or equity infusion');
    } else if (portfolioAnalysis.riskProfile.leverageRatio < 0.5) {
      recommendations.push('Conservative leverage - consider debt optimization for tax benefits');
    }
    
    // Occupancy analysis
    if (portfolioAnalysis.portfolioMetrics.occupancyRate < 0.9) {
      recommendations.push('Focus on improving occupancy through leasing and property management');
    }
    
    // Cap rate analysis
    if (portfolioAnalysis.portfolioMetrics.weightedCapRate < 0.055) {
      recommendations.push('Low cap rate market - consider strategic acquisitions or dispositions');
    }
    
    recommendations.forEach((rec, index) => {
      worksheet.push([`${index + 1}.`, rec, '', '', '', '', '', '', '', '']);
    });
    
    if (recommendations.length === 0) {
      worksheet.push(['Portfolio performance and risk profile within target parameters', '', '', '', '', '', '', '', '', '']);
    }
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Market Outlook and Action Plan
    worksheet.push(['🔮 MARKET OUTLOOK & ACTION PLAN', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Next 90 Days:', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['• Complete quarterly portfolio valuation updates', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['• Review and optimize financing for properties with upcoming maturity', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['• Analyze acquisition opportunities in underrepresented markets', '', '', '', '', '', '', '', '', '']);
    
    worksheet.push(['Next 6 Months:', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['• Strategic review of underperforming assets', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['• Market rent analysis and lease optimization', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['• Portfolio diversification strategy implementation', '', '', '', '', '', '', '', '', '']);
    
    worksheet.push(['Next 12 Months:', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['• Annual portfolio strategic plan review', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['• Tax optimization and 1031 exchange planning', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['• Market cycle positioning and risk management', '', '', '', '', '', '', '', '', '']);
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Professional Disclaimers
    worksheet.push(['📝 ANALYSIS METHODOLOGY & DISCLAIMERS', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- Portfolio analysis based on current market conditions and property fundamentals', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- Performance projections based on historical trends and market analysis', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- Risk assessments reflect current property and market conditions', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- Consult real estate, tax, and legal professionals before making investment decisions', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- Past performance does not guarantee future results', '', '', '', '', '', '', '', '', '']);
    
    return worksheet;
  }
}