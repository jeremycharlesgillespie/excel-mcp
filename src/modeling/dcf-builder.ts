import { CellValue } from '../excel/excel-manager.js';

export interface DCFProjection {
  year: number;
  revenue: number;
  revenueGrowthRate: number;
  grossMargin: number;
  operatingMargin: number;
  taxRate: number;
  capex: number;
  capexAsPercentOfRevenue: number;
  workingCapitalChange: number;
  depreciation: number;
}

export interface DCFAssumptions {
  projectionYears: number;
  discountRate: number; // WACC
  terminalGrowthRate: number;
  initialRevenue: number;
  revenueCAGR: number; // Compound Annual Growth Rate
  terminalMargin: number;
  terminalCapexRate: number; // Terminal capex as % of revenue
  terminalTaxRate: number;
}

export interface DCFValuation {
  projections: DCFProjection[];
  freeCashFlows: number[];
  presentValues: number[];
  terminalValue: number;
  terminalValuePV: number;
  enterpriseValue: number;
  sensitivityAnalysis: {
    waccRange: number[];
    terminalGrowthRange: number[];
    valuationMatrix: number[][];
  };
  keyMetrics: {
    impliedEVRevenue: number; // EV/Revenue multiple
    impliedEVEBITDA: number;   // EV/EBITDA multiple
    breakEvenWACC: number;     // WACC at which NPV = 0
  };
}

export class DCFBuilder {
  
  private calculateProjection(
    year: number,
    baseRevenue: number,
    assumptions: DCFAssumptions
  ): DCFProjection {
    
    // Revenue growth with potential deceleration
    const growthRate = year <= 5 
      ? assumptions.revenueCAGR * (1 - (year - 1) * 0.02) // 2% deceleration per year
      : assumptions.terminalGrowthRate;
    
    const revenue = year === 1 
      ? baseRevenue * (1 + growthRate)
      : baseRevenue * Math.pow(1 + assumptions.revenueCAGR, year);
    
    // Margins progression toward terminal values
    const marginProgression = Math.min(1, year / assumptions.projectionYears);
    const grossMargin = 0.6 + (assumptions.terminalMargin - 0.6) * marginProgression;
    const operatingMargin = grossMargin * 0.8; // Assume 80% of gross margin
    
    // Capex as % of revenue, declining over time
    const capexAsPercentOfRevenue = year <= 3 
      ? 0.08 - (year - 1) * 0.01  // High early capex, declining
      : assumptions.terminalCapexRate;
    
    const capex = revenue * capexAsPercentOfRevenue;
    const depreciation = capex * 0.85; // Assume 85% of capex becomes depreciation
    
    // Working capital change (2% of revenue increase)
    const workingCapitalChange = year === 1 
      ? revenue * 0.02 
      : (revenue - baseRevenue * Math.pow(1 + assumptions.revenueCAGR, year - 1)) * 0.02;
    
    return {
      year,
      revenue,
      revenueGrowthRate: growthRate,
      grossMargin,
      operatingMargin,
      taxRate: assumptions.terminalTaxRate,
      capex,
      capexAsPercentOfRevenue,
      workingCapitalChange,
      depreciation
    };
  }

  private calculateFreeCashFlow(projection: DCFProjection): number {
    const operatingIncome = projection.revenue * projection.operatingMargin;
    const ebit = operatingIncome;
    const taxes = ebit * projection.taxRate;
    const nopat = ebit - taxes; // Net Operating Profit After Tax
    
    // Free Cash Flow = NOPAT + Depreciation - Capex - Working Capital Change
    const freeCashFlow = nopat + projection.depreciation - projection.capex - projection.workingCapitalChange;
    
    return freeCashFlow;
  }

  private calculateTerminalValue(
    finalProjection: DCFProjection,
    assumptions: DCFAssumptions
  ): number {
    const finalFCF = this.calculateFreeCashFlow(finalProjection);
    const terminalFCF = finalFCF * (1 + assumptions.terminalGrowthRate);
    
    // Terminal Value = Terminal FCF / (WACC - Terminal Growth Rate)
    const terminalValue = terminalFCF / (assumptions.discountRate - assumptions.terminalGrowthRate);
    
    return terminalValue;
  }

  private createSensitivityAnalysis(
    projections: DCFProjection[],
    baseAssumptions: DCFAssumptions
  ): { waccRange: number[]; terminalGrowthRange: number[]; valuationMatrix: number[][] } {
    
    const waccRange = [];
    const terminalGrowthRange = [];
    const valuationMatrix: number[][] = [];
    
    // Create ranges
    for (let i = -0.02; i <= 0.02; i += 0.005) {
      waccRange.push(baseAssumptions.discountRate + i);
    }
    
    for (let i = -0.01; i <= 0.01; i += 0.0025) {
      terminalGrowthRange.push(baseAssumptions.terminalGrowthRate + i);
    }
    
    // Calculate valuations for each combination
    for (const wacc of waccRange) {
      const row: number[] = [];
      for (const terminalGrowth of terminalGrowthRange) {
        
        // Calculate free cash flows
        const freeCashFlows = projections.map(p => this.calculateFreeCashFlow(p));
        
        // Calculate present values
        const presentValues = freeCashFlows.map((fcf, index) => 
          fcf / Math.pow(1 + wacc, index + 1)
        );
        
        // Calculate terminal value with adjusted assumptions
        const finalFCF = freeCashFlows[freeCashFlows.length - 1];
        const terminalFCF = finalFCF * (1 + terminalGrowth);
        const terminalValue = terminalFCF / (wacc - terminalGrowth);
        const terminalValuePV = terminalValue / Math.pow(1 + wacc, projections.length);
        
        // Enterprise Value
        const enterpriseValue = presentValues.reduce((sum, pv) => sum + pv, 0) + terminalValuePV;
        row.push(enterpriseValue);
      }
      valuationMatrix.push(row);
    }
    
    return {
      waccRange,
      terminalGrowthRange,
      valuationMatrix
    };
  }

  buildDCFModel(assumptions: DCFAssumptions): DCFValuation {
    const projections: DCFProjection[] = [];
    
    // Create projections for each year
    for (let year = 1; year <= assumptions.projectionYears; year++) {
      const projection = this.calculateProjection(year, assumptions.initialRevenue, assumptions);
      projections.push(projection);
    }
    
    // Calculate free cash flows
    const freeCashFlows = projections.map(p => this.calculateFreeCashFlow(p));
    
    // Calculate present values
    const presentValues = freeCashFlows.map((fcf, index) => 
      fcf / Math.pow(1 + assumptions.discountRate, index + 1)
    );
    
    // Calculate terminal value
    const terminalValue = this.calculateTerminalValue(
      projections[projections.length - 1], 
      assumptions
    );
    const terminalValuePV = terminalValue / Math.pow(1 + assumptions.discountRate, assumptions.projectionYears);
    
    // Enterprise Value
    const enterpriseValue = presentValues.reduce((sum, pv) => sum + pv, 0) + terminalValuePV;
    
    // Sensitivity analysis
    const sensitivityAnalysis = this.createSensitivityAnalysis(projections, assumptions);
    
    // Key metrics
    const finalProjection = projections[projections.length - 1];
    const impliedEVRevenue = enterpriseValue / finalProjection.revenue;
    const finalOperatingIncome = finalProjection.revenue * finalProjection.operatingMargin;
    const finalEBITDA = finalOperatingIncome + finalProjection.depreciation;
    const impliedEVEBITDA = enterpriseValue / finalEBITDA;
    
    // Find break-even WACC (simplified - binary search would be more accurate)
    let breakEvenWACC = assumptions.discountRate;
    for (let wacc = 0.05; wacc <= 0.25; wacc += 0.001) {
      const testPVs = freeCashFlows.map((fcf, index) => fcf / Math.pow(1 + wacc, index + 1));
      const testTerminalPV = terminalValue / Math.pow(1 + wacc, assumptions.projectionYears);
      const testEV = testPVs.reduce((sum, pv) => sum + pv, 0) + testTerminalPV;
      
      if (Math.abs(testEV) < 1000) { // Close to zero
        breakEvenWACC = wacc;
        break;
      }
    }
    
    return {
      projections,
      freeCashFlows,
      presentValues,
      terminalValue,
      terminalValuePV,
      enterpriseValue,
      sensitivityAnalysis,
      keyMetrics: {
        impliedEVRevenue,
        impliedEVEBITDA,
        breakEvenWACC
      }
    };
  }

  generateDCFWorksheet(valuation: DCFValuation, assumptions: DCFAssumptions): Array<Array<CellValue | string | number>> {
    const worksheet: Array<Array<CellValue | string | number>> = [];
    
    // Header
    worksheet.push(['DISCOUNTED CASH FLOW (DCF) VALUATION MODEL', '', '', '', '', '', '', '', '']);
    worksheet.push(['Generated by Excel Finance MCP', '', '', '', '', '', '', '', '']);
    worksheet.push(['Date: ' + new Date().toLocaleDateString(), '', '', '', '', '', '', '', '']);
    worksheet.push(['', '', '', '', '', '', '', '', '']);
    
    // Key Assumptions
    worksheet.push(['KEY ASSUMPTIONS:', '', '', '', '', '', '', '', '']);
    worksheet.push(['Assumption', 'Value', 'Unit', 'Description', 'GAAP/Standard Reference', '', '', '', '']);
    worksheet.push(['Projection Period', assumptions.projectionYears, 'years', 'Explicit forecast period', 'DCF Best Practice: 5-10 years']);
    worksheet.push(['Discount Rate (WACC)', (assumptions.discountRate * 100).toFixed(2) + '%', 'annual', 'Weighted Average Cost of Capital', 'GAAP Concept 7 - Present Value']);
    worksheet.push(['Terminal Growth Rate', (assumptions.terminalGrowthRate * 100).toFixed(2) + '%', 'annual', 'Long-term growth assumption', 'Should not exceed GDP growth']);
    worksheet.push(['Initial Revenue', assumptions.initialRevenue.toFixed(0), 'currency', 'Base year revenue', 'ASC 606 Revenue Recognition']);
    worksheet.push(['Revenue CAGR', (assumptions.revenueCAGR * 100).toFixed(2) + '%', 'annual', 'Compound Annual Growth Rate', 'Management projections']);
    worksheet.push(['Terminal Margin', (assumptions.terminalMargin * 100).toFixed(1) + '%', 'of revenue', 'Steady-state profitability', 'Industry benchmarking']);
    worksheet.push(['Terminal Tax Rate', (assumptions.terminalTaxRate * 100).toFixed(1) + '%', 'effective rate', 'Long-term tax assumption', 'ASC 740 Income Taxes']);
    worksheet.push(['', '', '', '', '', '', '', '', '']);
    
    // Financial Projections
    worksheet.push(['FINANCIAL PROJECTIONS:', '', '', '', '', '', '', '', '']);
    
    // Header row
    const projectionHeaders = ['Metric', 'Unit', ...valuation.projections.map(p => `Year ${p.year}`), 'Terminal'];
    worksheet.push(projectionHeaders);
    
    // Revenue
    const revenueRow = ['Revenue', 'millions', ...valuation.projections.map(p => (p.revenue / 1000000).toFixed(1))];
    worksheet.push(revenueRow);
    
    // Growth rates
    const growthRow = ['Revenue Growth', '%', ...valuation.projections.map(p => (p.revenueGrowthRate * 100).toFixed(1) + '%')];
    worksheet.push(growthRow);
    
    // Margins
    const grossMarginRow = ['Gross Margin', '%', ...valuation.projections.map(p => (p.grossMargin * 100).toFixed(1) + '%')];
    worksheet.push(grossMarginRow);
    
    const operatingMarginRow = ['Operating Margin', '%', ...valuation.projections.map(p => (p.operatingMargin * 100).toFixed(1) + '%')];
    worksheet.push(operatingMarginRow);
    
    // Operating Income
    const operatingIncomeRow = ['Operating Income', 'millions', ...valuation.projections.map(p => (p.revenue * p.operatingMargin / 1000000).toFixed(1))];
    worksheet.push(operatingIncomeRow);
    
    // Tax and NOPAT
    const taxRow = ['Taxes', 'millions', ...valuation.projections.map(p => (p.revenue * p.operatingMargin * p.taxRate / 1000000).toFixed(1))];
    worksheet.push(taxRow);
    
    const nopatRow = ['NOPAT', 'millions', ...valuation.projections.map(p => (p.revenue * p.operatingMargin * (1 - p.taxRate) / 1000000).toFixed(1))];
    worksheet.push(nopatRow);
    
    // Capex and Depreciation
    const capexRow = ['Capital Expenditures', 'millions', ...valuation.projections.map(p => (p.capex / 1000000).toFixed(1))];
    worksheet.push(capexRow);
    
    const depreciationRow = ['Depreciation', 'millions', ...valuation.projections.map(p => (p.depreciation / 1000000).toFixed(1))];
    worksheet.push(depreciationRow);
    
    // Working capital
    const wcRow = ['Working Capital Change', 'millions', ...valuation.projections.map(p => (p.workingCapitalChange / 1000000).toFixed(1))];
    worksheet.push(wcRow);
    
    // Free Cash Flow
    const fcfRow = ['Free Cash Flow', 'millions', ...valuation.freeCashFlows.map(fcf => (fcf / 1000000).toFixed(1))];
    worksheet.push(fcfRow);
    
    worksheet.push(['', '', '', '', '', '', '', '', '']);
    
    // Valuation
    worksheet.push(['DCF VALUATION:', '', '', '', '', '', '', '', '']);
    worksheet.push(['Component', 'Value (millions)', 'Formula', 'Description', '', '', '', '', '']);
    
    // Present values
    for (let i = 0; i < valuation.presentValues.length; i++) {
      worksheet.push([
        `Year ${i + 1} PV`,
        (valuation.presentValues[i] / 1000000).toFixed(1),
        `=FCF${i + 1}/(1+WACC)^${i + 1}`,
        `Present value of year ${i + 1} free cash flow`,
        '', '', '', '', ''
      ]);
    }
    
    worksheet.push([
      'Terminal Value',
      (valuation.terminalValue / 1000000).toFixed(1),
      '=Terminal_FCF/(WACC-g)',
      'Present value of cash flows beyond projection period',
      '', '', '', '', ''
    ]);
    
    worksheet.push([
      'Terminal Value PV',
      (valuation.terminalValuePV / 1000000).toFixed(1),
      `=Terminal_Value/(1+WACC)^${assumptions.projectionYears}`,
      'Discounted terminal value',
      '', '', '', '', ''
    ]);
    
    worksheet.push([
      'Enterprise Value',
      (valuation.enterpriseValue / 1000000).toFixed(1),
      '=SUM(PV_Years)+Terminal_Value_PV',
      'Total company valuation',
      '', '', '', '', ''
    ]);
    
    worksheet.push(['', '', '', '', '', '', '', '', '']);
    
    // Key Metrics
    worksheet.push(['VALUATION METRICS:', '', '', '', '', '', '', '', '']);
    worksheet.push(['Metric', 'Value', 'Benchmark Range', 'Interpretation', '', '', '', '', '']);
    worksheet.push(['EV/Revenue', valuation.keyMetrics.impliedEVRevenue.toFixed(1) + 'x', '1.0x - 5.0x', 'Revenue multiple implied by DCF']);
    worksheet.push(['EV/EBITDA', valuation.keyMetrics.impliedEVEBITDA.toFixed(1) + 'x', '8.0x - 15.0x', 'EBITDA multiple implied by DCF']);
    worksheet.push(['Break-even WACC', (valuation.keyMetrics.breakEvenWACC * 100).toFixed(2) + '%', 'Compare to actual WACC', 'Discount rate yielding NPV = 0']);
    
    worksheet.push(['', '', '', '', '', '', '', '', '']);
    
    // Sensitivity Analysis
    worksheet.push(['SENSITIVITY ANALYSIS:', '', '', '', '', '', '', '', '']);
    worksheet.push(['Enterprise Value sensitivity to WACC and Terminal Growth Rate', '', '', '', '', '', '', '', '']);
    
    // Sensitivity table header
    const sensHeaders = ['WACC \\ Terminal Growth', ...valuation.sensitivityAnalysis.terminalGrowthRange.map(g => (g * 100).toFixed(2) + '%')];
    worksheet.push(sensHeaders);
    
    // Sensitivity table rows
    for (let i = 0; i < valuation.sensitivityAnalysis.waccRange.length; i++) {
      const row = [
        (valuation.sensitivityAnalysis.waccRange[i] * 100).toFixed(2) + '%',
        ...valuation.sensitivityAnalysis.valuationMatrix[i].map(v => (v / 1000000).toFixed(0))
      ];
      worksheet.push(row);
    }
    
    worksheet.push(['', '', '', '', '', '', '', '', '']);
    
    // Risk Analysis
    worksheet.push(['RISK ASSESSMENT:', '', '', '', '', '', '', '', '']);
    worksheet.push(['Risk Factor', 'Impact', 'Mitigation', 'Probability', '', '', '', '', '']);
    worksheet.push(['Revenue Growth Risk', 'High', 'Conservative projections, scenario analysis', 'Medium', '', '', '', '', '']);
    worksheet.push(['Margin Compression', 'Medium', 'Cost management, operational efficiency', 'Low', '', '', '', '', '']);
    worksheet.push(['WACC Changes', 'High', 'Capital structure optimization', 'Medium', '', '', '', '', '']);
    worksheet.push(['Terminal Value Risk', 'Very High', 'Multiple valuation approaches', 'High', '', '', '', '', '']);
    worksheet.push(['Market Conditions', 'Medium', 'Stress testing, flexibility planning', 'High', '', '', '', '', '']);
    
    worksheet.push(['', '', '', '', '', '', '', '', '']);
    
    // Professional Standards
    worksheet.push(['PROFESSIONAL STANDARDS & REFERENCES:', '', '', '', '', '', '', '', '']);
    worksheet.push(['- FASB Concept 7: Using Cash Flow Information and Present Value', '', '', '', '', '', '', '', '']);
    worksheet.push(['- AICPA Business Valuation Standards', '', '', '', '', '', '', '', '']);
    worksheet.push(['- CFA Institute Equity Valuation Guidelines', '', '', '', '', '', '', '', '']);
    worksheet.push(['- ASC 820: Fair Value Measurements', '', '', '', '', '', '', '', '']);
    worksheet.push(['- International Valuation Standards (IVS)', '', '', '', '', '', '', '', '']);
    
    worksheet.push(['', '', '', '', '', '', '', '', '']);
    worksheet.push(['LIMITATIONS & ASSUMPTIONS:', '', '', '', '', '', '', '', '']);
    worksheet.push(['- Model based on management projections and market assumptions', '', '', '', '', '', '', '', '']);
    worksheet.push(['- Terminal value represents significant portion of total value', '', '', '', '', '', '', '', '']);
    worksheet.push(['- Sensitivity to key assumptions (WACC, growth rates)', '', '', '', '', '', '', '', '']);
    worksheet.push(['- Does not include control premium or liquidity discounts', '', '', '', '', '', '', '', '']);
    worksheet.push(['- Should be used in conjunction with other valuation methods', '', '', '', '', '', '', '', '']);
    
    return worksheet;
  }
}