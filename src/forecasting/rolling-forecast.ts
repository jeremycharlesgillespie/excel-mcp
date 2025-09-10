import { CellValue } from '../excel/excel-manager.js';

export interface ForecastInput {
  category: string;
  subcategory?: string;
  historical: number[]; // Historical values (e.g., last 12 months)
  seasonalityPattern?: number[]; // Seasonal adjustment factors (length should match forecast periods)
  growthRate?: number; // Annual growth rate (decimal, e.g., 0.05 for 5%)
  volatility?: number; // Standard deviation for uncertainty modeling
  driver?: string; // Name of the driver variable if this is driver-based
  driverMultiplier?: number; // Multiplier for driver-based forecasting
}

export interface ForecastDriver {
  name: string;
  values: number[]; // Future values of the driver
  description: string;
}

export interface ForecastScenario {
  name: string;
  description: string;
  adjustmentFactor: number; // Multiplier for the scenario (1.0 = base case)
  probabilityWeight: number; // Probability weight (should sum to 1.0 across scenarios)
}

export interface CashFlowCategory {
  name: string;
  subcategories: ForecastInput[];
  isInflow: boolean; // true for revenue, false for expenses
}

export interface RollingForecastResult {
  periods: string[]; // Period labels (e.g., "2024-01", "2024-02")
  forecasts: { [category: string]: number[] };
  statistics: {
    totalInflows: number[];
    totalOutflows: number[];
    netCashFlow: number[];
    cumulativeCashFlow: number[];
    runningBalance: number[];
  };
  confidence: {
    lower90: number[];
    upper90: number[];
    lower95: number[];
    upper95: number[];
  };
  keyMetrics: {
    burnRate: number; // Average monthly cash burn
    runwayMonths: number; // Months until cash runs out
    breakEvenMonth: number | null; // Month when net cash flow turns positive
    cashMinimumDate: string | null; // Date of minimum cash balance
  };
}

export class RollingForecastEngine {
  private generatePeriodLabels(startDate: Date, periods: number, frequency: 'weekly' | 'monthly'): string[] {
    const labels: string[] = [];
    const current = new Date(startDate);
    
    for (let i = 0; i < periods; i++) {
      if (frequency === 'weekly') {
        labels.push(`Week ${i + 1} (${current.toISOString().split('T')[0]})`);
        current.setDate(current.getDate() + 7);
      } else {
        labels.push(`${current.getFullYear()}-${String(current.getMonth() + 1).padStart(2, '0')}`);
        current.setMonth(current.getMonth() + 1);
      }
    }
    
    return labels;
  }

  private applySeasonality(baseValue: number, seasonalityPattern: number[] | undefined, periodIndex: number): number {
    if (!seasonalityPattern || seasonalityPattern.length === 0) {
      return baseValue;
    }
    
    const seasonalIndex = periodIndex % seasonalityPattern.length;
    return baseValue * seasonalityPattern[seasonalIndex];
  }

  private calculateTrendProjection(historical: number[], periods: number, growthRate?: number): number[] {
    if (historical.length === 0) {
      return new Array(periods).fill(0);
    }

    // Use linear regression for trend if no growth rate specified
    let trend = 0;
    if (growthRate !== undefined) {
      trend = growthRate / 12; // Convert annual to monthly growth
    } else if (historical.length > 1) {
      // Calculate trend using linear regression
      const n = historical.length;
      const xSum = (n * (n + 1)) / 2;
      const ySum = historical.reduce((sum, val) => sum + val, 0);
      const xySum = historical.reduce((sum, val, index) => sum + val * (index + 1), 0);
      const xxSum = (n * (n + 1) * (2 * n + 1)) / 6;
      
      trend = (n * xySum - xSum * ySum) / (n * xxSum - xSum * xSum);
    }

    const lastValue = historical[historical.length - 1];
    const projections: number[] = [];
    
    for (let i = 1; i <= periods; i++) {
      const trendProjection = lastValue + trend * i;
      projections.push(Math.max(0, trendProjection)); // Ensure non-negative
    }
    
    return projections;
  }

  private applyDriverBasedForecast(input: ForecastInput, drivers: ForecastDriver[], periods: number): number[] {
    if (!input.driver || !input.driverMultiplier) {
      return this.calculateTrendProjection(input.historical, periods, input.growthRate);
    }

    const driver = drivers.find(d => d.name === input.driver);
    if (!driver) {
      console.warn(`Driver ${input.driver} not found, using trend projection`);
      return this.calculateTrendProjection(input.historical, periods, input.growthRate);
    }

    return driver.values.slice(0, periods).map(driverValue => driverValue * input.driverMultiplier!);
  }

  create13WeekForecast(
    categories: CashFlowCategory[],
    drivers: ForecastDriver[] = [],
    startDate: Date = new Date()
  ): RollingForecastResult {
    
    const periods = 13;
    const periodLabels = this.generatePeriodLabels(startDate, periods, 'weekly');
    const forecasts: { [category: string]: number[] } = {};
    
    // Generate forecasts for each category
    for (const category of categories) {
      const categoryForecasts: number[] = new Array(periods).fill(0);
      
      for (const input of category.subcategories) {
        // Get base projection
        let projection = this.applyDriverBasedForecast(input, drivers, periods);
        
        // Apply seasonality
        projection = projection.map((value, index) => 
          this.applySeasonality(value, input.seasonalityPattern, index)
        );
        
        // Add to category total
        for (let i = 0; i < periods; i++) {
          categoryForecasts[i] += projection[i];
        }
      }
      
      forecasts[category.name] = categoryForecasts;
    }

    // Calculate cash flow statistics
    const totalInflows = new Array(periods).fill(0);
    const totalOutflows = new Array(periods).fill(0);
    
    for (const category of categories) {
      const categoryValues = forecasts[category.name];
      for (let i = 0; i < periods; i++) {
        if (category.isInflow) {
          totalInflows[i] += categoryValues[i];
        } else {
          totalOutflows[i] += categoryValues[i];
        }
      }
    }
    
    const netCashFlow = totalInflows.map((inflow, index) => inflow - totalOutflows[index]);
    const cumulativeCashFlow = netCashFlow.reduce((acc, flow, index) => {
      acc.push(index === 0 ? flow : acc[index - 1] + flow);
      return acc;
    }, [] as number[]);
    
    // Assume starting cash balance (can be parameterized)
    const initialCashBalance = 100000; // This should be a parameter
    const runningBalance = cumulativeCashFlow.map(cumulative => initialCashBalance + cumulative);
    
    // Calculate confidence intervals (simplified - in practice would use more sophisticated modeling)
    const volatility = 0.15; // 15% standard deviation
    const lower90 = netCashFlow.map(flow => flow * (1 - 1.645 * volatility));
    const upper90 = netCashFlow.map(flow => flow * (1 + 1.645 * volatility));
    const lower95 = netCashFlow.map(flow => flow * (1 - 1.96 * volatility));
    const upper95 = netCashFlow.map(flow => flow * (1 + 1.96 * volatility));
    
    // Calculate key metrics
    const avgNegativeCashFlow = netCashFlow.filter(flow => flow < 0).reduce((sum, flow) => sum + flow, 0) / 
                               Math.max(1, netCashFlow.filter(flow => flow < 0).length);
    const burnRate = Math.abs(avgNegativeCashFlow);
    const runwayMonths = initialCashBalance / (burnRate * 4.33); // Convert weekly to monthly
    
    const breakEvenIndex = netCashFlow.findIndex(flow => flow > 0);
    const breakEvenMonth = breakEvenIndex !== -1 ? breakEvenIndex + 1 : null;
    
    const minBalanceIndex = runningBalance.indexOf(Math.min(...runningBalance));
    const cashMinimumDate = periodLabels[minBalanceIndex];
    
    return {
      periods: periodLabels,
      forecasts,
      statistics: {
        totalInflows,
        totalOutflows,
        netCashFlow,
        cumulativeCashFlow,
        runningBalance
      },
      confidence: {
        lower90,
        upper90,
        lower95,
        upper95
      },
      keyMetrics: {
        burnRate,
        runwayMonths,
        breakEvenMonth,
        cashMinimumDate
      }
    };
  }

  create52WeekForecast(
    categories: CashFlowCategory[],
    drivers: ForecastDriver[] = [],
    startDate: Date = new Date()
  ): RollingForecastResult {
    
    const periods = 52;
    const periodLabels = this.generatePeriodLabels(startDate, periods, 'weekly');
    const forecasts: { [category: string]: number[] } = {};
    
    // Generate forecasts for each category
    for (const category of categories) {
      const categoryForecasts: number[] = new Array(periods).fill(0);
      
      for (const input of category.subcategories) {
        // Get base projection with stronger trend modeling for longer period
        let projection = this.applyDriverBasedForecast(input, drivers, periods);
        
        // Apply seasonality with yearly pattern
        projection = projection.map((value, index) => 
          this.applySeasonality(value, input.seasonalityPattern, index)
        );
        
        // Add uncertainty for longer forecasts
        if (input.volatility) {
          projection = projection.map((value, index) => {
            const uncertaintyFactor = 1 + (Math.random() - 0.5) * 2 * input.volatility! * Math.sqrt(index + 1);
            return Math.max(0, value * uncertaintyFactor);
          });
        }
        
        // Add to category total
        for (let i = 0; i < periods; i++) {
          categoryForecasts[i] += projection[i];
        }
      }
      
      forecasts[category.name] = categoryForecasts;
    }

    // Calculate statistics (similar to 13-week but with quarterly aggregation)
    const totalInflows = new Array(periods).fill(0);
    const totalOutflows = new Array(periods).fill(0);
    
    for (const category of categories) {
      const categoryValues = forecasts[category.name];
      for (let i = 0; i < periods; i++) {
        if (category.isInflow) {
          totalInflows[i] += categoryValues[i];
        } else {
          totalOutflows[i] += categoryValues[i];
        }
      }
    }
    
    const netCashFlow = totalInflows.map((inflow, index) => inflow - totalOutflows[index]);
    const cumulativeCashFlow = netCashFlow.reduce((acc, flow, index) => {
      acc.push(index === 0 ? flow : acc[index - 1] + flow);
      return acc;
    }, [] as number[]);
    
    const initialCashBalance = 100000;
    const runningBalance = cumulativeCashFlow.map(cumulative => initialCashBalance + cumulative);
    
    // Wider confidence intervals for longer forecast
    const volatility = 0.25; // 25% standard deviation for yearly forecast
    const lower90 = netCashFlow.map((flow, index) => flow * (1 - 1.645 * volatility * Math.sqrt((index + 1) / 52)));
    const upper90 = netCashFlow.map((flow, index) => flow * (1 + 1.645 * volatility * Math.sqrt((index + 1) / 52)));
    const lower95 = netCashFlow.map((flow, index) => flow * (1 - 1.96 * volatility * Math.sqrt((index + 1) / 52)));
    const upper95 = netCashFlow.map((flow, index) => flow * (1 + 1.96 * volatility * Math.sqrt((index + 1) / 52)));
    
    // Calculate key metrics
    const avgNegativeCashFlow = netCashFlow.filter(flow => flow < 0).reduce((sum, flow) => sum + flow, 0) / 
                               Math.max(1, netCashFlow.filter(flow => flow < 0).length);
    const burnRate = Math.abs(avgNegativeCashFlow);
    const runwayMonths = initialCashBalance / (burnRate * 4.33);
    
    const breakEvenIndex = netCashFlow.findIndex(flow => flow > 0);
    const breakEvenMonth = breakEvenIndex !== -1 ? Math.ceil((breakEvenIndex + 1) / 4.33) : null;
    
    const minBalanceIndex = runningBalance.indexOf(Math.min(...runningBalance));
    const cashMinimumDate = periodLabels[minBalanceIndex];
    
    return {
      periods: periodLabels,
      forecasts,
      statistics: {
        totalInflows,
        totalOutflows,
        netCashFlow,
        cumulativeCashFlow,
        runningBalance
      },
      confidence: {
        lower90,
        upper90,
        lower95,
        upper95
      },
      keyMetrics: {
        burnRate,
        runwayMonths,
        breakEvenMonth,
        cashMinimumDate
      }
    };
  }

  generateForecastWorksheet(result: RollingForecastResult, title: string): Array<Array<CellValue | string | number>> {
    const worksheet: Array<Array<CellValue | string | number>> = [];
    
    // Header
    worksheet.push([title.toUpperCase(), '', '', '', '', '', '', '']);
    worksheet.push(['Generated: ' + new Date().toLocaleDateString(), '', '', '', '', '', '', '']);
    worksheet.push(['', '', '', '', '', '', '', '']);
    
    // Key Metrics Summary
    worksheet.push(['KEY CASH FLOW METRICS:', '', '', '', '', '', '', '']);
    worksheet.push(['Metric', 'Value', 'Unit', 'Interpretation', '', '', '', '']);
    worksheet.push(['Current Cash Burn Rate', result.keyMetrics.burnRate.toFixed(0), 'per period', 'Average negative cash flow']);
    worksheet.push(['Cash Runway', result.keyMetrics.runwayMonths.toFixed(1), 'months', 'Time until cash depletion']);
    
    if (result.keyMetrics.breakEvenMonth) {
      worksheet.push(['Break-even Period', result.keyMetrics.breakEvenMonth.toString(), 'period #', 'When cash flow turns positive']);
    }
    
    if (result.keyMetrics.cashMinimumDate) {
      worksheet.push(['Minimum Cash Date', result.keyMetrics.cashMinimumDate, 'date', 'Lowest projected cash balance']);
    }
    
    worksheet.push(['', '', '', '', '', '', '', '']);
    
    // Detailed Forecast Table
    worksheet.push(['DETAILED CASH FLOW FORECAST:', '', '', '', '', '', '', '']);
    
    // Header row with periods
    const headerRow = ['Period', ...result.periods.slice(0, Math.min(result.periods.length, 26))]; // Limit columns for Excel
    worksheet.push(headerRow);
    
    // Cash inflows by category
    worksheet.push(['CASH INFLOWS:', ...new Array(headerRow.length - 1).fill('')]);
    
    for (const [category, values] of Object.entries(result.forecasts)) {
      // Check if this is an inflow category (heuristic based on positive values)
      const isInflow = values.some(v => v > 0) && values.reduce((sum, v) => sum + v, 0) > 0;
      if (isInflow) {
        worksheet.push([category, ...values.slice(0, Math.min(values.length, 26)).map(v => v.toFixed(0))]);
      }
    }
    
    worksheet.push(['Total Inflows', ...result.statistics.totalInflows.slice(0, Math.min(result.statistics.totalInflows.length, 26)).map(v => v.toFixed(0))]);
    worksheet.push(['', '', '', '', '', '', '', '']);
    
    // Cash outflows by category
    worksheet.push(['CASH OUTFLOWS:', ...new Array(headerRow.length - 1).fill('')]);
    
    for (const [category, values] of Object.entries(result.forecasts)) {
      const isOutflow = values.some(v => v > 0) && !(values.some(v => v > 0) && values.reduce((sum, v) => sum + v, 0) > 0);
      if (isOutflow) {
        worksheet.push([category, ...values.slice(0, Math.min(values.length, 26)).map(v => v.toFixed(0))]);
      }
    }
    
    worksheet.push(['Total Outflows', ...result.statistics.totalOutflows.slice(0, Math.min(result.statistics.totalOutflows.length, 26)).map(v => v.toFixed(0))]);
    worksheet.push(['', '', '', '', '', '', '', '']);
    
    // Net cash flow and running balance
    worksheet.push(['NET CASH FLOW:', ...new Array(headerRow.length - 1).fill('')]);
    worksheet.push(['Net Flow', ...result.statistics.netCashFlow.slice(0, Math.min(result.statistics.netCashFlow.length, 26)).map(v => v.toFixed(0))]);
    worksheet.push(['Cumulative Flow', ...result.statistics.cumulativeCashFlow.slice(0, Math.min(result.statistics.cumulativeCashFlow.length, 26)).map(v => v.toFixed(0))]);
    worksheet.push(['Running Balance', ...result.statistics.runningBalance.slice(0, Math.min(result.statistics.runningBalance.length, 26)).map(v => v.toFixed(0))]);
    
    worksheet.push(['', '', '', '', '', '', '', '']);
    
    // Confidence intervals
    worksheet.push(['CONFIDENCE INTERVALS (Net Cash Flow):', ...new Array(headerRow.length - 1).fill('')]);
    worksheet.push(['90% Lower Bound', ...result.confidence.lower90.slice(0, Math.min(result.confidence.lower90.length, 26)).map(v => v.toFixed(0))]);
    worksheet.push(['90% Upper Bound', ...result.confidence.upper90.slice(0, Math.min(result.confidence.upper90.length, 26)).map(v => v.toFixed(0))]);
    worksheet.push(['95% Lower Bound', ...result.confidence.lower95.slice(0, Math.min(result.confidence.lower95.length, 26)).map(v => v.toFixed(0))]);
    worksheet.push(['95% Upper Bound', ...result.confidence.upper95.slice(0, Math.min(result.confidence.upper95.length, 26)).map(v => v.toFixed(0))]);
    
    worksheet.push(['', '', '', '', '', '', '', '']);
    
    // Risk analysis
    worksheet.push(['RISK ANALYSIS:', '', '', '', '', '', '', '']);
    const negativePeriods = result.statistics.netCashFlow.filter(flow => flow < 0).length;
    const totalPeriods = result.statistics.netCashFlow.length;
    const riskPercentage = (negativePeriods / totalPeriods * 100).toFixed(1);
    
    worksheet.push(['Periods with Negative Cash Flow', negativePeriods.toString(), `out of ${totalPeriods}`, `${riskPercentage}% of periods`]);
    
    const minBalance = Math.min(...result.statistics.runningBalance);
    const maxBalance = Math.max(...result.statistics.runningBalance);
    
    worksheet.push(['Minimum Projected Balance', minBalance.toFixed(0), 'currency', 'Lowest cash position']);
    worksheet.push(['Maximum Projected Balance', maxBalance.toFixed(0), 'currency', 'Highest cash position']);
    worksheet.push(['Balance Volatility', (maxBalance - minBalance).toFixed(0), 'currency', 'Range of cash positions']);
    
    worksheet.push(['', '', '', '', '', '', '', '']);
    worksheet.push(['ASSUMPTIONS & LIMITATIONS:', '', '', '', '', '', '', '']);
    worksheet.push(['- Forecast based on historical trends and specified growth rates', '', '', '', '', '', '', '']);
    worksheet.push(['- Seasonality patterns applied where specified', '', '', '', '', '', '', '']);
    worksheet.push(['- Confidence intervals are estimates based on assumed volatility', '', '', '', '', '', '', '']);
    worksheet.push(['- External factors (market changes, competition) not modeled', '', '', '', '', '', '', '']);
    worksheet.push(['- Review and update forecast regularly with actual results', '', '', '', '', '', '', '']);
    
    return worksheet;
  }
}