import { CellValue } from '../excel/excel-manager.js';

export interface MonteCarloVariable {
  name: string;
  distributionType: 'normal' | 'uniform' | 'triangular' | 'lognormal' | 'beta';
  parameters: {
    min?: number;
    max?: number;
    mean?: number;
    stdDev?: number;
    mode?: number; // For triangular distribution
    alpha?: number; // For beta distribution
    beta?: number;  // For beta distribution
  };
  correlations?: { [variableName: string]: number };
}

export interface MonteCarloResult {
  variable: string;
  iterations: number;
  results: number[];
  statistics: {
    mean: number;
    median: number;
    stdDev: number;
    var: number;
    min: number;
    max: number;
    percentiles: {
      p5: number;
      p10: number;
      p25: number;
      p75: number;
      p90: number;
      p95: number;
    };
    confidenceIntervals: {
      ci90: { lower: number; upper: number };
      ci95: { lower: number; upper: number };
      ci99: { lower: number; upper: number };
    };
  };
  histogram: { bucket: string; count: number; percentage: number }[];
}

export interface ScenarioDefinition {
  name: string;
  description: string;
  variables: MonteCarloVariable[];
  formula: string; // Excel-style formula using variable names
  iterations: number;
}

export class MonteCarloEngine {
  private random(): number {
    return Math.random();
  }

  private boxMullerTransform(): number {
    // Box-Muller transform for normal distribution
    const u = 0.001 + this.random() * 0.998; // Avoid 0 and 1
    const v = 0.001 + this.random() * 0.998;
    return Math.sqrt(-2.0 * Math.log(u)) * Math.cos(2.0 * Math.PI * v);
  }

  private sampleFromDistribution(variable: MonteCarloVariable): number {
    const { distributionType, parameters } = variable;

    switch (distributionType) {
      case 'uniform':
        const min = parameters.min ?? 0;
        const max = parameters.max ?? 1;
        return min + (max - min) * this.random();

      case 'normal':
        const mean = parameters.mean ?? 0;
        const stdDev = parameters.stdDev ?? 1;
        return mean + stdDev * this.boxMullerTransform();

      case 'lognormal':
        const logMean = parameters.mean ?? 0;
        const logStdDev = parameters.stdDev ?? 1;
        const normalSample = logMean + logStdDev * this.boxMullerTransform();
        return Math.exp(normalSample);

      case 'triangular':
        const triMin = parameters.min ?? 0;
        const triMax = parameters.max ?? 1;
        const mode = parameters.mode ?? (triMin + triMax) / 2;
        const u = this.random();
        
        const fc = (mode - triMin) / (triMax - triMin);
        if (u < fc) {
          return triMin + Math.sqrt(u * (triMax - triMin) * (mode - triMin));
        } else {
          return triMax - Math.sqrt((1 - u) * (triMax - triMin) * (triMax - mode));
        }

      case 'beta':
        const alpha = parameters.alpha ?? 2;
        const beta = parameters.beta ?? 5;
        // Using acceptance-rejection method for beta distribution
        let x: number, y: number;
        do {
          x = Math.pow(this.random(), 1 / alpha);
          y = Math.pow(this.random(), 1 / beta);
        } while (x + y > 1);
        return x / (x + y);

      default:
        throw new Error(`Unsupported distribution type: ${distributionType}`);
    }
  }

  private evaluateFormula(formula: string, variables: { [name: string]: number }): number {
    // Simple formula evaluator - in production, you'd want a more robust parser
    let expression = formula;
    
    // Replace variable names with values
    for (const [name, value] of Object.entries(variables)) {
      const regex = new RegExp(`\\b${name}\\b`, 'g');
      expression = expression.replace(regex, value.toString());
    }

    // Basic safety check and evaluation
    if (!/^[\d+\-*/(). ]+$/.test(expression)) {
      throw new Error(`Invalid formula characters: ${expression}`);
    }

    try {
      return Function(`"use strict"; return (${expression})`)();
    } catch (error) {
      throw new Error(`Formula evaluation error: ${error}`);
    }
  }

  private calculateStatistics(results: number[]): MonteCarloResult['statistics'] {
    const sorted = [...results].sort((a, b) => a - b);
    const n = results.length;
    
    const mean = results.reduce((sum, val) => sum + val, 0) / n;
    const variance = results.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) / (n - 1);
    const stdDev = Math.sqrt(variance);
    
    const getPercentile = (p: number): number => {
      const index = Math.ceil(n * p / 100) - 1;
      return sorted[Math.max(0, Math.min(index, n - 1))];
    };

    return {
      mean,
      median: getPercentile(50),
      stdDev,
      var: variance,
      min: sorted[0],
      max: sorted[n - 1],
      percentiles: {
        p5: getPercentile(5),
        p10: getPercentile(10),
        p25: getPercentile(25),
        p75: getPercentile(75),
        p90: getPercentile(90),
        p95: getPercentile(95)
      },
      confidenceIntervals: {
        ci90: { lower: getPercentile(5), upper: getPercentile(95) },
        ci95: { lower: getPercentile(2.5), upper: getPercentile(97.5) },
        ci99: { lower: getPercentile(0.5), upper: getPercentile(99.5) }
      }
    };
  }

  private createHistogram(results: number[], bins: number = 20): { bucket: string; count: number; percentage: number }[] {
    const min = Math.min(...results);
    const max = Math.max(...results);
    const binWidth = (max - min) / bins;
    
    const histogram: { bucket: string; count: number; percentage: number }[] = [];
    
    for (let i = 0; i < bins; i++) {
      const bucketMin = min + i * binWidth;
      const bucketMax = min + (i + 1) * binWidth;
      const count = results.filter(val => val >= bucketMin && val < bucketMax).length;
      
      histogram.push({
        bucket: `${bucketMin.toFixed(2)} - ${bucketMax.toFixed(2)}`,
        count,
        percentage: (count / results.length) * 100
      });
    }
    
    return histogram;
  }

  runSimulation(scenario: ScenarioDefinition): MonteCarloResult {
    const results: number[] = [];
    const { variables, formula, iterations } = scenario;

    for (let i = 0; i < iterations; i++) {
      const variableValues: { [name: string]: number } = {};
      
      // Sample from each variable's distribution
      for (const variable of variables) {
        variableValues[variable.name] = this.sampleFromDistribution(variable);
      }
      
      // Evaluate the target formula
      const result = this.evaluateFormula(formula, variableValues);
      results.push(result);
    }

    const statistics = this.calculateStatistics(results);
    const histogram = this.createHistogram(results);

    return {
      variable: scenario.name,
      iterations,
      results,
      statistics,
      histogram
    };
  }

  generateScenarioWorksheet(result: MonteCarloResult, scenario: ScenarioDefinition): Array<Array<CellValue | string | number>> {
    const worksheet: Array<Array<CellValue | string | number>> = [];
    
    // Header
    worksheet.push(['MONTE CARLO SIMULATION ANALYSIS', '', '', '', '', '']);
    worksheet.push([`Scenario: ${scenario.name}`, '', '', '', '', '']);
    worksheet.push([`Description: ${scenario.description}`, '', '', '', '', '']);
    worksheet.push([`Iterations: ${result.iterations}`, '', '', '', '', '']);
    worksheet.push([`Formula: ${scenario.formula}`, '', '', '', '', '']);
    worksheet.push(['', '', '', '', '', '']);
    
    // Variable Definitions
    worksheet.push(['INPUT VARIABLES:', '', '', '', '', '']);
    worksheet.push(['Variable', 'Distribution', 'Parameters', 'Description', '', '']);
    
    for (const variable of scenario.variables) {
      const paramStr = Object.entries(variable.parameters)
        .map(([key, value]) => `${key}: ${value}`)
        .join(', ');
      
      worksheet.push([
        variable.name,
        variable.distributionType,
        paramStr,
        `Monte Carlo input variable`,
        '', ''
      ]);
    }
    
    worksheet.push(['', '', '', '', '', '']);
    
    // Results Summary
    worksheet.push(['SIMULATION RESULTS:', '', '', '', '', '']);
    worksheet.push(['Statistic', 'Value', 'Formula', 'Description', '', '']);
    
    const stats = result.statistics;
    worksheet.push(['Mean', stats.mean, '=AVERAGE(results)', 'Expected value']);
    worksheet.push(['Median', stats.median, '=MEDIAN(results)', '50th percentile']);
    worksheet.push(['Std Deviation', stats.stdDev, '=STDEV.S(results)', 'Volatility measure']);
    worksheet.push(['Variance', stats.var, '=VAR.S(results)', 'Risk measure']);
    worksheet.push(['Minimum', stats.min, '=MIN(results)', 'Best case observed']);
    worksheet.push(['Maximum', stats.max, '=MAX(results)', 'Worst case observed']);
    
    worksheet.push(['', '', '', '', '', '']);
    
    // Percentiles
    worksheet.push(['RISK ANALYSIS:', '', '', '', '', '']);
    worksheet.push(['Percentile', 'Value', 'Probability', 'Risk Assessment', '', '']);
    
    worksheet.push(['5th Percentile', stats.percentiles.p5, '5%', 'Extreme downside risk']);
    worksheet.push(['10th Percentile', stats.percentiles.p10, '10%', 'High downside risk']);
    worksheet.push(['25th Percentile', stats.percentiles.p25, '25%', 'Moderate downside risk']);
    worksheet.push(['75th Percentile', stats.percentiles.p75, '75%', 'Moderate upside potential']);
    worksheet.push(['90th Percentile', stats.percentiles.p90, '90%', 'High upside potential']);
    worksheet.push(['95th Percentile', stats.percentiles.p95, '95%', 'Extreme upside potential']);
    
    worksheet.push(['', '', '', '', '', '']);
    
    // Confidence Intervals
    worksheet.push(['CONFIDENCE INTERVALS:', '', '', '', '', '']);
    worksheet.push(['Confidence Level', 'Lower Bound', 'Upper Bound', 'Width', 'Interpretation', '']);
    
    const ci90 = stats.confidenceIntervals.ci90;
    const ci95 = stats.confidenceIntervals.ci95;
    const ci99 = stats.confidenceIntervals.ci99;
    
    worksheet.push(['90%', ci90.lower, ci90.upper, ci90.upper - ci90.lower, '90% chance result falls in this range']);
    worksheet.push(['95%', ci95.lower, ci95.upper, ci95.upper - ci95.lower, '95% chance result falls in this range']);
    worksheet.push(['99%', ci99.lower, ci99.upper, ci99.upper - ci99.lower, '99% chance result falls in this range']);
    
    worksheet.push(['', '', '', '', '', '']);
    
    // Histogram Data
    worksheet.push(['FREQUENCY DISTRIBUTION:', '', '', '', '', '']);
    worksheet.push(['Range', 'Frequency', 'Percentage', 'Cumulative %', '', '']);
    
    let cumulative = 0;
    for (const bucket of result.histogram) {
      cumulative += bucket.percentage;
      worksheet.push([
        bucket.bucket,
        bucket.count,
        bucket.percentage.toFixed(2) + '%',
        cumulative.toFixed(2) + '%',
        '', ''
      ]);
    }
    
    worksheet.push(['', '', '', '', '', '']);
    
    // Risk Metrics
    worksheet.push(['RISK METRICS:', '', '', '', '', '']);
    worksheet.push(['Metric', 'Value', 'Calculation', 'Interpretation', '', '']);
    
    const downside = result.results.filter(r => r < stats.mean);
    const downsideDeviation = downside.length > 0 
      ? Math.sqrt(downside.reduce((sum, val) => sum + Math.pow(val - stats.mean, 2), 0) / downside.length)
      : 0;
    
    const valueAtRisk95 = stats.percentiles.p5; // 95% VaR (5th percentile)
    const expectedShortfall = downside.filter(r => r <= valueAtRisk95).reduce((sum, val) => sum + val, 0) / 
                              Math.max(1, downside.filter(r => r <= valueAtRisk95).length);
    
    worksheet.push(['Value at Risk (95%)', valueAtRisk95, '5th percentile', 'Maximum expected loss at 95% confidence']);
    worksheet.push(['Expected Shortfall', expectedShortfall || 0, 'Average of worst 5%', 'Expected loss given VaR breach']);
    worksheet.push(['Downside Deviation', downsideDeviation, 'Std dev of below-mean results', 'Downside volatility measure']);
    worksheet.push(['Coefficient of Variation', stats.stdDev / Math.abs(stats.mean), 'StdDev / |Mean|', 'Risk per unit of return']);
    
    worksheet.push(['', '', '', '', '', '']);
    worksheet.push(['PROFESSIONAL REFERENCES:', '', '', '', '', '']);
    worksheet.push(['- Basel III Risk Management Guidelines', '', '', '', '', '']);
    worksheet.push(['- CFA Institute Risk Management Standards', '', '', '', '', '']);
    worksheet.push(['- COSO Enterprise Risk Management Framework', '', '', '', '', '']);
    
    return worksheet;
  }

  createSensitivityAnalysis(baseScenario: ScenarioDefinition, sensitivityRange: number = 0.2): Array<Array<CellValue | string | number>> {
    const worksheet: Array<Array<CellValue | string | number>> = [];
    const results: { [variableName: string]: { change: number; result: number }[] } = {};
    
    // Header
    worksheet.push(['SENSITIVITY ANALYSIS', '', '', '', '', '']);
    worksheet.push([`Base Scenario: ${baseScenario.name}`, '', '', '', '', '']);
    worksheet.push(['', '', '', '', '', '']);
    
    // Run base case
    const baseResult = this.runSimulation(baseScenario);
    const baseValue = baseResult.statistics.mean;
    
    worksheet.push(['BASE CASE RESULT:', baseValue, '', '', '', '']);
    worksheet.push(['', '', '', '', '', '']);
    
    // Sensitivity analysis for each variable
    for (const variable of baseScenario.variables) {
      const sensitivityResults: { change: number; result: number }[] = [];
      
      for (let changePercent = -sensitivityRange; changePercent <= sensitivityRange; changePercent += 0.05) {
        const modifiedScenario = { ...baseScenario };
        const modifiedVariable = { ...variable };
        
        // Adjust the variable's parameters
        if (modifiedVariable.parameters.mean !== undefined) {
          modifiedVariable.parameters.mean *= (1 + changePercent);
        }
        if (modifiedVariable.parameters.min !== undefined && modifiedVariable.parameters.max !== undefined) {
          const center = (modifiedVariable.parameters.min + modifiedVariable.parameters.max) / 2;
          const range = modifiedVariable.parameters.max - modifiedVariable.parameters.min;
          modifiedVariable.parameters.min = center - range / 2 * (1 + changePercent);
          modifiedVariable.parameters.max = center + range / 2 * (1 + changePercent);
        }
        
        modifiedScenario.variables = baseScenario.variables.map(v => 
          v.name === variable.name ? modifiedVariable : v
        );
        
        const result = this.runSimulation(modifiedScenario);
        sensitivityResults.push({
          change: changePercent,
          result: result.statistics.mean
        });
      }
      
      results[variable.name] = sensitivityResults;
    }
    
    // Output sensitivity table
    worksheet.push(['SENSITIVITY TO INPUT VARIABLES:', '', '', '', '', '']);
    worksheet.push(['Variable Change (%)', ...baseScenario.variables.map(v => v.name)]);
    
    const changes = results[baseScenario.variables[0].name].map(r => r.change);
    
    for (let i = 0; i < changes.length; i++) {
      const row = [
        (changes[i] * 100).toFixed(1) + '%',
        ...baseScenario.variables.map(variable => {
          const result = results[variable.name][i].result;
          const impact = ((result - baseValue) / baseValue * 100).toFixed(2) + '%';
          return impact;
        })
      ];
      worksheet.push(row);
    }
    
    return worksheet;
  }
}