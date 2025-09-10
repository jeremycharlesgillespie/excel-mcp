import { Tool, ToolResult } from '../types/index.js';
import { ExcelManager } from '../excel/excel-manager.js';
import { MonteCarloEngine, ScenarioDefinition } from '../analytics/monte-carlo.js';
import { RollingForecastEngine, CashFlowCategory, ForecastDriver, ForecastInput } from '../forecasting/rolling-forecast.js';
import { DCFBuilder, DCFAssumptions } from '../modeling/dcf-builder.js';
import { ExecutiveDashboard, DashboardData, KPIMetric, RiskIndicator } from '../dashboard/executive-dashboard.js';

const excelManager = new ExcelManager();
const monteCarloEngine = new MonteCarloEngine();
const forecastEngine = new RollingForecastEngine();
const dcfBuilder = new DCFBuilder();
const dashboard = new ExecutiveDashboard();

export const analyticsTools: Tool[] = [
  {
    name: "analytics_monte_carlo_simulation",
    description: "Run Monte Carlo simulation for risk analysis and scenario planning",
    inputSchema: {
      type: "object",
      properties: {
        scenarioName: { type: "string", description: "Name of the scenario being analyzed" },
        description: { type: "string", description: "Description of what this simulation models" },
        formula: { type: "string", description: "Formula using variable names (e.g., 'revenue - costs - taxes')" },
        iterations: { type: "number", default: 10000, description: "Number of simulation iterations" },
        variables: {
          type: "array",
          items: {
            type: "object",
            properties: {
              name: { type: "string", description: "Variable name to use in formula" },
              distributionType: { 
                type: "string", 
                enum: ["normal", "uniform", "triangular", "lognormal", "beta"],
                description: "Probability distribution type"
              },
              parameters: {
                type: "object",
                description: "Distribution parameters",
                properties: {
                  min: { type: "number" },
                  max: { type: "number" },
                  mean: { type: "number" },
                  stdDev: { type: "number" },
                  mode: { type: "number" },
                  alpha: { type: "number" },
                  beta: { type: "number" }
                }
              }
            },
            required: ["name", "distributionType", "parameters"]
          }
        },
        worksheetName: { type: "string", default: "Monte Carlo Analysis" }
      },
      required: ["scenarioName", "formula", "variables"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const scenario: ScenarioDefinition = {
          name: args.scenarioName,
          description: args.description || `Monte Carlo simulation for ${args.scenarioName}`,
          formula: args.formula,
          iterations: args.iterations || 10000,
          variables: args.variables
        };
        
        // Run the simulation
        const result = monteCarloEngine.runSimulation(scenario);
        
        // Generate Excel worksheet
        const worksheetData = monteCarloEngine.generateScenarioWorksheet(result, scenario);
        const worksheetName = args.worksheetName || "Monte Carlo Analysis";
        
        await excelManager.addWorksheet(worksheetName);
        await excelManager.writeWorksheet(worksheetName, worksheetData);
        await excelManager.autoFitColumnWidths(worksheetName, { minWidth: 80, maxWidth: 300 });
        
        return {
          success: true,
          message: `Completed Monte Carlo simulation: ${args.scenarioName} (${result.iterations} iterations)`,
          data: {
            worksheetName,
            statistics: result.statistics,
            riskMetrics: {
              valueAtRisk95: result.statistics.percentiles.p5,
              expectedReturn: result.statistics.mean,
              volatility: result.statistics.stdDev,
              probabilityOfLoss: (result.results.filter(r => r < 0).length / result.results.length * 100).toFixed(2) + '%'
            }
          }
        };
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "analytics_sensitivity_analysis",
    description: "Perform sensitivity analysis to understand variable impact on outcomes",
    inputSchema: {
      type: "object",
      properties: {
        scenarioName: { type: "string" },
        description: { type: "string" },
        formula: { type: "string" },
        variables: {
          type: "array",
          items: {
            type: "object",
            properties: {
              name: { type: "string" },
              distributionType: { 
                type: "string", 
                enum: ["normal", "uniform", "triangular", "lognormal", "beta"]
              },
              parameters: { type: "object" }
            },
            required: ["name", "distributionType", "parameters"]
          }
        },
        sensitivityRange: { 
          type: "number", 
          default: 0.2, 
          description: "Range for sensitivity analysis (+/- percentage, e.g., 0.2 for ±20%)"
        },
        worksheetName: { type: "string", default: "Sensitivity Analysis" }
      },
      required: ["scenarioName", "formula", "variables"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const scenario: ScenarioDefinition = {
          name: args.scenarioName,
          description: args.description || `Sensitivity analysis for ${args.scenarioName}`,
          formula: args.formula,
          iterations: 1000, // Fewer iterations for sensitivity analysis
          variables: args.variables
        };
        
        // Create sensitivity analysis
        const worksheetData = monteCarloEngine.createSensitivityAnalysis(
          scenario, 
          args.sensitivityRange || 0.2
        );
        
        const worksheetName = args.worksheetName || "Sensitivity Analysis";
        
        await excelManager.addWorksheet(worksheetName);
        await excelManager.writeWorksheet(worksheetName, worksheetData);
        await excelManager.autoFitColumnWidths(worksheetName, { minWidth: 80, maxWidth: 200 });
        
        return {
          success: true,
          message: `Completed sensitivity analysis for ${args.scenarioName}`,
          data: {
            worksheetName,
            sensitivityRange: args.sensitivityRange || 0.2,
            variableCount: args.variables.length
          }
        };
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "analytics_scenario_comparison",
    description: "Compare multiple scenarios (Best/Base/Worst case) with probability weighting",
    inputSchema: {
      type: "object",
      properties: {
        scenarios: {
          type: "array",
          items: {
            type: "object",
            properties: {
              name: { type: "string", description: "Scenario name (e.g., 'Best Case', 'Base Case', 'Worst Case')" },
              probability: { type: "number", description: "Probability weight (0-1, should sum to 1)" },
              description: { type: "string" },
              formula: { type: "string" },
              variables: {
                type: "array",
                items: {
                  type: "object",
                  properties: {
                    name: { type: "string" },
                    distributionType: { type: "string", enum: ["normal", "uniform", "triangular", "lognormal", "beta"] },
                    parameters: { type: "object" }
                  },
                  required: ["name", "distributionType", "parameters"]
                }
              }
            },
            required: ["name", "probability", "formula", "variables"]
          }
        },
        worksheetName: { type: "string", default: "Scenario Comparison" }
      },
      required: ["scenarios"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const results = [];
        
        // Run simulation for each scenario
        for (const scenarioData of args.scenarios) {
          const scenario: ScenarioDefinition = {
            name: scenarioData.name,
            description: scenarioData.description || `Scenario: ${scenarioData.name}`,
            formula: scenarioData.formula,
            iterations: 5000,
            variables: scenarioData.variables
          };
          
          const result = monteCarloEngine.runSimulation(scenario);
          results.push({
            ...result,
            probability: scenarioData.probability,
            scenarioName: scenarioData.name
          });
        }
        
        // Calculate probability-weighted outcomes
        const weightedMean = results.reduce((sum, r) => sum + r.statistics.mean * r.probability, 0);
        const weightedStdDev = Math.sqrt(
          results.reduce((sum, r) => 
            sum + r.probability * (r.statistics.var + Math.pow(r.statistics.mean - weightedMean, 2)), 0)
        );
        
        // Create comparison worksheet
        const worksheet: Array<Array<any>> = [];
        
        worksheet.push(['SCENARIO COMPARISON ANALYSIS', '', '', '', '', '', '']);
        worksheet.push(['Date: ' + new Date().toLocaleDateString(), '', '', '', '', '', '']);
        worksheet.push(['', '', '', '', '', '', '']);
        
        // Summary table
        worksheet.push(['SCENARIO SUMMARY:', '', '', '', '', '', '']);
        worksheet.push(['Scenario', 'Probability', 'Expected Value', 'Std Deviation', 'Best Case', 'Worst Case', 'VaR (95%)']);
        
        for (const result of results) {
          worksheet.push([
            result.scenarioName,
            (result.probability * 100).toFixed(1) + '%',
            result.statistics.mean.toFixed(2),
            result.statistics.stdDev.toFixed(2),
            result.statistics.max.toFixed(2),
            result.statistics.min.toFixed(2),
            result.statistics.percentiles.p5.toFixed(2)
          ]);
        }
        
        worksheet.push(['', '', '', '', '', '', '']);
        
        // Weighted outcomes
        worksheet.push(['PROBABILITY-WEIGHTED ANALYSIS:', '', '', '', '', '', '']);
        worksheet.push(['Metric', 'Value', 'Calculation', 'Interpretation', '', '', '']);
        worksheet.push(['Expected Outcome', weightedMean.toFixed(2), 'Σ(Probability × Mean)', 'Probability-weighted expected value']);
        worksheet.push(['Combined Risk', weightedStdDev.toFixed(2), 'Portfolio volatility formula', 'Overall uncertainty measure']);
        
        const bestCase = Math.max(...results.map(r => r.statistics.max));
        const worstCase = Math.min(...results.map(r => r.statistics.min));
        
        worksheet.push(['Range Width', (bestCase - worstCase).toFixed(2), 'Best Case - Worst Case', 'Total outcome range']);
        worksheet.push(['Risk-Adjusted Return', (weightedMean / weightedStdDev).toFixed(2), 'Expected Return / Risk', 'Return per unit of risk']);
        
        worksheet.push(['', '', '', '', '', '', '']);
        
        // Decision matrix
        worksheet.push(['DECISION SUPPORT MATRIX:', '', '', '', '', '', '']);
        worksheet.push(['Scenario', 'Probability of Success', 'Risk Level', 'Recommendation', '', '', '']);
        
        for (const result of results) {
          const successProb = (result.results.filter(r => r > 0).length / result.results.length * 100).toFixed(1);
          const riskLevel = result.statistics.stdDev > weightedStdDev * 1.5 ? 'High' : 
                           result.statistics.stdDev > weightedStdDev * 0.8 ? 'Medium' : 'Low';
          const recommendation = result.statistics.mean > weightedMean && riskLevel !== 'High' ? 
                               'Favorable' : result.statistics.mean < weightedMean ? 'Unfavorable' : 'Monitor';
          
          worksheet.push([
            result.scenarioName,
            successProb + '%',
            riskLevel,
            recommendation,
            '', '', ''
          ]);
        }
        
        worksheet.push(['', '', '', '', '', '', '']);
        worksheet.push(['EXECUTIVE SUMMARY:', '', '', '', '', '', '']);
        worksheet.push([`Expected probability-weighted outcome: ${weightedMean.toFixed(2)}`, '', '', '', '', '', '']);
        worksheet.push([`Overall risk level: ${weightedStdDev.toFixed(2)} standard deviation`, '', '', '', '', '', '']);
        worksheet.push([`Outcome range: ${worstCase.toFixed(2)} to ${bestCase.toFixed(2)}`, '', '', '', '', '', '']);
        
        const worksheetName = args.worksheetName || "Scenario Comparison";
        
        await excelManager.addWorksheet(worksheetName);
        await excelManager.writeWorksheet(worksheetName, worksheet);
        await excelManager.autoFitColumnWidths(worksheetName, { minWidth: 100, maxWidth: 250 });
        
        return {
          success: true,
          message: `Completed scenario comparison analysis for ${results.length} scenarios`,
          data: {
            worksheetName,
            weightedExpectedValue: weightedMean,
            combinedRisk: weightedStdDev,
            scenarioCount: results.length,
            outcomeRange: { min: worstCase, max: bestCase }
          }
        };
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "analytics_13_week_forecast",
    description: "Generate 13-week rolling cash flow forecast with confidence intervals and risk analysis",
    inputSchema: {
      type: "object",
      properties: {
        categories: {
          type: "array",
          items: {
            type: "object",
            properties: {
              name: { type: "string", description: "Category name (e.g., 'Revenue', 'Operating Expenses')" },
              isInflow: { type: "boolean", description: "true for revenue/income, false for expenses" },
              subcategories: {
                type: "array",
                items: {
                  type: "object",
                  properties: {
                    category: { type: "string", description: "Subcategory name" },
                    historical: { 
                      type: "array", 
                      items: { type: "number" },
                      description: "Historical weekly values (last 12+ weeks recommended)"
                    },
                    growthRate: { type: "number", description: "Annual growth rate (decimal, e.g., 0.05 for 5%)" },
                    seasonalityPattern: {
                      type: "array",
                      items: { type: "number" },
                      description: "Weekly seasonal multipliers (52 values for full year pattern)"
                    },
                    volatility: { type: "number", description: "Standard deviation for uncertainty (decimal)" },
                    driver: { type: "string", description: "Name of driver variable if applicable" },
                    driverMultiplier: { type: "number", description: "Multiplier for driver-based forecasting" }
                  },
                  required: ["category", "historical"]
                }
              }
            },
            required: ["name", "isInflow", "subcategories"]
          }
        },
        drivers: {
          type: "array",
          items: {
            type: "object",
            properties: {
              name: { type: "string" },
              values: { 
                type: "array", 
                items: { type: "number" },
                description: "Future driver values for 13 weeks"
              },
              description: { type: "string" }
            },
            required: ["name", "values", "description"]
          },
          default: []
        },
        startDate: { type: "string", description: "Start date in YYYY-MM-DD format", default: "today" },
        worksheetName: { type: "string", default: "13-Week Forecast" }
      },
      required: ["categories"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const startDate = args.startDate && args.startDate !== "today" 
          ? new Date(args.startDate) 
          : new Date();
        
        // Convert input format to internal format
        const categories: CashFlowCategory[] = args.categories.map((cat: any) => ({
          name: cat.name,
          isInflow: cat.isInflow,
          subcategories: cat.subcategories.map((sub: any): ForecastInput => ({
            category: sub.category,
            subcategory: sub.subcategory,
            historical: sub.historical,
            seasonalityPattern: sub.seasonalityPattern,
            growthRate: sub.growthRate,
            volatility: sub.volatility,
            driver: sub.driver,
            driverMultiplier: sub.driverMultiplier
          }))
        }));
        
        const drivers: ForecastDriver[] = args.drivers || [];
        
        // Generate forecast
        const result = forecastEngine.create13WeekForecast(categories, drivers, startDate);
        
        // Create Excel worksheet
        const worksheetData = forecastEngine.generateForecastWorksheet(result, "13-Week Cash Flow Forecast");
        const worksheetName = args.worksheetName || "13-Week Forecast";
        
        await excelManager.addWorksheet(worksheetName);
        await excelManager.writeWorksheet(worksheetName, worksheetData);
        await excelManager.autoFitColumnWidths(worksheetName, { minWidth: 80, maxWidth: 150 });
        
        return {
          success: true,
          message: `Generated 13-week cash flow forecast`,
          data: {
            worksheetName,
            keyMetrics: result.keyMetrics,
            forecastSummary: {
              periods: result.periods.length,
              totalInflow: result.statistics.totalInflows.reduce((sum, val) => sum + val, 0).toFixed(0),
              totalOutflow: result.statistics.totalOutflows.reduce((sum, val) => sum + val, 0).toFixed(0),
              netCashFlow: result.statistics.netCashFlow.reduce((sum, val) => sum + val, 0).toFixed(0)
            }
          }
        };
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "analytics_52_week_forecast",
    description: "Generate 52-week rolling cash flow forecast with advanced trend modeling and uncertainty analysis",
    inputSchema: {
      type: "object",
      properties: {
        categories: {
          type: "array",
          items: {
            type: "object",
            properties: {
              name: { type: "string" },
              isInflow: { type: "boolean" },
              subcategories: {
                type: "array",
                items: {
                  type: "object",
                  properties: {
                    category: { type: "string" },
                    historical: { 
                      type: "array", 
                      items: { type: "number" },
                      description: "Historical weekly values (minimum 26 weeks recommended)"
                    },
                    growthRate: { type: "number" },
                    seasonalityPattern: { type: "array", items: { type: "number" } },
                    volatility: { type: "number" },
                    driver: { type: "string" },
                    driverMultiplier: { type: "number" }
                  },
                  required: ["category", "historical"]
                }
              }
            },
            required: ["name", "isInflow", "subcategories"]
          }
        },
        drivers: {
          type: "array",
          items: {
            type: "object",
            properties: {
              name: { type: "string" },
              values: { 
                type: "array", 
                items: { type: "number" },
                description: "Future driver values for 52 weeks"
              },
              description: { type: "string" }
            },
            required: ["name", "values", "description"]
          },
          default: []
        },
        startDate: { type: "string", default: "today" },
        worksheetName: { type: "string", default: "52-Week Forecast" }
      },
      required: ["categories"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const startDate = args.startDate && args.startDate !== "today" 
          ? new Date(args.startDate) 
          : new Date();
        
        const categories: CashFlowCategory[] = args.categories.map((cat: any) => ({
          name: cat.name,
          isInflow: cat.isInflow,
          subcategories: cat.subcategories.map((sub: any): ForecastInput => ({
            category: sub.category,
            subcategory: sub.subcategory,
            historical: sub.historical,
            seasonalityPattern: sub.seasonalityPattern,
            growthRate: sub.growthRate,
            volatility: sub.volatility,
            driver: sub.driver,
            driverMultiplier: sub.driverMultiplier
          }))
        }));
        
        const drivers: ForecastDriver[] = args.drivers || [];
        
        // Generate 52-week forecast
        const result = forecastEngine.create52WeekForecast(categories, drivers, startDate);
        
        // Create Excel worksheet
        const worksheetData = forecastEngine.generateForecastWorksheet(result, "52-Week Cash Flow Forecast");
        const worksheetName = args.worksheetName || "52-Week Forecast";
        
        await excelManager.addWorksheet(worksheetName);
        await excelManager.writeWorksheet(worksheetName, worksheetData);
        await excelManager.autoFitColumnWidths(worksheetName, { minWidth: 80, maxWidth: 150 });
        
        return {
          success: true,
          message: `Generated 52-week cash flow forecast`,
          data: {
            worksheetName,
            keyMetrics: result.keyMetrics,
            forecastSummary: {
              periods: result.periods.length,
              totalInflow: result.statistics.totalInflows.reduce((sum, val) => sum + val, 0).toFixed(0),
              totalOutflow: result.statistics.totalOutflows.reduce((sum, val) => sum + val, 0).toFixed(0),
              netCashFlow: result.statistics.netCashFlow.reduce((sum, val) => sum + val, 0).toFixed(0),
              yearEndBalance: result.statistics.runningBalance[result.statistics.runningBalance.length - 1].toFixed(0)
            }
          }
        };
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "analytics_dcf_valuation",
    description: "Build comprehensive DCF (Discounted Cash Flow) valuation model with sensitivity analysis",
    inputSchema: {
      type: "object",
      properties: {
        projectionYears: { type: "number", default: 5, description: "Number of explicit forecast years" },
        discountRate: { type: "number", description: "WACC - Weighted Average Cost of Capital (decimal, e.g., 0.12 for 12%)" },
        terminalGrowthRate: { type: "number", default: 0.025, description: "Long-term growth rate (decimal, e.g., 0.025 for 2.5%)" },
        initialRevenue: { type: "number", description: "Base year revenue in currency units" },
        revenueCAGR: { type: "number", description: "Revenue Compound Annual Growth Rate (decimal, e.g., 0.15 for 15%)" },
        terminalMargin: { type: "number", default: 0.20, description: "Terminal operating margin (decimal, e.g., 0.20 for 20%)" },
        terminalCapexRate: { type: "number", default: 0.03, description: "Terminal capex as % of revenue (decimal)" },
        terminalTaxRate: { type: "number", default: 0.25, description: "Terminal tax rate (decimal, e.g., 0.25 for 25%)" },
        worksheetName: { type: "string", default: "DCF Valuation" }
      },
      required: ["discountRate", "initialRevenue", "revenueCAGR"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const assumptions: DCFAssumptions = {
          projectionYears: args.projectionYears || 5,
          discountRate: args.discountRate,
          terminalGrowthRate: args.terminalGrowthRate || 0.025,
          initialRevenue: args.initialRevenue,
          revenueCAGR: args.revenueCAGR,
          terminalMargin: args.terminalMargin || 0.20,
          terminalCapexRate: args.terminalCapexRate || 0.03,
          terminalTaxRate: args.terminalTaxRate || 0.25
        };
        
        // Build DCF model
        const valuation = dcfBuilder.buildDCFModel(assumptions);
        
        // Create Excel worksheet
        const worksheetData = dcfBuilder.generateDCFWorksheet(valuation, assumptions);
        const worksheetName = args.worksheetName || "DCF Valuation";
        
        await excelManager.addWorksheet(worksheetName);
        await excelManager.writeWorksheet(worksheetName, worksheetData);
        await excelManager.autoFitColumnWidths(worksheetName, { minWidth: 90, maxWidth: 200 });
        
        return {
          success: true,
          message: `Generated DCF valuation model with ${assumptions.projectionYears}-year projections`,
          data: {
            worksheetName,
            valuation: {
              enterpriseValue: (valuation.enterpriseValue / 1000000).toFixed(1) + "M",
              terminalValuePercent: ((valuation.terminalValuePV / valuation.enterpriseValue) * 100).toFixed(1) + "%",
              evRevenueMultiple: valuation.keyMetrics.impliedEVRevenue.toFixed(1) + "x",
              evEbitdaMultiple: valuation.keyMetrics.impliedEVEBITDA.toFixed(1) + "x",
              breakEvenWACC: (valuation.keyMetrics.breakEvenWACC * 100).toFixed(2) + "%"
            },
            assumptions: {
              projectionYears: assumptions.projectionYears,
              wacc: (assumptions.discountRate * 100).toFixed(2) + "%",
              terminalGrowth: (assumptions.terminalGrowthRate * 100).toFixed(2) + "%",
              revenueCAGR: (assumptions.revenueCAGR * 100).toFixed(1) + "%"
            }
          }
        };
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "analytics_executive_dashboard",
    description: "Generate comprehensive executive dashboard with automated insights and KPI tracking",
    inputSchema: {
      type: "object",
      properties: {
        reportDate: { type: "string", description: "Report date in YYYY-MM-DD format", default: "today" },
        kpis: {
          type: "array",
          items: {
            type: "object",
            properties: {
              name: { type: "string" },
              current: { type: "number" },
              target: { type: "number" },
              previous: { type: "number" },
              unit: { type: "string" },
              category: { type: "string", enum: ["financial", "operational", "strategic", "risk"] },
              frequency: { type: "string", enum: ["daily", "weekly", "monthly", "quarterly"] },
              threshold: {
                type: "object",
                properties: {
                  excellent: { type: "number" },
                  good: { type: "number" },
                  warning: { type: "number" },
                  critical: { type: "number" }
                },
                required: ["excellent", "good", "warning", "critical"]
              },
              trend: { type: "string", enum: ["increasing", "decreasing", "stable"] },
              importance: { type: "string", enum: ["high", "medium", "low"] }
            },
            required: ["name", "current", "target", "previous", "unit", "category", "threshold", "trend", "importance"]
          }
        },
        risks: {
          type: "array",
          items: {
            type: "object",
            properties: {
              name: { type: "string" },
              score: { type: "number", minimum: 0, maximum: 100 },
              level: { type: "string", enum: ["low", "medium", "high", "critical"] },
              category: { type: "string", enum: ["liquidity", "operational", "strategic", "market", "credit"] },
              mitigation: { type: "string" },
              trend: { type: "string", enum: ["improving", "stable", "deteriorating"] },
              impact: { type: "number", description: "Financial impact estimate" }
            },
            required: ["name", "score", "level", "category", "mitigation", "trend", "impact"]
          }
        },
        cashFlow: {
          type: "object",
          properties: {
            current: { type: "number", description: "Current cash position" },
            projected13Week: { type: "number", description: "13-week net cash flow projection" },
            projected52Week: { type: "number", description: "52-week net cash flow projection" },
            burnRate: { type: "number", description: "Monthly cash burn rate" },
            runwayMonths: { type: "number", description: "Cash runway in months" }
          },
          required: ["current", "projected13Week", "projected52Week", "burnRate", "runwayMonths"]
        },
        financial: {
          type: "object",
          properties: {
            revenue: {
              type: "object",
              properties: {
                current: { type: "number" },
                target: { type: "number" },
                variance: { type: "number", description: "Variance percentage" }
              },
              required: ["current", "target", "variance"]
            },
            ebitda: {
              type: "object",
              properties: {
                current: { type: "number" },
                target: { type: "number" },
                margin: { type: "number", description: "EBITDA margin percentage" }
              },
              required: ["current", "target", "margin"]
            },
            workingCapital: {
              type: "object",
              properties: {
                current: { type: "number" },
                target: { type: "number" },
                days: { type: "number", description: "Days in working capital cycle" }
              },
              required: ["current", "target", "days"]
            },
            debtToEquity: {
              type: "object",
              properties: {
                current: { type: "number" },
                target: { type: "number" },
                trend: { type: "string" }
              },
              required: ["current", "target", "trend"]
            }
          },
          required: ["revenue", "ebitda", "workingCapital", "debtToEquity"]
        },
        worksheetName: { type: "string", default: "Executive Dashboard" }
      },
      required: ["kpis", "cashFlow", "financial"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const reportDate = args.reportDate && args.reportDate !== "today" 
          ? new Date(args.reportDate) 
          : new Date();
        
        const dashboardData: DashboardData = {
          reportDate,
          kpis: args.kpis,
          drivers: [], // Could be expanded in future
          risks: args.risks || [],
          cashFlow: args.cashFlow,
          financial: args.financial,
          insights: [] // Will be auto-generated
        };
        
        // Generate executive dashboard
        const worksheetData = dashboard.generateExecutiveDashboard(dashboardData);
        const worksheetName = args.worksheetName || "Executive Dashboard";
        
        await excelManager.addWorksheet(worksheetName);
        await excelManager.writeWorksheet(worksheetName, worksheetData);
        await excelManager.autoFitColumnWidths(worksheetName, { minWidth: 100, maxWidth: 300 });
        
        // Calculate summary metrics for response
        const criticalKPIs = args.kpis.filter((kpi: KPIMetric) => {
          return kpi.current < kpi.threshold.warning;
        });
        
        const highRisks = (args.risks || []).filter((r: RiskIndicator) => r.level === 'high' || r.level === 'critical');
        const overallHealthScore = Math.max(0, 100 - (criticalKPIs.length / args.kpis.length * 40) - (highRisks.length * 15));
        
        return {
          success: true,
          message: `Generated executive dashboard with ${args.kpis.length} KPIs and automated insights`,
          data: {
            worksheetName,
            summary: {
              overallHealthScore: overallHealthScore.toFixed(0) + '/100',
              criticalKPIs: criticalKPIs.length + '/' + args.kpis.length,
              highRisks: highRisks.length + '/' + (args.risks?.length || 0),
              cashRunway: args.cashFlow.runwayMonths.toFixed(1) + ' months',
              revenueVariance: args.financial.revenue.variance.toFixed(1) + '%'
            },
            insights: {
              cashFlowAlert: args.cashFlow.runwayMonths < 6,
              revenueAlert: args.financial.revenue.variance < -10,
              riskAlert: highRisks.length > 0,
              kpiAlert: criticalKPIs.length > 0
            }
          }
        };
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "analytics_kpi_dashboard",
    description: "Generate focused KPI performance dashboard with status tracking and trend analysis",
    inputSchema: {
      type: "object",
      properties: {
        kpis: {
          type: "array",
          items: {
            type: "object",
            properties: {
              name: { type: "string" },
              current: { type: "number" },
              target: { type: "number" },
              previous: { type: "number" },
              unit: { type: "string" },
              category: { type: "string", enum: ["financial", "operational", "strategic", "risk"] },
              frequency: { type: "string", enum: ["daily", "weekly", "monthly", "quarterly"] },
              threshold: {
                type: "object",
                properties: {
                  excellent: { type: "number" },
                  good: { type: "number" },
                  warning: { type: "number" },
                  critical: { type: "number" }
                },
                required: ["excellent", "good", "warning", "critical"]
              },
              trend: { type: "string", enum: ["increasing", "decreasing", "stable"] },
              importance: { type: "string", enum: ["high", "medium", "low"] }
            },
            required: ["name", "current", "target", "previous", "unit", "category", "threshold", "trend", "importance"]
          }
        },
        worksheetName: { type: "string", default: "KPI Dashboard" }
      },
      required: ["kpis"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        // Generate KPI-focused dashboard
        const worksheetData = dashboard.generateKPIDashboard(args.kpis);
        const worksheetName = args.worksheetName || "KPI Dashboard";
        
        await excelManager.addWorksheet(worksheetName);
        await excelManager.writeWorksheet(worksheetName, worksheetData);
        await excelManager.autoFitColumnWidths(worksheetName, { minWidth: 80, maxWidth: 200 });
        
        // Calculate KPI summary
        const criticalKPIs = args.kpis.filter((kpi: KPIMetric) => kpi.current < kpi.threshold.warning);
        const excellentKPIs = args.kpis.filter((kpi: KPIMetric) => kpi.current >= kpi.threshold.excellent);
        
        const categoryBreakdown = args.kpis.reduce((acc: any, kpi: KPIMetric) => {
          if (!acc[kpi.category]) acc[kpi.category] = 0;
          acc[kpi.category]++;
          return acc;
        }, {});
        
        return {
          success: true,
          message: `Generated KPI dashboard with ${args.kpis.length} performance indicators`,
          data: {
            worksheetName,
            kpiSummary: {
              totalKPIs: args.kpis.length,
              criticalKPIs: criticalKPIs.length,
              excellentKPIs: excellentKPIs.length,
              categoryBreakdown,
              performanceRate: ((args.kpis.length - criticalKPIs.length) / args.kpis.length * 100).toFixed(1) + '%'
            }
          }
        };
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  }
];