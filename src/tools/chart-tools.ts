import { Tool, ToolResult } from '../types/index.js';
import { ChartIntelligence } from '../excel/chart-intelligence.js';

// Type definitions for xlsx-chart since it doesn't have TypeScript definitions
interface XLSXChartOptions {
  file?: string;
  chart: 'column' | 'bar' | 'line' | 'area' | 'radar' | 'scatter' | 'pie';
  titles: string[];
  fields: string[];
  data: { [title: string]: { [field: string]: number } };
  chartTitle?: string;
  mixedChart?: boolean;
  template?: string;
}

interface XLSXChart {
  writeFile(options: XLSXChartOptions, callback: (err: Error | null) => void): void;
  generate(options: XLSXChartOptions, callback: (err: Error | null, data?: Buffer) => void): void;
}

// Import the xlsx-chart module (will be dynamically required)
let XLSXChart: new () => XLSXChart;

export const chartTools: Tool[] = [
  {
    name: "excel_create_native_chart",
    description: "Create a native Excel chart file with data visualization",
    inputSchema: {
      type: "object",
      properties: {
        fileName: {
          type: "string",
          description: "Output Excel file name (e.g., 'sales-chart.xlsx')"
        },
        chartType: {
          type: "string",
          enum: ["column", "bar", "line", "area", "radar", "scatter", "pie"],
          description: "Type of chart to create"
        },
        chartTitle: {
          type: "string",
          description: "Title for the chart"
        },
        categories: {
          type: "array",
          items: { type: "string" },
          description: "Category names (x-axis labels)"
        },
        series: {
          type: "array",
          items: {
            type: "object",
            properties: {
              name: { type: "string", description: "Series name" },
              data: {
                type: "array",
                items: { type: "number" },
                description: "Data values for this series"
              }
            },
            required: ["name", "data"]
          },
          description: "Data series for the chart"
        }
      },
      required: ["fileName", "chartType", "categories", "series"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        // Dynamically import xlsx-chart
        if (!XLSXChart) {
          const xlsxChartModule = require('xlsx-chart');
          XLSXChart = xlsxChartModule;
        }

        const xlsxChart = new XLSXChart();

        // Transform the input data to xlsx-chart format
        const data: { [title: string]: { [field: string]: number } } = {};
        
        args.series.forEach((series: any) => {
          data[series.name] = {};
          args.categories.forEach((category: string, index: number) => {
            data[series.name][category] = series.data[index] || 0;
          });
        });

        const chartOptions: XLSXChartOptions = {
          file: args.fileName,
          chart: args.chartType,
          titles: args.series.map((s: any) => s.name),
          fields: args.categories,
          data: data,
          chartTitle: args.chartTitle
        };

        return new Promise((resolve) => {
          xlsxChart.writeFile(chartOptions, (err: Error | null) => {
            if (err) {
              resolve({
                success: false,
                error: `Failed to create chart: ${err.message}`
              });
            } else {
              resolve({
                success: true,
                message: `Created native Excel chart: ${args.fileName}`,
                data: {
                  fileName: args.fileName,
                  chartType: args.chartType,
                  seriesCount: args.series.length,
                  categoryCount: args.categories.length,
                  chartTitle: args.chartTitle
                }
              });
            }
          });
        });
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "excel_create_financial_chart",
    description: "Create financial charts with pre-configured templates for common financial visualizations",
    inputSchema: {
      type: "object",
      properties: {
        fileName: {
          type: "string",
          description: "Output Excel file name"
        },
        chartTemplate: {
          type: "string",
          enum: ["revenue_trend", "expense_breakdown", "cash_flow", "profit_loss", "portfolio_performance"],
          description: "Pre-configured financial chart template"
        },
        data: {
          type: "object",
          properties: {
            periods: {
              type: "array",
              items: { type: "string" },
              description: "Time periods (months, quarters, years)"
            },
            values: {
              type: "object",
              additionalProperties: {
                type: "array",
                items: { type: "number" }
              },
              description: "Financial values for each metric"
            }
          },
          required: ["periods", "values"]
        },
        title: {
          type: "string",
          description: "Chart title"
        }
      },
      required: ["fileName", "chartTemplate", "data"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        if (!XLSXChart) {
          const xlsxChartModule = require('xlsx-chart');
          XLSXChart = xlsxChartModule;
        }

        const xlsxChart = new XLSXChart();

        // Configure chart based on template
        let chartType: XLSXChartOptions['chart'];
        let chartTitle: string;

        switch (args.chartTemplate) {
          case 'revenue_trend':
            chartType = 'line';
            chartTitle = args.title || 'Revenue Trend Analysis';
            break;
          case 'expense_breakdown':
            chartType = 'pie';
            chartTitle = args.title || 'Expense Breakdown';
            break;
          case 'cash_flow':
            chartType = 'column';
            chartTitle = args.title || 'Cash Flow Analysis';
            break;
          case 'profit_loss':
            chartType = 'column';
            chartTitle = args.title || 'Profit & Loss Statement';
            break;
          case 'portfolio_performance':
            chartType = 'area';
            chartTitle = args.title || 'Portfolio Performance';
            break;
          default:
            chartType = 'column';
            chartTitle = args.title || 'Financial Analysis';
        }

        // Transform data
        const data: { [title: string]: { [field: string]: number } } = {};
        const seriesNames = Object.keys(args.data.values);
        
        seriesNames.forEach(seriesName => {
          data[seriesName] = {};
          args.data.periods.forEach((period: string, index: number) => {
            data[seriesName][period] = args.data.values[seriesName][index] || 0;
          });
        });

        const chartOptions: XLSXChartOptions = {
          file: args.fileName,
          chart: chartType,
          titles: seriesNames,
          fields: args.data.periods,
          data: data,
          chartTitle: chartTitle
        };

        return new Promise((resolve) => {
          xlsxChart.writeFile(chartOptions, (err: Error | null) => {
            if (err) {
              resolve({
                success: false,
                error: `Failed to create financial chart: ${err.message}`
              });
            } else {
              resolve({
                success: true,
                message: `Created ${args.chartTemplate} chart: ${args.fileName}`,
                data: {
                  fileName: args.fileName,
                  template: args.chartTemplate,
                  chartType: chartType,
                  title: chartTitle,
                  dataPoints: args.data.periods.length,
                  metrics: seriesNames
                }
              });
            }
          });
        });
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "excel_create_dashboard_chart",
    description: "Create a multi-series dashboard chart for comprehensive data visualization",
    inputSchema: {
      type: "object",
      properties: {
        fileName: {
          type: "string",
          description: "Output Excel file name"
        },
        dashboardTitle: {
          type: "string",
          description: "Title for the dashboard"
        },
        primaryChart: {
          type: "object",
          properties: {
            type: {
              type: "string",
              enum: ["column", "bar", "line", "area"],
              description: "Primary chart type"
            },
            title: { type: "string" },
            categories: {
              type: "array",
              items: { type: "string" }
            },
            series: {
              type: "array",
              items: {
                type: "object",
                properties: {
                  name: { type: "string" },
                  data: { type: "array", items: { type: "number" } }
                }
              }
            }
          },
          required: ["type", "categories", "series"]
        }
      },
      required: ["fileName", "primaryChart"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        if (!XLSXChart) {
          const xlsxChartModule = require('xlsx-chart');
          XLSXChart = xlsxChartModule;
        }

        const xlsxChart = new XLSXChart();
        const chart = args.primaryChart;

        // Transform data
        const data: { [title: string]: { [field: string]: number } } = {};
        
        chart.series.forEach((series: any) => {
          data[series.name] = {};
          chart.categories.forEach((category: string, index: number) => {
            data[series.name][category] = series.data[index] || 0;
          });
        });

        const chartOptions: XLSXChartOptions = {
          file: args.fileName,
          chart: chart.type,
          titles: chart.series.map((s: any) => s.name),
          fields: chart.categories,
          data: data,
          chartTitle: args.dashboardTitle || chart.title || 'Dashboard'
        };

        return new Promise((resolve) => {
          xlsxChart.writeFile(chartOptions, (err: Error | null) => {
            if (err) {
              resolve({
                success: false,
                error: `Failed to create dashboard chart: ${err.message}`
              });
            } else {
              resolve({
                success: true,
                message: `Created dashboard chart: ${args.fileName}`,
                data: {
                  fileName: args.fileName,
                  chartType: chart.type,
                  title: args.dashboardTitle || chart.title,
                  seriesCount: chart.series.length,
                  categoryCount: chart.categories.length
                }
              });
            }
          });
        });
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "excel_chart_from_data_range",
    description: "Create a native Excel chart from existing Excel data (requires data to be pre-loaded)",
    inputSchema: {
      type: "object",
      properties: {
        outputFileName: {
          type: "string",
          description: "Output chart file name"
        },
        chartType: {
          type: "string",
          enum: ["column", "bar", "line", "area", "radar", "scatter", "pie"],
          description: "Type of chart"
        },
        dataMatrix: {
          type: "array",
          items: {
            type: "array",
            items: { type: "string" }
          },
          description: "2D array of data with headers in first row and first column"
        },
        chartTitle: {
          type: "string",
          description: "Title for the chart"
        }
      },
      required: ["outputFileName", "chartType", "dataMatrix"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        if (!XLSXChart) {
          const xlsxChartModule = require('xlsx-chart');
          XLSXChart = xlsxChartModule;
        }

        const xlsxChart = new XLSXChart();
        const matrix = args.dataMatrix;

        if (matrix.length < 2 || matrix[0].length < 2) {
          return {
            success: false,
            error: "Data matrix must have at least 2 rows and 2 columns (including headers)"
          };
        }

        // Extract headers and data
        const fields = matrix[0].slice(1); // Column headers (skip first cell)
        const titles = matrix.slice(1).map((row: string[]) => row[0]); // Row headers
        
        // Transform data
        const data: { [title: string]: { [field: string]: number } } = {};
        
        matrix.slice(1).forEach((row: string[], rowIndex: number) => {
          const title = titles[rowIndex];
          data[title] = {};
          
          fields.forEach((field: string, fieldIndex: number) => {
            const value = parseFloat(row[fieldIndex + 1]) || 0;
            data[title][field] = value;
          });
        });

        const chartOptions: XLSXChartOptions = {
          file: args.outputFileName,
          chart: args.chartType,
          titles: titles,
          fields: fields,
          data: data,
          chartTitle: args.chartTitle || 'Data Visualization'
        };

        return new Promise((resolve) => {
          xlsxChart.writeFile(chartOptions, (err: Error | null) => {
            if (err) {
              resolve({
                success: false,
                error: `Failed to create chart from data: ${err.message}`
              });
            } else {
              resolve({
                success: true,
                message: `Created chart from data matrix: ${args.outputFileName}`,
                data: {
                  fileName: args.outputFileName,
                  chartType: args.chartType,
                  title: args.chartTitle,
                  seriesCount: titles.length,
                  categoryCount: fields.length,
                  dataPoints: titles.length * fields.length
                }
              });
            }
          });
        });
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "excel_create_smart_chart",
    description: "Intelligently select and create the best chart type based on data description and context",
    inputSchema: {
      type: "object",
      properties: {
        fileName: {
          type: "string",
          description: "Output Excel file name"
        },
        dataDescription: {
          type: "string",
          description: "Description of what the data represents (e.g., 'monthly cash flow', 'expense breakdown', 'quarterly revenue trend')"
        },
        categories: {
          type: "array",
          items: { type: "string" },
          description: "Category names (x-axis labels)"
        },
        series: {
          type: "array",
          items: {
            type: "object",
            properties: {
              name: { type: "string", description: "Series name" },
              data: {
                type: "array",
                items: { type: "number" },
                description: "Data values for this series"
              }
            },
            required: ["name", "data"]
          },
          description: "Data series for the chart"
        },
        chartTitle: {
          type: "string",
          description: "Optional title for the chart"
        },
        forceChartType: {
          type: "string",
          enum: ["column", "bar", "line", "area", "radar", "scatter", "pie"],
          description: "Override intelligent selection with specific chart type"
        }
      },
      required: ["fileName", "dataDescription", "categories", "series"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        if (!XLSXChart) {
          const xlsxChartModule = require('xlsx-chart');
          XLSXChart = xlsxChartModule;
        }

        const xlsxChart = new XLSXChart();

        // Get intelligent chart recommendation
        const recommendation = ChartIntelligence.recommendChart(
          args.dataDescription,
          args.categories,
          args.series.map((s: any) => s.name),
          args.series.map((s: any) => s.data)
        );

        // Use forced chart type if provided, otherwise use recommendation
        const selectedChartType = args.forceChartType || recommendation.primaryChart;

        // Transform the input data to xlsx-chart format
        const data: { [title: string]: { [field: string]: number } } = {};
        
        args.series.forEach((series: any) => {
          data[series.name] = {};
          args.categories.forEach((category: string, index: number) => {
            data[series.name][category] = series.data[index] || 0;
          });
        });

        const chartOptions: XLSXChartOptions = {
          file: args.fileName,
          chart: selectedChartType,
          titles: args.series.map((s: any) => s.name),
          fields: args.categories,
          data: data,
          chartTitle: args.chartTitle || args.dataDescription
        };

        return new Promise((resolve) => {
          xlsxChart.writeFile(chartOptions, (err: Error | null) => {
            if (err) {
              resolve({
                success: false,
                error: `Failed to create smart chart: ${err.message}`
              });
            } else {
              resolve({
                success: true,
                message: `Created intelligent chart: ${args.fileName}`,
                data: {
                  fileName: args.fileName,
                  selectedChartType: selectedChartType,
                  recommendedType: recommendation.primaryChart,
                  confidence: recommendation.confidence,
                  reasoning: recommendation.reasoning,
                  alternatives: recommendation.alternatives,
                  warnings: recommendation.warnings,
                  explanation: ChartIntelligence.explainRecommendation(recommendation),
                  dataContext: args.dataDescription,
                  seriesCount: args.series.length,
                  categoryCount: args.categories.length
                }
              });
            }
          });
        });
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "excel_recommend_chart_type",
    description: "Get chart type recommendations without creating a chart - useful for planning and decision making",
    inputSchema: {
      type: "object",
      properties: {
        dataDescription: {
          type: "string",
          description: "Description of what the data represents"
        },
        categories: {
          type: "array",
          items: { type: "string" },
          description: "Category names or labels"
        },
        seriesNames: {
          type: "array",
          items: { type: "string" },
          description: "Names of data series"
        },
        sampleData: {
          type: "array",
          items: {
            type: "array",
            items: { type: "number" }
          },
          description: "Optional sample data for analysis"
        }
      },
      required: ["dataDescription", "categories", "seriesNames"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const recommendation = ChartIntelligence.recommendChart(
          args.dataDescription,
          args.categories,
          args.seriesNames,
          args.sampleData
        );

        return {
          success: true,
          message: `Chart recommendation for: ${args.dataDescription}`,
          data: {
            primaryRecommendation: {
              chartType: recommendation.primaryChart,
              confidence: recommendation.confidence,
              reasoning: recommendation.reasoning
            },
            alternatives: recommendation.alternatives,
            warnings: recommendation.warnings,
            fullExplanation: ChartIntelligence.explainRecommendation(recommendation),
            dataContext: {
              description: args.dataDescription,
              categoryCount: args.categories.length,
              seriesCount: args.seriesNames.length,
              hasTimeElements: args.categories.some((cat: string) => 
                /month|quarter|year|week|day|q[1-4]|jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec/i.test(cat)
              )
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
    name: "excel_create_business_chart",
    description: "Create charts optimized for specific business scenarios with intelligent defaults",
    inputSchema: {
      type: "object",
      properties: {
        fileName: {
          type: "string",
          description: "Output Excel file name"
        },
        businessScenario: {
          type: "string",
          enum: [
            "monthly_cash_flow",
            "quarterly_revenue",
            "expense_breakdown", 
            "profit_trend",
            "department_comparison",
            "portfolio_performance",
            "budget_vs_actual",
            "regional_sales",
            "customer_segments",
            "seasonal_analysis"
          ],
          description: "Pre-defined business scenario for optimal chart selection"
        },
        data: {
          type: "object",
          properties: {
            categories: {
              type: "array",
              items: { type: "string" }
            },
            series: {
              type: "array",
              items: {
                type: "object",
                properties: {
                  name: { type: "string" },
                  data: { type: "array", items: { type: "number" } }
                }
              }
            }
          },
          required: ["categories", "series"]
        },
        title: {
          type: "string",
          description: "Custom chart title"
        }
      },
      required: ["fileName", "businessScenario", "data"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        if (!XLSXChart) {
          const xlsxChartModule = require('xlsx-chart');
          XLSXChart = xlsxChartModule;
        }

        const xlsxChart = new XLSXChart();

        // Map business scenarios to optimal chart types and descriptions
        const scenarioMapping: { [key: string]: { chartType: XLSXChartOptions['chart'], description: string } } = {
          monthly_cash_flow: { chartType: 'area', description: 'Monthly cash flow shows cumulative effect over time' },
          quarterly_revenue: { chartType: 'line', description: 'Quarterly revenue trends are best shown with line charts' },
          expense_breakdown: { chartType: 'pie', description: 'Expense breakdown shows parts of the whole budget' },
          profit_trend: { chartType: 'line', description: 'Profit trends over time require line visualization' },
          department_comparison: { chartType: 'column', description: 'Department comparison uses column charts for clear comparison' },
          portfolio_performance: { chartType: 'line', description: 'Portfolio performance tracking uses line charts' },
          budget_vs_actual: { chartType: 'column', description: 'Budget vs actual comparison uses grouped columns' },
          regional_sales: { chartType: 'column', description: 'Regional sales comparison uses column charts' },
          customer_segments: { chartType: 'pie', description: 'Customer segments show market composition' },
          seasonal_analysis: { chartType: 'line', description: 'Seasonal patterns require line charts to show trends' }
        };

        const scenario = scenarioMapping[args.businessScenario];
        if (!scenario) {
          return {
            success: false,
            error: `Unknown business scenario: ${args.businessScenario}`
          };
        }

        // Transform data
        const data: { [title: string]: { [field: string]: number } } = {};
        
        args.data.series.forEach((series: any) => {
          data[series.name] = {};
          args.data.categories.forEach((category: string, index: number) => {
            data[series.name][category] = series.data[index] || 0;
          });
        });

        const chartOptions: XLSXChartOptions = {
          file: args.fileName,
          chart: scenario.chartType,
          titles: args.data.series.map((s: any) => s.name),
          fields: args.data.categories,
          data: data,
          chartTitle: args.title || args.businessScenario.replace(/_/g, ' ').toUpperCase()
        };

        return new Promise((resolve) => {
          xlsxChart.writeFile(chartOptions, (err: Error | null) => {
            if (err) {
              resolve({
                success: false,
                error: `Failed to create business chart: ${err.message}`
              });
            } else {
              resolve({
                success: true,
                message: `Created ${args.businessScenario} chart: ${args.fileName}`,
                data: {
                  fileName: args.fileName,
                  businessScenario: args.businessScenario,
                  chartType: scenario.chartType,
                  reasoning: scenario.description,
                  title: args.title || args.businessScenario.replace(/_/g, ' ').toUpperCase(),
                  seriesCount: args.data.series.length,
                  categoryCount: args.data.categories.length
                }
              });
            }
          });
        });
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  }
];