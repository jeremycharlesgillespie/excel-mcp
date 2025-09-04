import { Tool, ToolResult } from '../types/index.js';
import { PythonBridge } from '../utils/python-bridge.js';

const pythonBridge = new PythonBridge();

export const financialTools: Tool[] = [
  {
    name: "calculate_npv",
    description: "Calculate Net Present Value of cash flows",
    inputSchema: {
      type: "object",
      properties: {
        rate: { type: "number", description: "Discount rate as decimal (e.g., 0.1 for 10%)" },
        cashFlows: { type: "array", items: { type: "number" } },
        initialInvestment: { type: "number", default: 0 }
      },
      required: ["rate", "cashFlows"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'financial_calculations',
          function: 'FinancialCalculator.npv',
          args: [args.rate, args.cashFlows, args.initialInvestment || 0]
        });
        return result;
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "calculate_irr",
    description: "Calculate Internal Rate of Return",
    inputSchema: {
      type: "object",
      properties: {
        cashFlows: { type: "array", items: { type: "number" }, description: "Cash flows including initial investment" }
      },
      required: ["cashFlows"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'financial_calculations',
          function: 'FinancialCalculator.irr',
          args: [args.cashFlows]
        });
        return result;
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "calculate_mirr",
    description: "Calculate Modified Internal Rate of Return",
    inputSchema: {
      type: "object",
      properties: {
        cashFlows: { type: "array", items: { type: "number" } },
        financeRate: { type: "number", description: "Finance rate for negative cash flows" },
        reinvestRate: { type: "number", description: "Reinvestment rate for positive cash flows" }
      },
      required: ["cashFlows", "financeRate", "reinvestRate"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'financial_calculations',
          function: 'FinancialCalculator.mirr',
          args: [args.cashFlows, args.financeRate, args.reinvestRate]
        });
        return result;
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "loan_amortization",
    description: "Generate loan amortization schedule",
    inputSchema: {
      type: "object",
      properties: {
        principal: { type: "number" },
        annualRate: { type: "number", description: "Annual interest rate as decimal" },
        years: { type: "number" },
        paymentFrequency: { 
          type: "string", 
          enum: ["monthly", "quarterly", "semi-annually", "annually"],
          default: "monthly"
        }
      },
      required: ["principal", "annualRate", "years"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'financial_calculations',
          function: 'FinancialCalculator.loan_amortization',
          args: [args.principal, args.annualRate, args.years, args.paymentFrequency || 'monthly']
        });
        return result;
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "calculate_bond_price",
    description: "Calculate bond price given yield and terms",
    inputSchema: {
      type: "object",
      properties: {
        faceValue: { type: "number" },
        couponRate: { type: "number", description: "Annual coupon rate as decimal" },
        yieldRate: { type: "number", description: "Required yield as decimal" },
        years: { type: "number" },
        frequency: { type: "number", default: 2, description: "Coupon payments per year" }
      },
      required: ["faceValue", "couponRate", "yieldRate", "years"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'financial_calculations',
          function: 'FinancialCalculator.bond_price',
          args: [args.faceValue, args.couponRate, args.yieldRate, args.years, args.frequency || 2]
        });
        return result;
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "calculate_wacc",
    description: "Calculate Weighted Average Cost of Capital",
    inputSchema: {
      type: "object",
      properties: {
        equityValue: { type: "number" },
        debtValue: { type: "number" },
        costOfEquity: { type: "number" },
        costOfDebt: { type: "number" },
        taxRate: { type: "number" }
      },
      required: ["equityValue", "debtValue", "costOfEquity", "costOfDebt", "taxRate"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'financial_calculations',
          function: 'FinancialCalculator.wacc',
          args: [args.equityValue, args.debtValue, args.costOfEquity, args.costOfDebt, args.taxRate]
        });
        return result;
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "depreciation_straight_line",
    description: "Calculate straight-line depreciation schedule",
    inputSchema: {
      type: "object",
      properties: {
        cost: { type: "number" },
        salvageValue: { type: "number" },
        usefulLife: { type: "number", description: "Useful life in years" }
      },
      required: ["cost", "salvageValue", "usefulLife"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'financial_calculations',
          function: 'DepreciationCalculator.straight_line',
          args: [args.cost, args.salvageValue, args.usefulLife]
        });
        return result;
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "depreciation_declining_balance",
    description: "Calculate declining balance depreciation schedule",
    inputSchema: {
      type: "object",
      properties: {
        cost: { type: "number" },
        salvageValue: { type: "number" },
        usefulLife: { type: "number" },
        rate: { type: "number", default: 2.0, description: "Declining balance rate (e.g., 2.0 for double-declining)" }
      },
      required: ["cost", "salvageValue", "usefulLife"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'financial_calculations',
          function: 'DepreciationCalculator.declining_balance',
          args: [args.cost, args.salvageValue, args.usefulLife, args.rate || 2.0]
        });
        return result;
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "calculate_financial_ratios",
    description: "Calculate comprehensive financial ratios",
    inputSchema: {
      type: "object",
      properties: {
        ratioType: { 
          type: "string",
          enum: ["current_ratio", "quick_ratio", "debt_to_equity", "return_on_assets", "return_on_equity", 
                 "gross_margin", "operating_margin", "net_margin", "inventory_turnover", "asset_turnover"]
        },
        values: { type: "object", description: "Financial values needed for calculation" }
      },
      required: ["ratioType", "values"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'financial_calculations',
          function: `RatioAnalysis.${args.ratioType}`,
          args: Object.values(args.values)
        });
        return result;
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "effective_annual_rate",
    description: "Calculate effective annual rate from nominal rate",
    inputSchema: {
      type: "object",
      properties: {
        nominalRate: { type: "number", description: "Nominal annual rate as decimal" },
        compoundingPeriods: { type: "number", description: "Number of compounding periods per year" }
      },
      required: ["nominalRate", "compoundingPeriods"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'financial_calculations',
          function: 'FinancialCalculator.effective_annual_rate',
          args: [args.nominalRate, args.compoundingPeriods]
        });
        return result;
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "payback_period",
    description: "Calculate payback period for investment",
    inputSchema: {
      type: "object",
      properties: {
        initialInvestment: { type: "number" },
        cashFlows: { type: "array", items: { type: "number" } },
        discounted: { type: "boolean", default: false },
        discountRate: { type: "number", description: "Required if discounted = true" }
      },
      required: ["initialInvestment", "cashFlows"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const functionName = args.discounted ? 'discounted_payback_period' : 'payback_period';
        const functionArgs = args.discounted 
          ? [args.initialInvestment, args.cashFlows, args.discountRate]
          : [args.initialInvestment, args.cashFlows];
        
        const result = await pythonBridge.callPythonFunction({
          module: 'financial_calculations',
          function: `FinancialCalculator.${functionName}`,
          args: functionArgs
        });
        return result;
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "profitability_index",
    description: "Calculate profitability index (PI)",
    inputSchema: {
      type: "object",
      properties: {
        initialInvestment: { type: "number" },
        cashFlows: { type: "array", items: { type: "number" } },
        rate: { type: "number", description: "Discount rate as decimal" }
      },
      required: ["initialInvestment", "cashFlows", "rate"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'financial_calculations',
          function: 'FinancialCalculator.profitability_index',
          args: [args.initialInvestment, args.cashFlows, args.rate]
        });
        return result;
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "capm_expected_return",
    description: "Calculate expected return using CAPM model",
    inputSchema: {
      type: "object",
      properties: {
        riskFreeRate: { type: "number" },
        beta: { type: "number" },
        marketReturn: { type: "number" }
      },
      required: ["riskFreeRate", "beta", "marketReturn"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'financial_calculations',
          function: 'FinancialCalculator.capm',
          args: [args.riskFreeRate, args.beta, args.marketReturn]
        });
        return result;
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "bond_duration",
    description: "Calculate Macaulay duration for bond",
    inputSchema: {
      type: "object",
      properties: {
        cashFlows: {
          type: "array",
          items: {
            type: "array",
            items: { type: "number" },
            minItems: 2,
            maxItems: 2
          },
          description: "Array of [time, cashFlow] pairs"
        },
        yieldRate: { type: "number" }
      },
      required: ["cashFlows", "yieldRate"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'financial_calculations',
          function: 'FinancialCalculator.duration',
          args: [args.cashFlows, args.yieldRate]
        });
        return result;
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "depreciation_sum_of_years",
    description: "Calculate sum-of-years-digits depreciation schedule",
    inputSchema: {
      type: "object",
      properties: {
        cost: { type: "number" },
        salvageValue: { type: "number" },
        usefulLife: { type: "number" }
      },
      required: ["cost", "salvageValue", "usefulLife"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'financial_calculations',
          function: 'DepreciationCalculator.sum_of_years_digits',
          args: [args.cost, args.salvageValue, args.usefulLife]
        });
        return result;
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "depreciation_units_production",
    description: "Calculate units of production depreciation",
    inputSchema: {
      type: "object",
      properties: {
        cost: { type: "number" },
        salvageValue: { type: "number" },
        totalUnits: { type: "number" },
        unitsPerPeriod: { type: "array", items: { type: "number" } }
      },
      required: ["cost", "salvageValue", "totalUnits", "unitsPerPeriod"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'financial_calculations',
          function: 'DepreciationCalculator.units_of_production',
          args: [args.cost, args.salvageValue, args.totalUnits, args.unitsPerPeriod]
        });
        return result;
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "depreciation_macrs",
    description: "Calculate MACRS depreciation schedule",
    inputSchema: {
      type: "object",
      properties: {
        cost: { type: "number" },
        recoveryPeriod: { 
          type: "number",
          enum: [3, 5, 7, 10],
          description: "MACRS recovery period in years"
        }
      },
      required: ["cost", "recoveryPeriod"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'financial_calculations',
          function: 'DepreciationCalculator.macrs',
          args: [args.cost, args.recoveryPeriod]
        });
        return result;
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "future_value",
    description: "Calculate future value of present amount",
    inputSchema: {
      type: "object",
      properties: {
        presentValue: { type: "number" },
        rate: { type: "number" },
        periods: { type: "number" }
      },
      required: ["presentValue", "rate", "periods"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'financial_calculations',
          function: 'FinancialCalculator.future_value',
          args: [args.presentValue, args.rate, args.periods]
        });
        return result;
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "present_value",
    description: "Calculate present value of future amount",
    inputSchema: {
      type: "object",
      properties: {
        futureValue: { type: "number" },
        rate: { type: "number" },
        periods: { type: "number" }
      },
      required: ["futureValue", "rate", "periods"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'financial_calculations',
          function: 'FinancialCalculator.present_value',
          args: [args.futureValue, args.rate, args.periods]
        });
        return result;
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  }
];