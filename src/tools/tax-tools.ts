import { Tool, ToolResult } from '../types/index.js';
import { PythonBridge } from '../utils/python-bridge.js';

const pythonBridge = new PythonBridge();

export const taxTools: Tool[] = [
  {
    name: "calculate_federal_income_tax",
    description: "Calculate federal income tax based on taxable income and filing status",
    inputSchema: {
      type: "object",
      properties: {
        taxableIncome: { type: "number" },
        filingStatus: { 
          type: "string",
          enum: ["single", "married_filing_jointly", "married_filing_separately", "head_of_household"],
          default: "single"
        }
      },
      required: ["taxableIncome"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'tax_calculations',
          function: 'TaxCalculator.calculate_federal_income_tax',
          args: [args.taxableIncome, args.filingStatus || 'single']
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
    name: "calculate_self_employment_tax",
    description: "Calculate self-employment tax (Social Security and Medicare)",
    inputSchema: {
      type: "object",
      properties: {
        netEarnings: { type: "number", description: "Net earnings from self-employment" }
      },
      required: ["netEarnings"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'tax_calculations',
          function: 'TaxCalculator.calculate_self_employment_tax',
          args: [args.netEarnings]
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
    name: "calculate_payroll_taxes",
    description: "Calculate payroll taxes for employer and employee",
    inputSchema: {
      type: "object",
      properties: {
        wages: { type: "number" },
        year: { type: "number", default: 2024 }
      },
      required: ["wages"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'tax_calculations',
          function: 'TaxCalculator.calculate_payroll_taxes',
          args: [args.wages, args.year || 2024]
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
    name: "calculate_state_taxes",
    description: "Calculate state income tax for supported states",
    inputSchema: {
      type: "object",
      properties: {
        state: { 
          type: "string", 
          enum: ["CA", "NY", "TX", "FL"],
          description: "State abbreviation" 
        },
        taxableIncome: { type: "number" },
        filingStatus: { 
          type: "string",
          enum: ["single", "married_filing_jointly"],
          default: "single"
        }
      },
      required: ["state", "taxableIncome"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'tax_calculations',
          function: 'TaxCalculator.calculate_state_taxes',
          args: [args.state, args.taxableIncome, args.filingStatus || 'single']
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
    name: "estimated_quarterly_taxes",
    description: "Calculate estimated quarterly tax payments",
    inputSchema: {
      type: "object",
      properties: {
        annualIncome: { type: "number" },
        filingStatus: { 
          type: "string",
          enum: ["single", "married_filing_jointly"],
          default: "single"
        },
        selfEmployed: { type: "boolean", default: false }
      },
      required: ["annualIncome"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'tax_calculations',
          function: 'TaxCalculator.estimated_quarterly_taxes',
          args: [args.annualIncome, args.filingStatus || 'single', args.selfEmployed || false]
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
    name: "calculate_depreciation_deduction",
    description: "Calculate tax depreciation deduction for an asset",
    inputSchema: {
      type: "object",
      properties: {
        assetId: { type: "string" },
        taxYear: { type: "number" }
      },
      required: ["assetId", "taxYear"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'tax_calculations',
          function: 'TaxCalculator.calculate_depreciation_deduction',
          args: [args.assetId, args.taxYear]
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
    name: "business_deductions_analysis",
    description: "Analyze allowable business deductions",
    inputSchema: {
      type: "object",
      properties: {
        expenses: {
          type: "array",
          items: {
            type: "object",
            properties: {
              category: { type: "string" },
              amount: { type: "number" },
              description: { type: "string" }
            }
          }
        },
        entityType: {
          type: "string",
          enum: ["Sole Proprietorship", "Partnership", "S Corporation", "C Corporation", "LLC"]
        }
      },
      required: ["expenses", "entityType"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'tax_calculations',
          function: 'TaxCalculator.calculate_business_deductions',
          args: [args.expenses, args.entityType]
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
    name: "tax_planning_strategies",
    description: "Get tax planning strategy recommendations",
    inputSchema: {
      type: "object",
      properties: {
        currentIncome: { type: "number" },
        projectedIncome: { type: "number" },
        entityType: {
          type: "string",
          enum: ["Sole Proprietorship", "Partnership", "S Corporation", "C Corporation", "LLC"]
        }
      },
      required: ["currentIncome", "projectedIncome", "entityType"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'tax_calculations',
          function: 'TaxCalculator.tax_planning_strategies',
          args: [args.currentIncome, args.projectedIncome, args.entityType]
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
    name: "business_tax_summary",
    description: "Generate comprehensive business tax summary",
    inputSchema: {
      type: "object",
      properties: {
        entityId: { type: "string" },
        taxYear: { type: "number" },
        financialData: {
          type: "object",
          properties: {
            revenue: { type: "number" },
            expenses: { type: "number" }
          },
          required: ["revenue", "expenses"]
        }
      },
      required: ["entityId", "taxYear", "financialData"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'tax_calculations',
          function: 'TaxCalculator.business_tax_summary',
          args: [args.entityId, args.taxYear, args.financialData]
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
    name: "tax_projection_scenarios",
    description: "Generate tax projections under different income scenarios",
    inputSchema: {
      type: "object",
      properties: {
        entityId: { type: "string" },
        scenarios: {
          type: "object",
          description: "Scenarios with revenue and expenses"
        }
      },
      required: ["entityId", "scenarios"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'tax_calculations',
          function: 'TaxCalculator.generate_tax_projection',
          args: [args.entityId, args.scenarios]
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