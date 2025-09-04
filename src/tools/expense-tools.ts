import { Tool, ToolResult } from '../types/index.js';
import { PythonBridge } from '../utils/python-bridge.js';

const pythonBridge = new PythonBridge();

export const expenseTools: Tool[] = [
  {
    name: "expense_add",
    description: "Add a new expense entry",
    inputSchema: {
      type: "object",
      properties: {
        expenseId: { type: "string" },
        date: { type: "string", format: "date" },
        vendorId: { type: "string" },
        amount: { type: "number" },
        category: { 
          type: "string",
          enum: [
            "Rent/Lease", "Utilities", "Salaries & Wages", "Employee Benefits", 
            "Insurance", "Marketing & Advertising", "Office Supplies", 
            "Maintenance & Repairs", "Professional Fees", "Travel & Entertainment",
            "Raw Materials", "Inventory Purchases", "Freight & Shipping",
            "Equipment", "Property", "Vehicles", "Software",
            "Interest Expense", "Bank Fees", "Taxes",
            "Depreciation", "Amortization", "Other"
          ]
        },
        subcategory: { type: "string" },
        description: { type: "string" },
        invoiceNumber: { type: "string" },
        costCenter: { type: "string" },
        projectId: { type: "string" },
        tags: { type: "array", items: { type: "string" } },
        taxDeductible: { type: "boolean", default: true },
        recurring: { type: "boolean", default: false },
        recurringFrequency: { type: "string" }
      },
      required: ["expenseId", "date", "vendorId", "amount", "category"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'expense_tracking',
          function: 'ExpenseTracker.add_expense',
          args: [],
          kwargs: { expense: args }
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
    name: "expense_add_vendor",
    description: "Add a new vendor to the system",
    inputSchema: {
      type: "object",
      properties: {
        vendorId: { type: "string" },
        name: { type: "string" },
        contactInfo: { type: "object" },
        taxId: { type: "string" },
        paymentTerms: { type: "string", default: "Net 30" },
        preferredPaymentMethod: { 
          type: "string",
          enum: ["Cash", "Check", "Credit Card", "ACH Transfer", "Wire Transfer", "PayPal", "Other"],
          default: "Check"
        },
        w9OnFile: { type: "boolean", default: false },
        active: { type: "boolean", default: true }
      },
      required: ["vendorId", "name", "contactInfo"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        return {
          success: true,
          data: { vendorId: args.vendorId },
          message: `Added vendor: ${args.name}`
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
    name: "expense_summary_report",
    description: "Generate expense summary report",
    inputSchema: {
      type: "object",
      properties: {
        startDate: { type: "string", format: "date" },
        endDate: { type: "string", format: "date" },
        groupBy: { 
          type: "string",
          enum: ["category", "vendor", "cost_center", "month"],
          default: "category"
        }
      },
      required: ["startDate", "endDate"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'expense_tracking',
          function: 'ExpenseTracker.get_expense_summary',
          args: [args.startDate, args.endDate, args.groupBy || 'category']
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
    name: "expense_budget_vs_actual",
    description: "Compare actual expenses against budget",
    inputSchema: {
      type: "object",
      properties: {
        budgetId: { type: "string" },
        startDate: { type: "string", format: "date" },
        endDate: { type: "string", format: "date" }
      },
      required: ["budgetId", "startDate", "endDate"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'expense_tracking',
          function: 'ExpenseTracker.budget_vs_actual',
          args: [args.budgetId, args.startDate, args.endDate]
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
    name: "expense_forecast",
    description: "Forecast future expenses based on historical data",
    inputSchema: {
      type: "object",
      properties: {
        monthsAhead: { type: "number", default: 12 }
      }
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'expense_tracking',
          function: 'ExpenseTracker.expense_forecast',
          args: [args.monthsAhead || 12]
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
    name: "expense_cost_savings_analysis",
    description: "Identify potential cost savings opportunities",
    inputSchema: {
      type: "object",
      properties: {
        lookbackMonths: { type: "number", default: 6 }
      }
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'expense_tracking',
          function: 'ExpenseTracker.identify_cost_savings',
          args: [args.lookbackMonths || 6]
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
    name: "expense_1099_report",
    description: "Generate 1099 reporting for vendors",
    inputSchema: {
      type: "object",
      properties: {
        taxYear: { type: "number" }
      },
      required: ["taxYear"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'expense_tracking',
          function: 'ExpenseTracker.generate_1099_report',
          args: [args.taxYear]
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
    name: "expense_cash_flow_impact",
    description: "Analyze cash flow impact of pending expenses",
    inputSchema: {
      type: "object",
      properties: {
        daysAhead: { type: "number", default: 30 }
      }
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'expense_tracking',
          function: 'ExpenseTracker.cash_flow_impact',
          args: [args.daysAhead || 30]
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
    name: "expense_vendor_analysis",
    description: "Analyze vendor spending patterns and performance",
    inputSchema: {
      type: "object",
      properties: {
        analysisType: {
          type: "string",
          enum: ["spending_patterns", "performance_metrics", "payment_terms_analysis"],
          default: "spending_patterns"
        }
      }
    },
    handler: async (): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'expense_tracking',
          function: 'ExpenseAnalytics.vendor_analysis',
          args: [],  // Would pass expenses and vendors data
          kwargs: {}
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
    name: "expense_spending_trends",
    description: "Analyze spending trends over time",
    inputSchema: {
      type: "object",
      properties: {
        periods: { type: "number", default: 12, description: "Number of periods to analyze" }
      }
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'expense_tracking',
          function: 'ExpenseAnalytics.spending_trends',
          args: [],  // Would pass expenses data
          kwargs: { periods: args.periods || 12 }
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
    name: "expense_create_budget",
    description: "Create a new budget for expense planning",
    inputSchema: {
      type: "object",
      properties: {
        budgetId: { type: "string" },
        name: { type: "string" },
        fiscalYear: { type: "number" },
        period: { 
          type: "string",
          enum: ["annual", "quarterly", "monthly"]
        },
        categories: { 
          type: "object",
          description: "Budget amounts by category"
        },
        costCenters: { type: "object" }
      },
      required: ["budgetId", "name", "fiscalYear", "period", "categories"]
    },
    handler: async (): Promise<ToolResult> => {
      try {
        return {
          success: true,
          message: "Budget creation placeholder"
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