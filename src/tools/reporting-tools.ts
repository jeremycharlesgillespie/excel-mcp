import { Tool, ToolResult } from '../types/index.js';
import { PythonBridge } from '../utils/python-bridge.js';
import { ExcelManager, WorksheetData } from '../excel/excel-manager.js';

const pythonBridge = new PythonBridge();
const excelManager = new ExcelManager();

export const reportingTools: Tool[] = [
  {
    name: "generate_income_statement",
    description: "Generate Income Statement (P&L) for specified period",
    inputSchema: {
      type: "object",
      properties: {
        startDate: { type: "string", format: "date" },
        endDate: { type: "string", format: "date" },
        outputToExcel: { type: "boolean", default: false },
        filePath: { type: "string" }
      },
      required: ["startDate", "endDate"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'financial_reporting',
          function: 'FinancialReportGenerator.generate_income_statement',
          args: [args.startDate, args.endDate]
        });

        if (args.outputToExcel && result.success && args.filePath) {
          // Convert to Excel format
          const incomeStatement = result.data;
          const worksheetData = formatIncomeStatementForExcel(incomeStatement);
          
          await excelManager.createWorkbook([worksheetData]);
          await excelManager.saveWorkbook(args.filePath);
        }

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
    name: "generate_balance_sheet",
    description: "Generate Balance Sheet as of specified date",
    inputSchema: {
      type: "object",
      properties: {
        asOfDate: { type: "string", format: "date" },
        outputToExcel: { type: "boolean", default: false },
        filePath: { type: "string" }
      },
      required: ["asOfDate"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'financial_reporting',
          function: 'FinancialReportGenerator.generate_balance_sheet',
          args: [args.asOfDate]
        });

        if (args.outputToExcel && result.success && args.filePath) {
          const balanceSheet = result.data;
          const worksheetData = formatBalanceSheetForExcel(balanceSheet);
          
          await excelManager.createWorkbook([worksheetData]);
          await excelManager.saveWorkbook(args.filePath);
        }

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
    name: "generate_trial_balance",
    description: "Generate Trial Balance as of specified date",
    inputSchema: {
      type: "object",
      properties: {
        asOfDate: { type: "string", format: "date" },
        outputToExcel: { type: "boolean", default: false },
        filePath: { type: "string" }
      },
      required: ["asOfDate"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'financial_reporting',
          function: 'FinancialReportGenerator.generate_trial_balance',
          args: [args.asOfDate]
        });

        if (args.outputToExcel && result.success && args.filePath) {
          const trialBalance = result.data;
          const worksheetData: WorksheetData = {
            name: 'Trial Balance',
            data: [
              ['Account Number', 'Account Name', 'Account Type', 'Debit', 'Credit'],
              ...trialBalance.map((row: any) => [
                row['Account Number'],
                row['Account Name'],
                row['Account Type'],
                row['Debit'] || '',
                row['Credit'] || ''
              ])
            ]
          };
          
          await excelManager.createWorkbook([worksheetData]);
          await excelManager.saveWorkbook(args.filePath);
        }

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
    name: "comparative_income_statement",
    description: "Generate comparative income statement for multiple periods",
    inputSchema: {
      type: "object",
      properties: {
        periods: {
          type: "array",
          items: {
            type: "object",
            properties: {
              startDate: { type: "string", format: "date" },
              endDate: { type: "string", format: "date" }
            },
            required: ["startDate", "endDate"]
          }
        },
        outputToExcel: { type: "boolean", default: false },
        filePath: { type: "string" }
      },
      required: ["periods"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const periodTuples = args.periods.map((p: any) => [p.startDate, p.endDate]);
        
        const result = await pythonBridge.callPythonFunction({
          module: 'financial_reporting',
          function: 'FinancialReportGenerator.comparative_income_statement',
          args: [periodTuples]
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
    name: "financial_ratios_analysis",
    description: "Calculate comprehensive financial ratios",
    inputSchema: {
      type: "object",
      properties: {
        asOfDate: { type: "string", format: "date" }
      },
      required: ["asOfDate"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'financial_reporting',
          function: 'FinancialReportGenerator.financial_ratios_analysis',
          args: [args.asOfDate]
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
    name: "create_rental_property_analysis",
    description: "Create rental property analysis Excel template",
    inputSchema: {
      type: "object",
      properties: {
        propertyData: {
          type: "object",
          properties: {
            name: { type: "string" },
            address: { type: "string" },
            grossRentalIncome: { type: "number" },
            vacancyLoss: { type: "number" },
            effectiveRentalIncome: { type: "number" },
            otherIncome: { type: "number" },
            propertyManagement: { type: "number" },
            maintenance: { type: "number" },
            insurance: { type: "number" },
            propertyTaxes: { type: "number" },
            utilities: { type: "number" },
            capRate: { type: "number" }
          }
        },
        filePath: { type: "string" }
      },
      required: ["propertyData", "filePath"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'financial_reporting',
          function: 'ReportTemplates.create_rental_property_analysis',
          args: [args.propertyData]
        });

        if (result.success && args.filePath) {
          const templateData = result.data;
          const worksheetData: WorksheetData = {
            name: 'Property Analysis',
            data: templateData
          };
          
          await excelManager.createWorkbook([worksheetData]);
          await excelManager.saveWorkbook(args.filePath);
        }

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
    name: "create_cash_flow_template",
    description: "Create cash flow projection Excel template",
    inputSchema: {
      type: "object",
      properties: {
        months: { type: "number", default: 12 },
        filePath: { type: "string" }
      },
      required: ["filePath"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'financial_reporting',
          function: 'ReportTemplates.create_cash_flow_template',
          args: [args.months || 12]
        });

        if (result.success && args.filePath) {
          const templateData = result.data;
          const worksheetData: WorksheetData = {
            name: 'Cash Flow Projection',
            data: templateData
          };
          
          await excelManager.createWorkbook([worksheetData]);
          await excelManager.saveWorkbook(args.filePath);
        }

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
    name: "create_budget_template",
    description: "Create budget vs actual Excel template",
    inputSchema: {
      type: "object",
      properties: {
        categories: { 
          type: "array", 
          items: { type: "string" },
          default: [
            "Salaries & Wages", "Rent", "Utilities", "Marketing", 
            "Office Supplies", "Professional Fees", "Travel", "Insurance"
          ]
        },
        filePath: { type: "string" }
      },
      required: ["filePath"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const categories = args.categories || [
          "Salaries & Wages", "Rent", "Utilities", "Marketing", 
          "Office Supplies", "Professional Fees", "Travel", "Insurance"
        ];
        
        const result = await pythonBridge.callPythonFunction({
          module: 'financial_reporting',
          function: 'ReportTemplates.create_budget_template',
          args: [categories]
        });

        if (result.success && args.filePath) {
          const templateData = result.data;
          const worksheetData: WorksheetData = {
            name: 'Budget vs Actual',
            data: templateData
          };
          
          await excelManager.createWorkbook([worksheetData]);
          await excelManager.saveWorkbook(args.filePath);
        }

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

function formatIncomeStatementForExcel(incomeStatement: any): WorksheetData {
  const data: any[][] = [
    ['INCOME STATEMENT', '', ''],
    [`Period: ${incomeStatement.period}`, '', ''],
    ['', '', ''],
    ['REVENUE', '', ''],
    ...incomeStatement.revenue.line_items.map((item: any) => [
      item.account_name, '', item.amount
    ]),
    ['Total Revenue', '', incomeStatement.revenue.total],
    ['', '', ''],
    ['COST OF GOODS SOLD', '', ''],
    ...incomeStatement.cost_of_goods_sold.line_items.map((item: any) => [
      item.account_name, '', item.amount
    ]),
    ['Total Cost of Goods Sold', '', incomeStatement.cost_of_goods_sold.total],
    ['', '', ''],
    ['GROSS PROFIT', '', incomeStatement.gross_profit],
    [`Gross Margin %`, '', `${incomeStatement.gross_margin_pct}%`],
    ['', '', ''],
    ['OPERATING EXPENSES', '', ''],
    ['Operating Expenses', '', incomeStatement.operating_expenses.operating],
    ['Administrative Expenses', '', incomeStatement.operating_expenses.administrative],
    ['Selling Expenses', '', incomeStatement.operating_expenses.selling],
    ['Total Operating Expenses', '', incomeStatement.operating_expenses.total],
    ['', '', ''],
    ['OPERATING INCOME', '', incomeStatement.operating_income],
    [`Operating Margin %`, '', `${incomeStatement.operating_margin_pct}%`],
    ['', '', ''],
    ['OTHER INCOME/EXPENSE', '', ''],
    ['Interest Expense', '', incomeStatement.other_income_expense.interest_expense],
    ['Other Income', '', incomeStatement.other_income_expense.other_income],
    ['', '', ''],
    ['NET INCOME', '', incomeStatement.net_income],
    [`Net Margin %`, '', `${incomeStatement.net_margin_pct}%`]
  ];

  return {
    name: 'Income Statement',
    data: data,
    columns: [
      { header: 'Item', key: 'item', width: 30 },
      { header: 'Notes', key: 'notes', width: 20 },
      { header: 'Amount', key: 'amount', width: 15 }
    ]
  };
}

function formatBalanceSheetForExcel(balanceSheet: any): WorksheetData {
  const data: any[][] = [
    ['BALANCE SHEET', '', ''],
    [`As of: ${balanceSheet.as_of_date}`, '', ''],
    ['', '', ''],
    ['ASSETS', '', ''],
    ['Current Assets:', '', ''],
    ...balanceSheet.assets.current_assets.line_items.map((item: any) => [
      `  ${item.account_name}`, '', item.amount
    ]),
    ['Total Current Assets', '', balanceSheet.assets.current_assets.total],
    ['', '', ''],
    ['Fixed Assets:', '', ''],
    ...balanceSheet.assets.fixed_assets.line_items.map((item: any) => [
      `  ${item.account_name}`, '', item.amount
    ]),
    ['Total Fixed Assets', '', balanceSheet.assets.fixed_assets.total],
    ['', '', ''],
    ['TOTAL ASSETS', '', balanceSheet.assets.total_assets],
    ['', '', ''],
    ['LIABILITIES & EQUITY', '', ''],
    ['Current Liabilities:', '', ''],
    ...balanceSheet.liabilities.current_liabilities.line_items.map((item: any) => [
      `  ${item.account_name}`, '', item.amount
    ]),
    ['Total Current Liabilities', '', balanceSheet.liabilities.current_liabilities.total],
    ['', '', ''],
    ['Long-term Liabilities:', '', ''],
    ...balanceSheet.liabilities.long_term_liabilities.line_items.map((item: any) => [
      `  ${item.account_name}`, '', item.amount
    ]),
    ['Total Long-term Liabilities', '', balanceSheet.liabilities.long_term_liabilities.total],
    ['', '', ''],
    ['Total Liabilities', '', balanceSheet.liabilities.total_liabilities],
    ['', '', ''],
    ['Equity:', '', ''],
    ...balanceSheet.equity.line_items.map((item: any) => [
      `  ${item.account_name}`, '', item.amount
    ]),
    ['Total Equity', '', balanceSheet.equity.total],
    ['', '', ''],
    ['TOTAL LIABILITIES & EQUITY', '', balanceSheet.total_liabilities_and_equity],
    ['', '', ''],
    ['Balanced:', '', balanceSheet.balanced ? 'YES' : 'NO']
  ];

  return {
    name: 'Balance Sheet',
    data: data,
    columns: [
      { header: 'Item', key: 'item', width: 30 },
      { header: 'Notes', key: 'notes', width: 20 },
      { header: 'Amount', key: 'amount', width: 15 }
    ]
  };
}