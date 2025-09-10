import { Tool, ToolResult } from '../types/index.js';
import { ExcelManager } from '../excel/excel-manager.js';
import { FormulaGenerator } from '../excel/formula-generator.js';
import { ProfessionalTemplates } from '../excel/professional-templates.js';

const excelManager = new ExcelManager();

export const excelTools: Tool[] = [
  {
    name: "excel_create_workbook",
    description: "Create a new Excel workbook with specified worksheets and data",
    inputSchema: {
      type: "object",
      properties: {
        worksheets: {
          type: "array",
          items: {
            type: "object",
            properties: {
              name: { type: "string" },
              data: {
                type: "array",
                items: { type: "array" }
              },
              columns: {
                type: "array",
                items: {
                  type: "object",
                  properties: {
                    header: { type: "string" },
                    key: { type: "string" },
                    width: { type: "number" }
                  }
                }
              }
            },
            required: ["name", "data"]
          }
        }
      },
      required: ["worksheets"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        await excelManager.createWorkbook(args.worksheets);
        return {
          success: true,
          message: `Created workbook with ${args.worksheets.length} worksheets`
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
    name: "excel_open_file",
    description: "Open an existing Excel file",
    inputSchema: {
      type: "object",
      properties: {
        filePath: { type: "string" }
      },
      required: ["filePath"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const fileInfo = await excelManager.openWorkbook(args.filePath);
        return {
          success: true,
          data: fileInfo
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
    name: "excel_save_file",
    description: "Save the current workbook to a file",
    inputSchema: {
      type: "object",
      properties: {
        filePath: { type: "string" }
      }
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        await excelManager.saveWorkbook(args.filePath);
        return {
          success: true,
          message: `Saved workbook to ${args.filePath || 'current file'}`
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
    name: "excel_read_worksheet",
    description: "Read data from a specific worksheet",
    inputSchema: {
      type: "object",
      properties: {
        worksheetName: { type: "string" }
      },
      required: ["worksheetName"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const data = await excelManager.readWorksheet(args.worksheetName);
        return {
          success: true,
          data: data
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
    name: "excel_write_worksheet",
    description: "Write data to a worksheet",
    inputSchema: {
      type: "object",
      properties: {
        worksheetName: { type: "string" },
        data: {
          type: "array",
          items: { type: "array" }
        },
        startRow: { type: "number", default: 1 },
        startCol: { type: "number", default: 1 }
      },
      required: ["worksheetName", "data"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        await excelManager.writeWorksheet(
          args.worksheetName,
          args.data,
          args.startRow || 1,
          args.startCol || 1
        );
        return {
          success: true,
          message: `Data written to ${args.worksheetName}`
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
    name: "excel_add_worksheet",
    description: "Add a new worksheet to the current workbook",
    inputSchema: {
      type: "object",
      properties: {
        name: { type: "string" }
      },
      required: ["name"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        await excelManager.addWorksheet(args.name);
        return {
          success: true,
          message: `Added worksheet: ${args.name}`
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
    name: "excel_delete_worksheet",
    description: "Delete a worksheet from the current workbook",
    inputSchema: {
      type: "object",
      properties: {
        name: { type: "string" }
      },
      required: ["name"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        await excelManager.deleteWorksheet(args.name);
        return {
          success: true,
          message: `Deleted worksheet: ${args.name}`
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
    name: "excel_create_named_range",
    description: "Create a named range in the workbook",
    inputSchema: {
      type: "object",
      properties: {
        name: { type: "string" },
        range: { type: "string" },
        worksheetName: { type: "string" }
      },
      required: ["name", "range"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        await excelManager.createNamedRange(args.name, args.range, args.worksheetName);
        return {
          success: true,
          message: `Created named range: ${args.name}`
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
    name: "excel_find_replace",
    description: "Find and replace text in a worksheet",
    inputSchema: {
      type: "object",
      properties: {
        worksheetName: { type: "string" },
        searchText: { type: "string" },
        replaceText: { type: "string" },
        matchCase: { type: "boolean", default: false },
        matchEntireCell: { type: "boolean", default: false }
      },
      required: ["worksheetName", "searchText", "replaceText"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const count = await excelManager.findAndReplace(
          args.worksheetName,
          args.searchText,
          args.replaceText,
          {
            matchCase: args.matchCase,
            matchEntireCell: args.matchEntireCell
          }
        );
        return {
          success: true,
          data: { replacements: count },
          message: `Replaced ${count} instances`
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
    name: "excel_get_formulas",
    description: "Extract all formulas from a worksheet",
    inputSchema: {
      type: "object",
      properties: {
        worksheetName: { type: "string" }
      },
      required: ["worksheetName"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const formulas = await excelManager.getCellFormulas(args.worksheetName);
        return {
          success: true,
          data: formulas
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
    name: "excel_protect_worksheet",
    description: "Protect a worksheet with a password",
    inputSchema: {
      type: "object",
      properties: {
        worksheetName: { type: "string" },
        password: { type: "string" },
        allowSelectLockedCells: { type: "boolean", default: true },
        allowSelectUnlockedCells: { type: "boolean", default: true },
        allowFormatCells: { type: "boolean", default: false },
        allowFormatColumns: { type: "boolean", default: false },
        allowFormatRows: { type: "boolean", default: false }
      },
      required: ["worksheetName", "password"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const protection = {
          selectLockedCells: args.allowSelectLockedCells,
          selectUnlockedCells: args.allowSelectUnlockedCells,
          formatCells: args.allowFormatCells,
          formatColumns: args.allowFormatColumns,
          formatRows: args.allowFormatRows
        };
        
        await excelManager.protectWorksheet(args.worksheetName, args.password, protection);
        return {
          success: true,
          message: `Protected worksheet: ${args.worksheetName}`
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
    name: "excel_merge_files",
    description: "Merge multiple Excel files into one workbook",
    inputSchema: {
      type: "object",
      properties: {
        inputFiles: {
          type: "array",
          items: { type: "string" }
        },
        outputPath: { type: "string" }
      },
      required: ["inputFiles", "outputPath"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        await excelManager.mergeFiles(args.inputFiles, args.outputPath);
        return {
          success: true,
          message: `Merged ${args.inputFiles.length} files into ${args.outputPath}`
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
    name: "excel_conditional_formatting",
    description: "Apply conditional formatting to a range",
    inputSchema: {
      type: "object",
      properties: {
        worksheetName: { type: "string" },
        range: { type: "string" },
        rules: {
          type: "array",
          items: {
            type: "object",
            properties: {
              type: { type: "string" },
              operator: { type: "string" },
              formulae: { type: "array", items: { type: "string" } },
              style: {
                type: "object",
                properties: {
                  fill: { type: "object" },
                  font: { type: "object" },
                  border: { type: "object" }
                }
              }
            }
          }
        }
      },
      required: ["worksheetName", "range", "rules"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        await excelManager.applyConditionalFormatting(args.worksheetName, args.range, args.rules);
        return {
          success: true,
          message: `Applied conditional formatting to ${args.range}`
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
    name: "excel_data_validation",
    description: "Add data validation to a range",
    inputSchema: {
      type: "object",
      properties: {
        worksheetName: { type: "string" },
        range: { type: "string" },
        validation: {
          type: "object",
          properties: {
            type: { type: "string" },
            operator: { type: "string" },
            formula1: { type: "string" },
            formula2: { type: "string" },
            allowBlank: { type: "boolean" },
            showInputMessage: { type: "boolean" },
            promptTitle: { type: "string" },
            prompt: { type: "string" },
            showErrorMessage: { type: "boolean" },
            errorTitle: { type: "string" },
            error: { type: "string" }
          }
        }
      },
      required: ["worksheetName", "range", "validation"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        await excelManager.validateData(args.worksheetName, args.range, args.validation);
        return {
          success: true,
          message: `Added data validation to ${args.range}`
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
    name: "excel_close_workbook",
    description: "Close the current workbook",
    inputSchema: {
      type: "object",
      properties: {}
    },
    handler: async (): Promise<ToolResult> => {
      try {
        excelManager.closeWorkbook();
        return {
          success: true,
          message: "Workbook closed"
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
    name: "excel_autofit_columns",
    description: "Auto-fit column widths based on content with min/max constraints",
    inputSchema: {
      type: "object",
      properties: {
        worksheetName: { type: "string", description: "Name of worksheet to auto-fit" },
        minWidth: { type: "number", default: 30, description: "Minimum width in pixels" },
        maxWidth: { type: "number", default: 300, description: "Maximum width in pixels" },
        paddingRatio: { type: "number", default: 1.2, description: "Padding multiplier for content width" }
      },
      required: ["worksheetName"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        await excelManager.autoFitColumnWidths(args.worksheetName, {
          minWidth: args.minWidth || 30,
          maxWidth: args.maxWidth || 300,
          paddingRatio: args.paddingRatio || 1.2
        });
        return {
          success: true,
          message: `Auto-fitted columns in ${args.worksheetName}`
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
    name: "excel_autofit_all_columns",
    description: "Auto-fit column widths for all worksheets in the workbook",
    inputSchema: {
      type: "object",
      properties: {
        minWidth: { type: "number", default: 30, description: "Minimum width in pixels" },
        maxWidth: { type: "number", default: 300, description: "Maximum width in pixels" },
        paddingRatio: { type: "number", default: 1.2, description: "Padding multiplier for content width" }
      }
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        await excelManager.autoFitAllColumnWidths({
          minWidth: args.minWidth || 30,
          maxWidth: args.maxWidth || 300,
          paddingRatio: args.paddingRatio || 1.2
        });
        return {
          success: true,
          message: "Auto-fitted columns in all worksheets"
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
    name: "excel_create_formula_reference",
    description: "Create a comprehensive formula reference sheet with accounting standards and links",
    inputSchema: {
      type: "object",
      properties: {
        worksheetName: { type: "string", default: "Formula Reference" }
      }
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const referenceData = FormulaGenerator.createFormulaDocumentationSheet();
        const worksheetName = args.worksheetName || "Formula Reference";
        
        await excelManager.addWorksheet(worksheetName);
        await excelManager.writeWorksheet(worksheetName, referenceData);
        await excelManager.autoFitColumnWidths(worksheetName, { minWidth: 50, maxWidth: 400 });
        
        return {
          success: true,
          message: `Created formula reference sheet: ${worksheetName}`,
          data: { 
            formulaCount: referenceData.length - 6, // Exclude headers and notes
            worksheetName 
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
    name: "excel_create_npv_analysis",
    description: "Create a comprehensive NPV analysis worksheet with formulas and documentation",
    inputSchema: {
      type: "object",
      properties: {
        projectName: { type: "string" },
        discountRate: { type: "number", description: "Annual discount rate (e.g., 0.10 for 10%)" },
        worksheetName: { type: "string", default: "NPV Analysis" }
      },
      required: ["projectName", "discountRate"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const worksheetData = ProfessionalTemplates.createNPVAnalysisWorksheet(args.projectName, args.discountRate);
        const worksheetName = args.worksheetName || "NPV Analysis";
        
        await excelManager.addWorksheet(worksheetName);
        await excelManager.writeWorksheet(worksheetName, worksheetData.data);
        await excelManager.autoFitColumnWidths(worksheetName, { minWidth: 80, maxWidth: 250 });
        
        return {
          success: true,
          message: `Created NPV analysis worksheet for ${args.projectName}`,
          data: { 
            worksheetName,
            discountRate: args.discountRate,
            projectName: args.projectName 
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
    name: "excel_create_loan_analysis",
    description: "Create a loan amortization analysis worksheet with payment schedule and accounting entries",
    inputSchema: {
      type: "object",
      properties: {
        loanAmount: { type: "number", description: "Principal loan amount" },
        annualRate: { type: "number", description: "Annual interest rate (e.g., 0.05 for 5%)" },
        years: { type: "number", description: "Loan term in years" },
        worksheetName: { type: "string", default: "Loan Analysis" }
      },
      required: ["loanAmount", "annualRate", "years"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const worksheetData = ProfessionalTemplates.createLoanAnalysisWorksheet(args.loanAmount, args.annualRate, args.years);
        const worksheetName = args.worksheetName || "Loan Analysis";
        
        await excelManager.addWorksheet(worksheetName);
        await excelManager.writeWorksheet(worksheetName, worksheetData.data);
        await excelManager.autoFitColumnWidths(worksheetName, { minWidth: 80, maxWidth: 200 });
        
        return {
          success: true,
          message: `Created loan analysis for $${args.loanAmount} at ${args.annualRate * 100}% for ${args.years} years`,
          data: { 
            worksheetName,
            loanAmount: args.loanAmount,
            annualRate: args.annualRate,
            years: args.years
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
    name: "excel_create_rent_roll",
    description: "Create a comprehensive rent roll analysis worksheet with occupancy and financial metrics",
    inputSchema: {
      type: "object",
      properties: {
        propertyName: { type: "string" },
        worksheetName: { type: "string", default: "Rent Roll" }
      },
      required: ["propertyName"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const worksheetData = ProfessionalTemplates.createRentRollTemplate(args.propertyName);
        const worksheetName = args.worksheetName || "Rent Roll";
        
        await excelManager.addWorksheet(worksheetName);
        await excelManager.writeWorksheet(worksheetName, worksheetData.data);
        await excelManager.autoFitColumnWidths(worksheetName, { minWidth: 80, maxWidth: 200 });
        
        return {
          success: true,
          message: `Created rent roll analysis for ${args.propertyName}`,
          data: { 
            worksheetName,
            propertyName: args.propertyName
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
    name: "excel_create_financial_ratios",
    description: "Create a financial ratios analysis worksheet with GAAP-compliant ratio calculations",
    inputSchema: {
      type: "object",
      properties: {
        companyName: { type: "string" },
        worksheetName: { type: "string", default: "Financial Ratios" }
      },
      required: ["companyName"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const worksheetData = ProfessionalTemplates.createFinancialRatiosWorksheet(args.companyName);
        const worksheetName = args.worksheetName || "Financial Ratios";
        
        await excelManager.addWorksheet(worksheetName);
        await excelManager.writeWorksheet(worksheetName, worksheetData.data);
        await excelManager.autoFitColumnWidths(worksheetName, { minWidth: 100, maxWidth: 300 });
        
        return {
          success: true,
          message: `Created financial ratios analysis for ${args.companyName}`,
          data: { 
            worksheetName,
            companyName: args.companyName
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
    name: "excel_create_cash_flow_projection",
    description: "Create a 12-month cash flow projection worksheet with operating, investing, and financing activities",
    inputSchema: {
      type: "object",
      properties: {
        entityName: { type: "string" },
        worksheetName: { type: "string", default: "Cash Flow Projection" }
      },
      required: ["entityName"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const worksheetData = ProfessionalTemplates.createCashFlowProjectionWorksheet(args.entityName);
        const worksheetName = args.worksheetName || "Cash Flow Projection";
        
        await excelManager.addWorksheet(worksheetName);
        await excelManager.writeWorksheet(worksheetName, worksheetData.data);
        await excelManager.autoFitColumnWidths(worksheetName, { minWidth: 60, maxWidth: 150 });
        
        return {
          success: true,
          message: `Created cash flow projection for ${args.entityName}`,
          data: { 
            worksheetName,
            entityName: args.entityName
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
    name: "excel_add_formula_with_documentation",
    description: "Add a financial formula to a cell with complete documentation and validation",
    inputSchema: {
      type: "object",
      properties: {
        worksheetName: { type: "string" },
        cell: { type: "string", description: "Cell address (e.g., 'A1')" },
        formulaType: { 
          type: "string",
          enum: [
            "npv", "loan_payment", "straight_line_depreciation", "declining_balance_depreciation", "sum_of_years_depreciation",
            "current_ratio", "quick_ratio", "cash_ratio", "roa", "roe", "debt_to_equity", "debt_to_assets",
            "cap_rate", "noi", "cash_on_cash", "dscr", "working_capital", "gross_margin", "operating_margin", 
            "net_margin", "ebitda", "free_cash_flow", "wacc", "break_even", "inventory_turnover", 
            "days_sales_outstanding", "price_to_earnings", "interest_coverage"
          ]
        },
        parameters: { type: "object", description: "Parameters specific to the formula type" },
        addDocumentation: { type: "boolean", default: true, description: "Add explanation in adjacent cells" }
      },
      required: ["worksheetName", "cell", "formulaType", "parameters"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        let formulaInfo;
        
        switch (args.formulaType) {
          case "npv":
            formulaInfo = FormulaGenerator.generateNPVFormula(
              args.parameters.rate, 
              args.parameters.cashFlowRange, 
              args.parameters.initialInvestment
            );
            break;
            
          case "loan_payment":
            formulaInfo = FormulaGenerator.generateLoanPaymentFormula(
              args.parameters.principal,
              args.parameters.rate,
              args.parameters.periods
            );
            break;
            
          case "straight_line_depreciation":
            formulaInfo = FormulaGenerator.generateDepreciationFormula(
              'straight-line',
              args.parameters.cost,
              args.parameters.salvage,
              args.parameters.life,
              args.parameters.year
            );
            break;
            
          case "declining_balance_depreciation":
            formulaInfo = FormulaGenerator.generateDepreciationFormula(
              'declining-balance',
              args.parameters.cost,
              args.parameters.salvage,
              args.parameters.life,
              args.parameters.year
            );
            break;
            
          case "sum_of_years_depreciation":
            formulaInfo = FormulaGenerator.generateDepreciationFormula(
              'sum-of-years',
              args.parameters.cost,
              args.parameters.salvage,
              args.parameters.life,
              args.parameters.year
            );
            break;
            
          case "current_ratio":
            formulaInfo = FormulaGenerator.generateCurrentRatioFormula(
              args.parameters.currentAssetsRange,
              args.parameters.currentLiabilitiesRange
            );
            break;
            
          case "quick_ratio":
            formulaInfo = FormulaGenerator.generateQuickRatioFormula(
              args.parameters.currentAssetsCell,
              args.parameters.inventoryCell,
              args.parameters.currentLiabilitiesCell
            );
            break;
            
          case "roa":
            formulaInfo = FormulaGenerator.generateROAFormula(
              args.parameters.netIncomeCell,
              args.parameters.totalAssetsCell
            );
            break;
            
          case "roe":
            formulaInfo = FormulaGenerator.generateROEFormula(
              args.parameters.netIncomeCell,
              args.parameters.shareholdersEquityCell
            );
            break;
            
          case "debt_to_equity":
            formulaInfo = FormulaGenerator.generateDebtToEquityFormula(
              args.parameters.totalDebtCell,
              args.parameters.totalEquityCell
            );
            break;
            
          case "cap_rate":
            formulaInfo = FormulaGenerator.generateCapRateFormula(
              args.parameters.noiCell,
              args.parameters.propertyValueCell
            );
            break;
            
          case "noi":
            formulaInfo = FormulaGenerator.generateNOIFormula(
              args.parameters.grossIncomeRange,
              args.parameters.operatingExpensesRange,
              args.parameters.vacancyLossCell
            );
            break;
            
          case "cash_on_cash":
            formulaInfo = FormulaGenerator.generateCashOnCashFormula(
              args.parameters.annualCashFlowCell,
              args.parameters.totalInvestmentCell
            );
            break;
            
          case "dscr":
            formulaInfo = FormulaGenerator.generateDSCRFormula(
              args.parameters.noiCell,
              args.parameters.debtServiceCell
            );
            break;
            
          case "working_capital":
            formulaInfo = FormulaGenerator.generateWorkingCapitalFormula(
              args.parameters.currentAssetsRange,
              args.parameters.currentLiabilitiesRange
            );
            break;
            
          case "gross_margin":
            formulaInfo = FormulaGenerator.generateGrossMarginFormula(
              args.parameters.revenueCell,
              args.parameters.cogsCell
            );
            break;
            
          case "operating_margin":
            formulaInfo = FormulaGenerator.generateOperatingMarginFormula(
              args.parameters.operatingIncomeCell,
              args.parameters.revenueCell
            );
            break;
            
          case "ebitda":
            formulaInfo = FormulaGenerator.generateEBITDAFormula(
              args.parameters.netIncomeCell,
              args.parameters.taxesCell,
              args.parameters.interestCell,
              args.parameters.depreciationCell,
              args.parameters.amortizationCell
            );
            break;
            
          case "free_cash_flow":
            formulaInfo = FormulaGenerator.generateFreeeCashFlowFormula(
              args.parameters.operatingCashFlowCell,
              args.parameters.capexCell
            );
            break;
            
          case "wacc":
            formulaInfo = FormulaGenerator.generateWACCFormula(
              args.parameters.equityValueCell,
              args.parameters.debtValueCell,
              args.parameters.costOfEquityCell,
              args.parameters.costOfDebtCell,
              args.parameters.taxRateCell
            );
            break;
            
          case "break_even":
            formulaInfo = FormulaGenerator.generateBreakEvenFormula(
              args.parameters.fixedCostsCell,
              args.parameters.pricePer,
              args.parameters.variableCostPer
            );
            break;
            
          case "inventory_turnover":
            formulaInfo = FormulaGenerator.generateInventoryTurnoverFormula(
              args.parameters.cogsCell,
              args.parameters.avgInventoryCell
            );
            break;
            
          case "days_sales_outstanding":
            formulaInfo = FormulaGenerator.generateDaysSalesOutstandingFormula(
              args.parameters.accountsReceivableCell,
              args.parameters.annualSalesCell
            );
            break;
            
          case "price_to_earnings":
            formulaInfo = FormulaGenerator.generatePriceToEarningsFormula(
              args.parameters.stockPriceCell,
              args.parameters.epsCell
            );
            break;
            
          default:
            throw new Error(`Formula type ${args.formulaType} not yet implemented`);
        }
        
        // Add the formula to the specified cell
        await excelManager.writeWorksheet(args.worksheetName, [[{
          formula: formulaInfo.formula,
          value: null
        }]], parseInt(args.cell.replace(/[A-Z]/g, '')), 
        args.cell.charCodeAt(0) - 64);
        
        // Add documentation if requested
        if (args.addDocumentation) {
          const docStartRow = parseInt(args.cell.replace(/[A-Z]/g, '')) + 2;
          const docCol = args.cell.charCodeAt(0) - 64;
          
          const documentation = [
            [`Formula: ${formulaInfo.formula}`],
            [`Description: ${formulaInfo.description}`],
            [`Standard: ${formulaInfo.accountingStandard}`],
            [`Reference: ${formulaInfo.referenceUrl}`],
            [`Validation: ${formulaInfo.validation}`],
            [`Parameters:`],
            ...formulaInfo.parameters.map(param => [`  ${param}`])
          ];
          
          await excelManager.writeWorksheet(args.worksheetName, documentation, docStartRow, docCol);
        }
        
        return {
          success: true,
          message: `Added ${args.formulaType} formula to ${args.cell}`,
          data: {
            formula: formulaInfo.formula,
            description: formulaInfo.description,
            reference: formulaInfo.referenceUrl
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
    name: "excel_write_calculation",
    description: "Write a calculation to a cell using a formula (not hardcoded value) for transparency",
    inputSchema: {
      type: "object",
      properties: {
        worksheetName: { type: "string" },
        cell: { type: "string", description: "Cell address (e.g., 'A3')" },
        operation: { 
          type: "string", 
          enum: ["sum", "average", "count", "min", "max", "product", "subtract", "divide", "multiply", "custom"],
          description: "Type of calculation to perform"
        },
        references: {
          type: "array",
          items: { type: "string" },
          description: "Cell references or ranges to include in calculation (e.g., ['A1', 'A2'] or ['A1:A10'])"
        },
        customFormula: {
          type: "string",
          description: "Custom formula if operation is 'custom' (e.g., '=A1*B1+C1')"
        },
        description: {
          type: "string",
          description: "Optional description of what this calculation represents"
        }
      },
      required: ["worksheetName", "cell", "operation"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        let formula: string;
        
        // Build the formula based on the operation
        if (args.operation === 'custom' && args.customFormula) {
          formula = args.customFormula.startsWith('=') ? args.customFormula : `=${args.customFormula}`;
        } else if (args.references && args.references.length > 0) {
          const refs = args.references.join(',');
          
          switch (args.operation) {
            case 'sum':
              formula = `=SUM(${refs})`;
              break;
            case 'average':
              formula = `=AVERAGE(${refs})`;
              break;
            case 'count':
              formula = `=COUNT(${refs})`;
              break;
            case 'min':
              formula = `=MIN(${refs})`;
              break;
            case 'max':
              formula = `=MAX(${refs})`;
              break;
            case 'product':
              formula = `=PRODUCT(${refs})`;
              break;
            case 'subtract':
              if (args.references.length === 2) {
                formula = `=${args.references[0]}-${args.references[1]}`;
              } else {
                throw new Error('Subtract operation requires exactly 2 references');
              }
              break;
            case 'divide':
              if (args.references.length === 2) {
                formula = `=${args.references[0]}/${args.references[1]}`;
              } else {
                throw new Error('Divide operation requires exactly 2 references');
              }
              break;
            case 'multiply':
              if (args.references.length === 2) {
                formula = `=${args.references[0]}*${args.references[1]}`;
              } else if (args.references.length > 2) {
                formula = `=PRODUCT(${refs})`;
              } else {
                throw new Error('Multiply operation requires at least 2 references');
              }
              break;
            default:
              throw new Error(`Unknown operation: ${args.operation}`);
          }
        } else {
          throw new Error('References are required for non-custom operations');
        }
        
        // Parse cell address
        const match = args.cell.match(/^([A-Z]+)(\d+)$/);
        if (!match) {
          throw new Error(`Invalid cell address: ${args.cell}`);
        }
        
        const col = match[1].charCodeAt(0) - 64;
        const row = parseInt(match[2]);
        
        // Write the formula to the cell
        await excelManager.writeCellWithFormula(
          args.worksheetName,
          row,
          col,
          formula,
          args.description
        );
        
        return {
          success: true,
          message: `Added formula ${formula} to cell ${args.cell}`,
          data: {
            cell: args.cell,
            formula: formula,
            description: args.description
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
    name: "excel_validate_formulas",
    description: "Validate that cells contain formulas instead of hardcoded values",
    inputSchema: {
      type: "object",
      properties: {
        worksheetName: { type: "string" },
        cells: {
          type: "array",
          items: { type: "string" },
          description: "List of cell addresses to validate (e.g., ['A3', 'B5', 'C10'])"
        }
      },
      required: ["worksheetName", "cells"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const results: { [key: string]: boolean } = {};
        const missingFormulas: string[] = [];
        
        for (const cellAddr of args.cells) {
          const match = cellAddr.match(/^([A-Z]+)(\d+)$/);
          if (!match) {
            throw new Error(`Invalid cell address: ${cellAddr}`);
          }
          
          const col = match[1].charCodeAt(0) - 64;
          const row = parseInt(match[2]);
          
          const hasFormula = await excelManager.validateCellHasFormula(args.worksheetName, row, col);
          results[cellAddr] = hasFormula;
          
          if (!hasFormula) {
            missingFormulas.push(cellAddr);
          }
        }
        
        return {
          success: true,
          message: missingFormulas.length > 0 
            ? `Warning: ${missingFormulas.length} cells are missing formulas: ${missingFormulas.join(', ')}`
            : 'All cells contain formulas',
          data: {
            results,
            missingFormulas
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
    name: "excel_ensure_formula_calculation",
    description: "Ensure all formulas in a worksheet will be calculated when the file is opened",
    inputSchema: {
      type: "object",
      properties: {
        worksheetName: { type: "string" }
      },
      required: ["worksheetName"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        await excelManager.ensureFormulasInWorksheet(args.worksheetName);
        
        return {
          success: true,
          message: `Configured ${args.worksheetName} to calculate all formulas automatically`
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
    name: "excel_audit_calculations",
    description: "Audit a worksheet to find cells that should contain formulas but have hardcoded values",
    inputSchema: {
      type: "object",
      properties: {
        worksheetName: { type: "string" },
        suspiciousPatterns: {
          type: "array",
          items: { type: "string" },
          default: ["SUM", "AVERAGE", "TOTAL", "SUBTOTAL", "TAX", "DISCOUNT", "PERCENT", "%"],
          description: "Text patterns that suggest a cell should contain a calculation"
        },
        checkNumericCells: {
          type: "boolean",
          default: true,
          description: "Check all numeric cells for potential formulas"
        }
      },
      required: ["worksheetName"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const data = await excelManager.readWorksheet(args.worksheetName);
        const suspiciousPatterns = args.suspiciousPatterns || ["SUM", "AVERAGE", "TOTAL", "SUBTOTAL", "TAX", "DISCOUNT", "PERCENT", "%"];
        const issues: any[] = [];
        
        data.forEach((row, rowIndex) => {
          row.forEach((cell, colIndex) => {
            const cellAddress = String.fromCharCode(65 + colIndex) + (rowIndex + 1);
            
            // Check if it's a numeric value without a formula
            if (typeof cell === 'number' && cell !== 0) {
              // Check if this appears to be in a calculation context
              const nearbyText = [];
              
              // Check cells around this one for suspicious text
              for (let r = Math.max(0, rowIndex - 1); r <= Math.min(data.length - 1, rowIndex + 1); r++) {
                for (let c = Math.max(0, colIndex - 2); c <= Math.min(row.length - 1, colIndex + 2); c++) {
                  if (r !== rowIndex || c !== colIndex) {
                    const nearbyCell = data[r] && data[r][c];
                    if (typeof nearbyCell === 'string') {
                      nearbyText.push(nearbyCell.toUpperCase());
                    }
                  }
                }
              }
              
              // Check if nearby text suggests this should be a formula
              const contextText = nearbyText.join(' ');
              const shouldHaveFormula = suspiciousPatterns.some((pattern: string) => 
                contextText.includes(pattern.toUpperCase())
              );
              
              if (shouldHaveFormula) {
                issues.push({
                  cell: cellAddress,
                  value: cell,
                  issue: 'Numeric value in calculation context - should likely be a formula',
                  context: nearbyText.filter(text => text.length > 0).slice(0, 3),
                  severity: 'WARNING'
                });
              }
            }
            
            // Check for cells with formula objects that have suspicious calculated values
            if (typeof cell === 'object' && cell !== null && !cell.formula && typeof cell.value === 'number') {
              issues.push({
                cell: cellAddress,
                value: cell.value,
                issue: 'Cell object with numeric value but no formula - likely hardcoded calculation',
                severity: 'ERROR'
              });
            }
          });
        });
        
        // Count cells with formulas vs without
        let totalCells = 0;
        let formulaCells = 0;
        let numericCells = 0;
        
        data.forEach(row => {
          row.forEach(cell => {
            if (cell !== null && cell !== undefined && cell !== '') {
              totalCells++;
              if (typeof cell === 'object' && cell.formula) {
                formulaCells++;
              } else if (typeof cell === 'number') {
                numericCells++;
              }
            }
          });
        });
        
        return {
          success: true,
          message: issues.length > 0 
            ? `Found ${issues.length} potential formula issues in ${args.worksheetName}`
            : `No formula issues found in ${args.worksheetName}`,
          data: {
            issues,
            summary: {
              totalCells,
              formulaCells,
              numericCells,
              formulaPercentage: totalCells > 0 ? Math.round((formulaCells / totalCells) * 100) : 0
            },
            recommendations: issues.length > 0 ? [
              'Review flagged cells to determine if they should contain formulas',
              'Use excel_write_calculation to replace hardcoded values with formulas',
              'Ensure all calculated values show their derivation transparently'
            ] : [
              'Worksheet appears to follow formula best practices',
              'All calculations are transparent and auditable'
            ]
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
    name: "excel_enforce_formula_rule",
    description: "Enforce the rule that ALL calculated values must use formulas - never hardcoded results",
    inputSchema: {
      type: "object",
      properties: {
        worksheetName: { type: "string" },
        autoFix: {
          type: "boolean",
          default: false,
          description: "Attempt to automatically convert simple calculations to formulas"
        }
      },
      required: ["worksheetName"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        // Run comprehensive audit
        const data = await excelManager.readWorksheet(args.worksheetName);
        const violations: any[] = [];
        
        data.forEach((row, rowIndex) => {
          row.forEach((cell, colIndex) => {
            const cellAddress = String.fromCharCode(65 + colIndex) + (rowIndex + 1);
            
            // Flag any numeric values that might be calculations
            if (typeof cell === 'number' && cell > 0) {
              // Check if this could be a calculated result
              const isPotentialCalculation = 
                cell % 1 === 0 && cell > 10 || // Round numbers over 10
                cell.toString().includes('.') && cell > 1; // Decimal values over 1
                
              if (isPotentialCalculation) {
                violations.push({
                  cell: cellAddress,
                  value: cell,
                  violation: 'RULE_VIOLATION: Calculated value without formula',
                  rule: 'All calculated values MUST use Excel formulas for transparency',
                  action: 'Convert to formula using excel_write_calculation tool'
                });
              }
            }
          });
        });
        
        return {
          success: true,
          message: violations.length > 0 
            ? `⚠️ RULE VIOLATIONS: Found ${violations.length} cells that violate the formula rule`
            : `✅ COMPLIANCE: All calculations appear to use formulas`,
          data: {
            compliant: violations.length === 0,
            violations,
            rule: {
              title: 'MANDATORY FORMULA RULE',
              description: 'Every calculated value in Excel must use a formula, never hardcoded results',
              rationale: 'Ensures transparency, auditability, and professional standards',
              examples: {
                wrong: 'Cell contains: 42 (hardcoded result)',
                right: 'Cell contains: =SUM(A1:A5) (shows calculation method)'
              }
            },
            nextSteps: violations.length > 0 ? [
              'Use excel_write_calculation to fix violations',
              'Replace hardcoded values with appropriate formulas',
              'Re-run this tool to verify compliance'
            ] : [
              'Worksheet is fully compliant with formula rule',
              'All calculations are transparent and professional'
            ]
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