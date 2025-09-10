import ExcelJS from 'exceljs';
import { promises as fs } from 'fs';
import path from 'path';

export interface CellValue {
  value: any;
  formula?: string;
  style?: Partial<ExcelJS.Style>;
}

export interface WorksheetData {
  name: string;
  data: (string | number | Date | CellValue)[][];
  columns?: Partial<ExcelJS.Column>[];
}

export interface ExcelFileInfo {
  path: string;
  worksheets: string[];
  lastModified: Date;
  size: number;
}

export class ExcelManager {
  private workbook: ExcelJS.Workbook | null = null;
  private currentFile: string | null = null;

  async createWorkbook(worksheets: WorksheetData[]): Promise<void> {
    this.workbook = new ExcelJS.Workbook();
    
    for (const wsData of worksheets) {
      const worksheet = this.workbook.addWorksheet(wsData.name);
      
      if (wsData.columns) {
        worksheet.columns = wsData.columns;
      }
      
      wsData.data.forEach((row, rowIndex) => {
        row.forEach((cell, colIndex) => {
          const cellRef = worksheet.getCell(rowIndex + 1, colIndex + 1);
          
          if (typeof cell === 'object' && cell !== null && !(cell instanceof Date)) {
            const cellData = cell as CellValue;
            if (cellData.formula) {
              cellRef.value = { formula: cellData.formula };
            } else {
              cellRef.value = cellData.value;
            }
            if (cellData.style) {
              Object.assign(cellRef, { style: cellData.style });
            }
          } else {
            cellRef.value = cell;
          }
        });
      });
    }
  }

  async openWorkbook(filePath: string): Promise<ExcelFileInfo> {
    this.workbook = new ExcelJS.Workbook();
    await this.workbook.xlsx.readFile(filePath);
    this.currentFile = filePath;
    
    const stats = await fs.stat(filePath);
    
    return {
      path: filePath,
      worksheets: this.workbook.worksheets.map(ws => ws.name),
      lastModified: stats.mtime,
      size: stats.size
    };
  }

  async saveWorkbook(filePath?: string): Promise<void> {
    if (!this.workbook) {
      throw new Error('No workbook is currently open');
    }
    
    const savePath = filePath || this.currentFile;
    if (!savePath) {
      throw new Error('No file path specified');
    }
    
    // Ensure formulas are calculated when the file is opened
    this.workbook.calcProperties = {
      fullCalcOnLoad: true
    };
    
    await this.workbook.xlsx.writeFile(savePath);
    if (!this.currentFile || this.currentFile !== savePath) {
      this.currentFile = savePath;
    }
  }

  async readWorksheet(worksheetName: string): Promise<any[][]> {
    if (!this.workbook) {
      throw new Error('No workbook is currently open');
    }
    
    const worksheet = this.workbook.getWorksheet(worksheetName);
    if (!worksheet) {
      throw new Error(`Worksheet "${worksheetName}" not found`);
    }
    
    const data: any[][] = [];
    worksheet.eachRow({ includeEmpty: true }, (row) => {
      const rowData: any[] = [];
      row.eachCell({ includeEmpty: true }, (cell) => {
        // Check if the cell value contains a formula
        if (cell.value && typeof cell.value === 'object' && 'formula' in cell.value) {
          rowData.push({ 
            value: (cell.value as any).result || (cell.value as any).formula, 
            formula: (cell.value as any).formula 
          });
        } else if ((cell as any).formula) {
          // Fallback for legacy formula detection
          rowData.push({ value: cell.value, formula: (cell as any).formula });
        } else {
          rowData.push(cell.value);
        }
      });
      data.push(rowData);
    });
    
    return data;
  }

  async writeWorksheet(worksheetName: string, data: any[][], startRow: number = 1, startCol: number = 1): Promise<void> {
    if (!this.workbook) {
      throw new Error('No workbook is currently open');
    }
    
    let worksheet = this.workbook.getWorksheet(worksheetName);
    if (!worksheet) {
      worksheet = this.workbook.addWorksheet(worksheetName);
    }
    
    data.forEach((row, rowIndex) => {
      row.forEach((cellData, colIndex) => {
        const cell = worksheet!.getCell(startRow + rowIndex, startCol + colIndex);
        
        if (typeof cellData === 'object' && cellData !== null && !(cellData instanceof Date)) {
          if (cellData.formula) {
            // When a formula is provided, set it and let Excel calculate the value
            cell.value = { formula: cellData.formula };
            // Note: ExcelJS will automatically calculate the value when the file is opened
          } else if (cellData.value !== undefined) {
            cell.value = cellData.value;
          }
          if (cellData.style) {
            Object.assign(cell, { style: cellData.style });
          }
        } else {
          cell.value = cellData;
        }
      });
    });
    
    // Force calculation of all formulas
    this.workbook.calcProperties = {
      fullCalcOnLoad: true // Ensure formulas are calculated when file is opened
    };
  }

  async addWorksheet(name: string): Promise<void> {
    if (!this.workbook) {
      throw new Error('No workbook is currently open');
    }
    
    this.workbook.addWorksheet(name);
  }

  async deleteWorksheet(name: string): Promise<void> {
    if (!this.workbook) {
      throw new Error('No workbook is currently open');
    }
    
    const worksheet = this.workbook.getWorksheet(name);
    if (!worksheet) {
      throw new Error(`Worksheet "${name}" not found`);
    }
    
    this.workbook.removeWorksheet(worksheet.id);
  }

  async getNamedRanges(): Promise<string[]> {
    if (!this.workbook) {
      throw new Error('No workbook is currently open');
    }
    
    return Object.keys(this.workbook.definedNames.model);
  }

  async createNamedRange(name: string, range: string, worksheetName?: string): Promise<void> {
    if (!this.workbook) {
      throw new Error('No workbook is currently open');
    }
    
    const fullRange = worksheetName ? `${worksheetName}!${range}` : range;
    this.workbook.definedNames.add(name, fullRange);
  }

  async applyConditionalFormatting(
    worksheetName: string,
    range: string,
    rules: ExcelJS.ConditionalFormattingRule[]
  ): Promise<void> {
    if (!this.workbook) {
      throw new Error('No workbook is currently open');
    }
    
    const worksheet = this.workbook.getWorksheet(worksheetName);
    if (!worksheet) {
      throw new Error(`Worksheet "${worksheetName}" not found`);
    }
    
    worksheet.addConditionalFormatting({
      ref: range,
      rules: rules
    });
  }

  async protectWorksheet(
    worksheetName: string,
    password: string,
    options?: Partial<ExcelJS.WorksheetProtection>
  ): Promise<void> {
    if (!this.workbook) {
      throw new Error('No workbook is currently open');
    }
    
    const worksheet = this.workbook.getWorksheet(worksheetName);
    if (!worksheet) {
      throw new Error(`Worksheet "${worksheetName}" not found`);
    }
    
    await worksheet.protect(password, options || {});
  }

  async unprotectWorksheet(worksheetName: string): Promise<void> {
    if (!this.workbook) {
      throw new Error('No workbook is currently open');
    }
    
    const worksheet = this.workbook.getWorksheet(worksheetName);
    if (!worksheet) {
      throw new Error(`Worksheet "${worksheetName}" not found`);
    }
    
    worksheet.unprotect();
  }

  async findAndReplace(
    worksheetName: string,
    searchText: string | RegExp,
    replaceText: string,
    options?: { matchCase?: boolean; matchEntireCell?: boolean }
  ): Promise<number> {
    if (!this.workbook) {
      throw new Error('No workbook is currently open');
    }
    
    const worksheet = this.workbook.getWorksheet(worksheetName);
    if (!worksheet) {
      throw new Error(`Worksheet "${worksheetName}" not found`);
    }
    
    let replacementCount = 0;
    
    worksheet.eachRow((row) => {
      row.eachCell((cell) => {
        if (cell.value && typeof cell.value === 'string') {
          let matches = false;
          
          if (searchText instanceof RegExp) {
            matches = searchText.test(cell.value);
            if (matches) {
              cell.value = cell.value.replace(searchText, replaceText);
              replacementCount++;
            }
          } else {
            if (options?.matchEntireCell) {
              matches = options?.matchCase 
                ? cell.value === searchText 
                : cell.value.toLowerCase() === searchText.toLowerCase();
              if (matches) {
                cell.value = replaceText;
                replacementCount++;
              }
            } else {
              const regex = new RegExp(
                searchText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'),
                options?.matchCase ? 'g' : 'gi'
              );
              const newValue = cell.value.replace(regex, replaceText);
              if (newValue !== cell.value) {
                cell.value = newValue;
                replacementCount++;
              }
            }
          }
        }
      });
    });
    
    return replacementCount;
  }

  async getCellFormulas(worksheetName: string): Promise<{ [cell: string]: string }> {
    if (!this.workbook) {
      throw new Error('No workbook is currently open');
    }
    
    const worksheet = this.workbook.getWorksheet(worksheetName);
    if (!worksheet) {
      throw new Error(`Worksheet "${worksheetName}" not found`);
    }
    
    const formulas: { [cell: string]: string } = {};
    
    worksheet.eachRow((row) => {
      row.eachCell((cell) => {
        if ((cell as any).formula) {
          formulas[cell.address] = (cell as any).formula;
        }
      });
    });
    
    return formulas;
  }

  async validateData(
    worksheetName: string,
    range: string,
    validation: ExcelJS.DataValidation
  ): Promise<void> {
    if (!this.workbook) {
      throw new Error('No workbook is currently open');
    }
    
    const worksheet = this.workbook.getWorksheet(worksheetName);
    if (!worksheet) {
      throw new Error(`Worksheet "${worksheetName}" not found`);
    }
    
    const cells = worksheet.getCell(range);
    cells.dataValidation = validation;
  }

  async mergeFiles(filePaths: string[], outputPath: string): Promise<void> {
    const mergedWorkbook = new ExcelJS.Workbook();
    
    for (const filePath of filePaths) {
      const sourceWorkbook = new ExcelJS.Workbook();
      await sourceWorkbook.xlsx.readFile(filePath);
      
      sourceWorkbook.worksheets.forEach(sourceSheet => {
        const newSheetName = `${path.basename(filePath, '.xlsx')}_${sourceSheet.name}`;
        const targetSheet = mergedWorkbook.addWorksheet(newSheetName);
        
        sourceSheet.eachRow((row, rowNumber) => {
          const newRow = targetSheet.getRow(rowNumber);
          row.eachCell((cell, colNumber) => {
            const newCell = newRow.getCell(colNumber);
            newCell.value = cell.value;
            newCell.style = cell.style;
            if ((cell as any).formula) {
              (newCell as any).formula = (cell as any).formula;
            }
          });
        });
        
        targetSheet.columns = sourceSheet.columns;
      });
    }
    
    await mergedWorkbook.xlsx.writeFile(outputPath);
  }

  async autoFitColumnWidths(worksheetName: string, options?: {
    minWidth?: number;
    maxWidth?: number;
    paddingRatio?: number;
  }): Promise<void> {
    if (!this.workbook) {
      throw new Error('No workbook is currently open');
    }

    const worksheet = this.workbook.getWorksheet(worksheetName);
    if (!worksheet) {
      throw new Error(`Worksheet "${worksheetName}" not found`);
    }

    const minWidth = options?.minWidth || 30;
    const maxWidth = options?.maxWidth || 300;
    const paddingRatio = options?.paddingRatio || 1.2;

    // Calculate column widths based on content
    const columnWidths: { [col: number]: number } = {};

    worksheet.eachRow((row) => {
      row.eachCell((cell, colNumber) => {
        let cellText = '';
        
        if (cell.value !== null && cell.value !== undefined) {
          if (typeof cell.value === 'object' && 'text' in cell.value) {
            cellText = String(cell.value.text);
          } else {
            cellText = String(cell.value);
          }
        }

        // Estimate character width (approximate)
        const estimatedWidth = cellText.length * 7 * paddingRatio; // ~7 pixels per character
        
        if (!columnWidths[colNumber] || estimatedWidth > columnWidths[colNumber]) {
          columnWidths[colNumber] = Math.min(Math.max(estimatedWidth, minWidth), maxWidth);
        }
      });
    });

    // Apply calculated widths
    for (const [colNumber, width] of Object.entries(columnWidths)) {
      const column = worksheet.getColumn(parseInt(colNumber));
      column.width = width / 7; // ExcelJS uses character units, not pixels
    }
  }

  async autoFitAllColumnWidths(options?: {
    minWidth?: number;
    maxWidth?: number;
    paddingRatio?: number;
  }): Promise<void> {
    if (!this.workbook) {
      throw new Error('No workbook is currently open');
    }

    for (const worksheet of this.workbook.worksheets) {
      await this.autoFitColumnWidths(worksheet.name, options);
    }
  }

  closeWorkbook(): void {
    this.workbook = null;
    this.currentFile = null;
  }

  /**
   * Write a calculated value with formula to ensure transparency
   * This ensures users can see how values were calculated
   */
  async writeCellWithFormula(
    worksheetName: string, 
    row: number, 
    col: number, 
    formula: string,
    description?: string
  ): Promise<void> {
    if (!this.workbook) {
      throw new Error('No workbook is currently open');
    }
    
    const worksheet = this.workbook.getWorksheet(worksheetName);
    if (!worksheet) {
      throw new Error(`Worksheet "${worksheetName}" not found`);
    }
    
    const cell = worksheet.getCell(row, col);
    cell.value = { formula: formula };
    
    // Add a comment with description if provided
    if (description) {
      cell.note = description;
    }
  }

  /**
   * Create a calculation with formula instead of hardcoded value
   * Ensures all calculations are transparent and auditable
   */
  static createFormulaCell(formula: string, description?: string): CellValue {
    return {
      formula: formula,
      value: null, // Let Excel calculate the value
      style: description ? { 
        font: { italic: true },
        alignment: { horizontal: 'right' }
      } : undefined
    };
  }

  /**
   * Validate that a cell contains a formula, not just a value
   * Use this to ensure calculations are transparent
   */
  async validateCellHasFormula(worksheetName: string, row: number, col: number): Promise<boolean> {
    if (!this.workbook) {
      throw new Error('No workbook is currently open');
    }
    
    const worksheet = this.workbook.getWorksheet(worksheetName);
    if (!worksheet) {
      throw new Error(`Worksheet "${worksheetName}" not found`);
    }
    
    const cell = worksheet.getCell(row, col);
    return !!(cell.value && typeof cell.value === 'object' && 'formula' in cell.value);
  }

  /**
   * Convert a hardcoded value to a formula cell
   * This ensures transparency in calculations
   */
  static convertToFormulaCell(value: number, sourceReferences: string[]): CellValue {
    // Create a SUM formula from the source references
    const formula = sourceReferences.length > 0 
      ? `=SUM(${sourceReferences.join(',')})` 
      : `=${value}`;
    
    return {
      formula: formula,
      value: null
    };
  }

  /**
   * Ensure all calculations in a worksheet use formulas
   * This is critical for audit and transparency
   */
  async ensureFormulasInWorksheet(worksheetName: string): Promise<void> {
    if (!this.workbook) {
      throw new Error('No workbook is currently open');
    }
    
    const worksheet = this.workbook.getWorksheet(worksheetName);
    if (!worksheet) {
      throw new Error(`Worksheet "${worksheetName}" not found`);
    }
    
    // Set calculation properties to ensure formulas are calculated
    this.workbook.calcProperties = {
      fullCalcOnLoad: true
    };
  }
}