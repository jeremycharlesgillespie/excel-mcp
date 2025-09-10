import { ExcelManager } from './src/excel/excel-manager.js';

async function testFormulaExample() {
  const manager = new ExcelManager();
  
  // Create a new workbook with sample data
  await manager.createWorkbook([{
    name: 'Sheet1',
    data: [
      ['Item', 'Quantity', 'Price', 'Total'],
      ['Product A', 2, 10, ExcelManager.createFormulaCell('=B2*C2', 'Quantity * Price')],
      ['Product B', 3, 15, ExcelManager.createFormulaCell('=B3*C3', 'Quantity * Price')],
      ['Product C', 5, 8, ExcelManager.createFormulaCell('=B4*C4', 'Quantity * Price')],
      ['', '', 'Subtotal:', ExcelManager.createFormulaCell('=SUM(D2:D4)', 'Sum of all totals')],
      ['', '', 'Tax (10%):', ExcelManager.createFormulaCell('=D5*0.1', '10% of subtotal')],
      ['', '', 'Grand Total:', ExcelManager.createFormulaCell('=D5+D6', 'Subtotal + Tax')]
    ]
  }]);
  
  // Save the workbook
  await manager.saveWorkbook('formula-example.xlsx');
  
  console.log('Created formula-example.xlsx with formulas that will:');
  console.log('1. Calculate totals for each item (Quantity * Price)');
  console.log('2. Sum all totals for subtotal');
  console.log('3. Calculate tax as 10% of subtotal');
  console.log('4. Calculate grand total as subtotal + tax');
  console.log('');
  console.log('When you open the file:');
  console.log('- Cells will show calculated values');
  console.log('- Clicking on a cell will show the formula in the formula bar');
  console.log('- All calculations are transparent and auditable');
}

// Example of validating formulas
async function validateFormulas() {
  const manager = new ExcelManager();
  await manager.openWorkbook('formula-example.xlsx');
  
  // Validate that calculation cells contain formulas
  const hasFormula1 = await manager.validateCellHasFormula('Sheet1', 2, 4); // D2
  const hasFormula2 = await manager.validateCellHasFormula('Sheet1', 5, 4); // D5
  const hasFormula3 = await manager.validateCellHasFormula('Sheet1', 7, 4); // D7
  
  console.log('Formula validation results:');
  console.log(`D2 has formula: ${hasFormula1}`);
  console.log(`D5 has formula: ${hasFormula2}`);
  console.log(`D7 has formula: ${hasFormula3}`);
}

// Run the test
testFormulaExample()
  .then(() => validateFormulas())
  .catch(console.error);