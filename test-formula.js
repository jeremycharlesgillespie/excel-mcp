const { ExcelManager } = require('./dist/excel/excel-manager.js');

async function testFormulaExample() {
  const manager = new ExcelManager();
  
  // Create a new workbook with sample data
  await manager.createWorkbook([{
    name: 'Sheet1',
    data: [
      ['Item', 'Quantity', 'Price', 'Total'],
      ['Product A', 2, 10, { formula: '=B2*C2', value: null }],
      ['Product B', 3, 15, { formula: '=B3*C3', value: null }],
      ['Product C', 5, 8, { formula: '=B4*C4', value: null }],
      ['', '', 'Subtotal:', { formula: '=SUM(D2:D4)', value: null }],
      ['', '', 'Tax (10%):', { formula: '=D5*0.1', value: null }],
      ['', '', 'Grand Total:', { formula: '=D5+D6', value: null }]
    ]
  }]);
  
  // Save the workbook
  await manager.saveWorkbook('formula-example.xlsx');
  
  console.log('✅ Created formula-example.xlsx with formulas that will:');
  console.log('   1. Calculate totals for each item (Quantity * Price)');
  console.log('   2. Sum all totals for subtotal');
  console.log('   3. Calculate tax as 10% of subtotal');
  console.log('   4. Calculate grand total as subtotal + tax');
  console.log('');
  console.log('📊 When you open the file in Excel:');
  console.log('   - Cells will show calculated values');
  console.log('   - Clicking on a cell will show the formula in the formula bar');
  console.log('   - All calculations are transparent and auditable');
  console.log('');
  console.log('Example: Cell D2 will show "20" but formula bar will show "=B2*C2"');
}

// Run the test
testFormulaExample()
  .then(() => console.log('\n✅ Test completed successfully!'))
  .catch(error => console.error('❌ Error:', error));