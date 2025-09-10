const XLSXChart = require('xlsx-chart');

async function testNativeCharts() {
  console.log('🧪 Testing native Excel chart creation...\n');
  
  // Test 1: Basic Column Chart
  console.log('📊 Test 1: Creating basic column chart');
  
  const xlsxChart1 = new XLSXChart();
  const columnChartOptions = {
    file: 'native-column-chart.xlsx',
    chart: 'column',
    titles: ['Q1 2024', 'Q2 2024', 'Q3 2024', 'Q4 2024'],
    fields: ['Revenue', 'Expenses', 'Net Income'],
    data: {
      'Q1 2024': {
        'Revenue': 100000,
        'Expenses': 75000,
        'Net Income': 25000
      },
      'Q2 2024': {
        'Revenue': 120000,
        'Expenses': 80000,
        'Net Income': 40000
      },
      'Q3 2024': {
        'Revenue': 110000,
        'Expenses': 70000,
        'Net Income': 40000
      },
      'Q4 2024': {
        'Revenue': 140000,
        'Expenses': 85000,
        'Net Income': 55000
      }
    }
  };

  return new Promise((resolve, reject) => {
    xlsxChart1.writeFile(columnChartOptions, (err) => {
      if (err) {
        console.log('❌ Column chart failed:', err.message);
        reject(err);
      } else {
        console.log('✅ Created native-column-chart.xlsx');
        
        // Test 2: Line Chart for Trends
        console.log('\n📈 Test 2: Creating line chart for trends');
        
        const xlsxChart2 = new XLSXChart();
        const lineChartOptions = {
          file: 'native-line-chart.xlsx',
          chart: 'line',
          titles: ['Stock A', 'Stock B', 'Stock C'],
          fields: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'],
          data: {
            'Stock A': {
              'Jan': 100, 'Feb': 105, 'Mar': 102, 'Apr': 108, 'May': 115, 'Jun': 120
            },
            'Stock B': {
              'Jan': 80, 'Feb': 85, 'Mar': 83, 'Apr': 90, 'May': 88, 'Jun': 95
            },
            'Stock C': {
              'Jan': 120, 'Feb': 118, 'Mar': 125, 'Apr': 130, 'May': 128, 'Jun': 135
            }
          }
        };

        xlsxChart2.writeFile(lineChartOptions, (err2) => {
          if (err2) {
            console.log('❌ Line chart failed:', err2.message);
            reject(err2);
          } else {
            console.log('✅ Created native-line-chart.xlsx');
            
            // Test 3: Pie Chart for Breakdown
            console.log('\n🥧 Test 3: Creating pie chart for expense breakdown');
            
            const xlsxChart3 = new XLSXChart();
            const pieChartOptions = {
              file: 'native-pie-chart.xlsx',
              chart: 'pie',
              titles: ['Expenses'],
              fields: ['Rent', 'Utilities', 'Marketing', 'Salaries', 'Other'],
              data: {
                'Expenses': {
                  'Rent': 5000,
                  'Utilities': 1200,
                  'Marketing': 3000,
                  'Salaries': 15000,
                  'Other': 2800
                }
              }
            };

            xlsxChart3.writeFile(pieChartOptions, (err3) => {
              if (err3) {
                console.log('❌ Pie chart failed:', err3.message);
                reject(err3);
              } else {
                console.log('✅ Created native-pie-chart.xlsx');
                
                // Test 4: Area Chart for Cash Flow
                console.log('\n📊 Test 4: Creating area chart for cash flow');
                
                const xlsxChart4 = new XLSXChart();
                const areaChartOptions = {
                  file: 'native-area-chart.xlsx',
                  chart: 'area',
                  titles: ['Operating CF', 'Investing CF', 'Financing CF'],
                  fields: ['Q1', 'Q2', 'Q3', 'Q4'],
                  data: {
                    'Operating CF': {
                      'Q1': 50000, 'Q2': 60000, 'Q3': 55000, 'Q4': 70000
                    },
                    'Investing CF': {
                      'Q1': -20000, 'Q2': -15000, 'Q3': -25000, 'Q4': -10000
                    },
                    'Financing CF': {
                      'Q1': 10000, 'Q2': 5000, 'Q3': 15000, 'Q4': -5000
                    }
                  }
                };

                xlsxChart4.writeFile(areaChartOptions, (err4) => {
                  if (err4) {
                    console.log('❌ Area chart failed:', err4.message);
                    reject(err4);
                  } else {
                    console.log('✅ Created native-area-chart.xlsx');
                    resolve({
                      success: true,
                      chartsCreated: 4,
                      files: [
                        'native-column-chart.xlsx',
                        'native-line-chart.xlsx', 
                        'native-pie-chart.xlsx',
                        'native-area-chart.xlsx'
                      ]
                    });
                  }
                });
              }
            });
          }
        });
      }
    });
  });
}

// Run the test
testNativeCharts()
  .then(result => {
    console.log('\n🎉 NATIVE CHART TEST RESULTS:');
    console.log(`   ✅ Success: ${result.success}`);
    console.log(`   📊 Charts Created: ${result.chartsCreated}`);
    console.log('   📁 Files:');
    result.files.forEach(file => console.log(`     - ${file}`));
    
    console.log('\n📄 What you can do with these files:');
    console.log('   1. Open any .xlsx file in Microsoft Excel');
    console.log('   2. You will see NATIVE Excel charts (not images)');
    console.log('   3. Charts are fully interactive and editable in Excel');
    console.log('   4. You can modify chart types, colors, and data');
    console.log('   5. Charts will update if you change the source data');
    console.log('   6. These are real Excel chart objects - fully professional!');
    
    console.log('\n🔧 Chart Types Demonstrated:');
    console.log('   📊 Column Chart: Quarterly financial performance');
    console.log('   📈 Line Chart: Stock price trends over time');  
    console.log('   🥧 Pie Chart: Expense breakdown by category');
    console.log('   📊 Area Chart: Multi-series cash flow analysis');
  })
  .catch(error => {
    console.error('\n❌ TEST FAILED:', error.message);
    console.log('\n🔧 Troubleshooting:');
    console.log('   - Ensure xlsx-chart is properly installed');
    console.log('   - Check that no Excel files are currently open');
    console.log('   - Verify write permissions in the current directory');
  });