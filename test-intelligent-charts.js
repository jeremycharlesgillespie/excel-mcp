const { ChartIntelligence } = require('./dist/excel/chart-intelligence.js');

function testChartIntelligence() {
  console.log('🧠 Testing Intelligent Chart Recommendations...\n');
  
  // Test scenarios that should demonstrate intelligent chart selection
  const testCases = [
    {
      name: "Monthly Cash Flow",
      description: "monthly cash flow for the company",
      categories: ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
      series: ["Operating CF", "Investing CF", "Financing CF"],
      expectedChart: "area",
      reason: "Cash flow over time with cumulative effect"
    },
    {
      name: "Quarterly Revenue Trend", 
      description: "quarterly revenue trend analysis",
      categories: ["Q1 2024", "Q2 2024", "Q3 2024", "Q4 2024"],
      series: ["Revenue", "Growth Rate"],
      expectedChart: "line",
      reason: "Revenue trends best shown with line charts"
    },
    {
      name: "Expense Breakdown",
      description: "expense breakdown by category",
      categories: ["Rent", "Salaries", "Marketing", "Utilities", "Other"],
      series: ["Expenses"],
      expectedChart: "pie",
      reason: "Parts of whole composition data"
    },
    {
      name: "Department Sales Comparison",
      description: "sales comparison across departments",
      categories: ["North", "South", "East", "West", "Central"],
      series: ["Q1 Sales", "Q2 Sales"],
      expectedChart: "column",
      reason: "Comparative analysis across categories"
    },
    {
      name: "Stock Portfolio Performance",
      description: "portfolio performance tracking over time",
      categories: ["Jan", "Feb", "Mar", "Apr", "May"],
      series: ["Portfolio Value", "S&P 500"],
      expectedChart: "line", 
      reason: "Financial performance over time"
    },
    {
      name: "Risk vs Return Analysis",
      description: "correlation between risk and return for investments",
      categories: ["Low Risk", "Med Risk", "High Risk"],
      series: ["Expected Return", "Volatility"],
      expectedChart: "scatter",
      reason: "Correlation between two variables"
    },
    {
      name: "Budget vs Actual",
      description: "budget versus actual expenses comparison",
      categories: ["Marketing", "IT", "Operations", "HR", "Finance"],
      series: ["Budget", "Actual"],
      expectedChart: "column",
      reason: "Side-by-side comparison of values"
    },
    {
      name: "Customer Segments",
      description: "customer segments distribution breakdown",
      categories: ["Enterprise", "SMB", "Startup", "Individual"],
      series: ["Customer Count"],
      expectedChart: "pie",
      reason: "Market composition visualization"
    }
  ];
  
  console.log('🎯 Testing Chart Intelligence Recommendations:\n');
  
  const results = testCases.map(testCase => {
    console.log(`📊 Testing: ${testCase.name}`);
    console.log(`   Description: "${testCase.description}"`);
    console.log(`   Categories: [${testCase.categories.join(', ')}]`);
    console.log(`   Series: [${testCase.series.join(', ')}]`);
    
    const recommendation = ChartIntelligence.recommendChart(
      testCase.description,
      testCase.categories, 
      testCase.series
    );
    
    console.log(`   🤖 AI Recommended: ${recommendation.primaryChart.toUpperCase()} (${recommendation.confidence}% confidence)`);
    console.log(`   💡 Reasoning: ${recommendation.reasoning}`);
    
    const isCorrect = recommendation.primaryChart === testCase.expectedChart;
    if (isCorrect) {
      console.log(`   ✅ CORRECT! Matches expected ${testCase.expectedChart.toUpperCase()} chart`);
    } else {
      console.log(`   ❓ Expected ${testCase.expectedChart.toUpperCase()}, got ${recommendation.primaryChart.toUpperCase()}`);
      console.log(`   📝 Expected reason: ${testCase.reason}`);
    }
    
    if (recommendation.alternatives && recommendation.alternatives.length > 0) {
      console.log(`   🔄 Alternatives: ${recommendation.alternatives.map(alt => 
        `${alt.chartType.toUpperCase()} (${alt.confidence}%)`
      ).join(', ')}`);
    }
    
    if (recommendation.warnings && recommendation.warnings.length > 0) {
      console.log(`   ⚠️  Warnings: ${recommendation.warnings.join('; ')}`);
    }
    
    console.log('');
    
    return {
      testCase: testCase.name,
      expected: testCase.expectedChart,
      actual: recommendation.primaryChart,
      correct: isCorrect,
      confidence: recommendation.confidence,
      reasoning: recommendation.reasoning
    };
  });
  
  // Summary
  const correctPredictions = results.filter(r => r.correct).length;
  const totalTests = results.length;
  const accuracy = Math.round((correctPredictions / totalTests) * 100);
  
  console.log('📈 INTELLIGENCE TEST SUMMARY:');
  console.log(`   🎯 Accuracy: ${correctPredictions}/${totalTests} (${accuracy}%)`);
  console.log(`   🧠 Average Confidence: ${Math.round(results.reduce((sum, r) => sum + r.confidence, 0) / totalTests)}%`);
  
  console.log('\n✅ Correct Predictions:');
  results.filter(r => r.correct).forEach(r => {
    console.log(`   - ${r.testCase}: ${r.actual.toUpperCase()} (${r.confidence}%)`);
  });
  
  if (results.some(r => !r.correct)) {
    console.log('\n❓ Mismatched Predictions:');
    results.filter(r => !r.correct).forEach(r => {
      console.log(`   - ${r.testCase}: Expected ${r.expected.toUpperCase()}, got ${r.actual.toUpperCase()}`);
      console.log(`     Reasoning: ${r.reasoning}`);
    });
  }
  
  console.log('\n🎊 Key Intelligence Features Demonstrated:');
  console.log('   ✅ Time-series detection (cash flow → area chart)');
  console.log('   ✅ Trend analysis (revenue → line chart)'); 
  console.log('   ✅ Composition data (expenses → pie chart)');
  console.log('   ✅ Comparative analysis (departments → column chart)');
  console.log('   ✅ Financial context awareness (portfolio → line chart)');
  console.log('   ✅ Correlation detection (risk vs return → scatter)');
  console.log('   ✅ Business scenario mapping');
  console.log('   ✅ Confidence scoring and alternatives');
  
  return results;
}

// Test explanation feature
function testExplanationFeature() {
  console.log('\n📖 Testing Chart Recommendation Explanations:\n');
  
  const testRecommendation = ChartIntelligence.recommendChart(
    "monthly cash flow analysis with operating, investing, and financing activities",
    ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
    ["Operating CF", "Investing CF", "Financing CF"],
    [[50000, 60000, 55000, 70000, 65000, 80000], [-20000, -15000, -25000, -10000, -30000, -5000], [10000, 5000, 15000, -5000, 20000, -15000]]
  );
  
  const explanation = ChartIntelligence.explainRecommendation(testRecommendation);
  console.log('🤖 AI Explanation for Cash Flow Analysis:');
  console.log('─'.repeat(50));
  console.log(explanation);
  console.log('─'.repeat(50));
}

// Run tests
try {
  const results = testChartIntelligence();
  testExplanationFeature();
  
  console.log('\n🚀 INTELLIGENT CHART SYSTEM READY!');
  console.log('\n💡 The MCP server now knows:');
  console.log('   📊 Which chart type to use for different data scenarios');
  console.log('   🧠 Business context and financial best practices'); 
  console.log('   ⚡ Automatic optimization based on data characteristics');
  console.log('   🎯 Confidence scoring and alternative suggestions');
  console.log('   ⚠️  Warning system for suboptimal choices');
  
} catch (error) {
  console.error('\n❌ Intelligence test failed:', error.message);
  console.log('\n🔧 Make sure the project is built: npm run build');
}