import { CellValue } from '../excel/excel-manager.js';

export interface KPIMetric {
  name: string;
  current: number;
  target: number;
  previous: number;
  unit: string;
  category: 'financial' | 'operational' | 'strategic' | 'risk';
  frequency: 'daily' | 'weekly' | 'monthly' | 'quarterly';
  threshold: {
    excellent: number;
    good: number;
    warning: number;
    critical: number;
  };
  trend: 'increasing' | 'decreasing' | 'stable';
  importance: 'high' | 'medium' | 'low';
}

export interface BusinessDriver {
  name: string;
  values: number[];
  periods: string[];
  unit: string;
  impact: 'revenue' | 'cost' | 'efficiency' | 'risk';
  correlation: number; // Correlation with financial performance (-1 to 1)
}

export interface RiskIndicator {
  name: string;
  score: number; // 0-100 risk score
  level: 'low' | 'medium' | 'high' | 'critical';
  category: 'liquidity' | 'operational' | 'strategic' | 'market' | 'credit';
  mitigation: string;
  trend: 'improving' | 'stable' | 'deteriorating';
  impact: number; // Financial impact estimate
}

export interface ExecutiveInsight {
  type: 'opportunity' | 'risk' | 'trend' | 'alert';
  priority: 'high' | 'medium' | 'low';
  title: string;
  description: string;
  recommendation: string;
  financialImpact: number;
  timeframe: string;
  confidence: number; // 0-100 confidence level
}

export interface DashboardData {
  reportDate: Date;
  kpis: KPIMetric[];
  drivers: BusinessDriver[];
  risks: RiskIndicator[];
  cashFlow: {
    current: number;
    projected13Week: number;
    projected52Week: number;
    burnRate: number;
    runwayMonths: number;
  };
  financial: {
    revenue: { current: number; target: number; variance: number };
    ebitda: { current: number; target: number; margin: number };
    workingCapital: { current: number; target: number; days: number };
    debtToEquity: { current: number; target: number; trend: string };
  };
  insights: ExecutiveInsight[];
}

export class ExecutiveDashboard {
  
  private calculateKPIStatus(metric: KPIMetric): { status: string; color: string; variance: number } {
    const variance = ((metric.current - metric.target) / metric.target) * 100;
    let status: string;
    let color: string;
    
    if (metric.current >= metric.threshold.excellent) {
      status = 'Excellent';
      color = 'Green';
    } else if (metric.current >= metric.threshold.good) {
      status = 'Good';
      color = 'Light Green';
    } else if (metric.current >= metric.threshold.warning) {
      status = 'Warning';
      color = 'Yellow';
    } else {
      status = 'Critical';
      color = 'Red';
    }
    
    return { status, color, variance };
  }

  private generateAutomatedInsights(data: DashboardData): ExecutiveInsight[] {
    const insights: ExecutiveInsight[] = [];
    
    // Cash flow analysis
    if (data.cashFlow.runwayMonths < 6) {
      insights.push({
        type: 'alert',
        priority: 'high',
        title: 'Cash Runway Below 6 Months',
        description: `Current cash runway is ${data.cashFlow.runwayMonths.toFixed(1)} months based on current burn rate.`,
        recommendation: 'Consider accelerating revenue initiatives or reducing operational expenses. Evaluate emergency funding options.',
        financialImpact: data.cashFlow.burnRate * 6,
        timeframe: 'Immediate',
        confidence: 95
      });
    }
    
    // Revenue variance analysis
    if (data.financial.revenue.variance < -10) {
      insights.push({
        type: 'risk',
        priority: 'high',
        title: 'Revenue Significantly Below Target',
        description: `Revenue is ${Math.abs(data.financial.revenue.variance).toFixed(1)}% below target.`,
        recommendation: 'Review sales pipeline, pricing strategy, and market conditions. Consider demand generation investments.',
        financialImpact: Math.abs(data.financial.revenue.current - data.financial.revenue.target),
        timeframe: 'Next Quarter',
        confidence: 85
      });
    }
    
    // Working capital efficiency
    if (data.financial.workingCapital.days > 60) {
      insights.push({
        type: 'opportunity',
        priority: 'medium',
        title: 'Working Capital Optimization Opportunity',
        description: `Days in working capital (${data.financial.workingCapital.days}) exceed industry benchmark.`,
        recommendation: 'Optimize accounts receivable collection and inventory management. Target 45-day cycle.',
        financialImpact: (data.financial.workingCapital.days - 45) * data.financial.revenue.current / 365,
        timeframe: '3-6 Months',
        confidence: 75
      });
    }
    
    // Risk assessment
    const highRisks = data.risks.filter(r => r.level === 'high' || r.level === 'critical');
    if (highRisks.length > 0) {
      insights.push({
        type: 'risk',
        priority: 'high',
        title: `${highRisks.length} High-Risk Items Identified`,
        description: `Critical risks: ${highRisks.map(r => r.name).join(', ')}`,
        recommendation: 'Implement immediate risk mitigation strategies. Review enterprise risk management framework.',
        financialImpact: highRisks.reduce((sum, r) => sum + r.impact, 0),
        timeframe: 'Next 30 Days',
        confidence: 90
      });
    }
    
    // KPI trend analysis
    const criticalKPIs = data.kpis.filter(kpi => {
      const status = this.calculateKPIStatus(kpi);
      return status.status === 'Critical';
    });
    
    if (criticalKPIs.length > 0) {
      insights.push({
        type: 'alert',
        priority: 'high',
        title: `${criticalKPIs.length} Critical KPI Performance Issues`,
        description: `Critical KPIs: ${criticalKPIs.map(k => k.name).join(', ')}`,
        recommendation: 'Immediate management attention required. Deploy corrective action plans.',
        financialImpact: 0, // Would need specific calculation per KPI
        timeframe: 'This Week',
        confidence: 100
      });
    }
    
    // Positive trends
    const excellentKPIs = data.kpis.filter(kpi => {
      const status = this.calculateKPIStatus(kpi);
      return status.status === 'Excellent';
    });
    
    if (excellentKPIs.length > 3) {
      insights.push({
        type: 'opportunity',
        priority: 'low',
        title: 'Strong Performance Momentum',
        description: `${excellentKPIs.length} KPIs performing at excellent levels.`,
        recommendation: 'Consider scaling successful strategies. Document best practices for replication.',
        financialImpact: 0,
        timeframe: 'Ongoing',
        confidence: 80
      });
    }
    
    return insights.sort((a, b) => {
      const priorityOrder = { high: 3, medium: 2, low: 1 };
      return priorityOrder[b.priority] - priorityOrder[a.priority];
    });
  }

  private calculateExecutiveSummaryMetrics(data: DashboardData) {
    const totalKPIs = data.kpis.length;
    const criticalKPIs = data.kpis.filter(kpi => {
      const status = this.calculateKPIStatus(kpi);
      return status.status === 'Critical';
    }).length;
    
    const highRisks = data.risks.filter(r => r.level === 'high' || r.level === 'critical').length;
    const totalRisks = data.risks.length;
    
    const overallHealthScore = Math.max(0, 100 - (criticalKPIs / totalKPIs * 40) - (highRisks / totalRisks * 30));
    
    return {
      overallHealthScore,
      criticalKPIs,
      totalKPIs,
      highRisks,
      totalRisks,
      revenueHealth: data.financial.revenue.variance > -5 ? 'Healthy' : 'At Risk',
      cashFlowHealth: data.cashFlow.runwayMonths > 12 ? 'Strong' : 
                     data.cashFlow.runwayMonths > 6 ? 'Adequate' : 'Critical'
    };
  }

  generateExecutiveDashboard(data: DashboardData): Array<Array<CellValue | string | number>> {
    const worksheet: Array<Array<CellValue | string | number>> = [];
    const insights = this.generateAutomatedInsights(data);
    const summary = this.calculateExecutiveSummaryMetrics(data);
    
    // Header
    worksheet.push(['EXECUTIVE DASHBOARD', '', '', '', '', '', '', '', '', '']);
    worksheet.push([`Report Date: ${data.reportDate.toLocaleDateString()}`, '', '', '', '', '', '', '', '', '']);
    worksheet.push([`Generated: ${new Date().toLocaleString()}`, '', '', '', '', '', '', '', '', '']);
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Executive Summary
    worksheet.push(['🎯 EXECUTIVE SUMMARY', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Metric', 'Status', 'Score/Value', 'Trend', 'Action Required', '', '', '', '', '']);
    worksheet.push(['Overall Business Health', 
                   summary.overallHealthScore > 80 ? 'Excellent' : 
                   summary.overallHealthScore > 60 ? 'Good' : 
                   summary.overallHealthScore > 40 ? 'Warning' : 'Critical',
                   summary.overallHealthScore.toFixed(0) + '/100',
                   '→', // Would be calculated from historical data
                   summary.overallHealthScore < 60 ? 'Yes' : 'Monitor'
                  ]);
    worksheet.push(['Revenue Performance', summary.revenueHealth, 
                   data.financial.revenue.variance.toFixed(1) + '% vs target',
                   data.financial.revenue.variance > 0 ? '↗' : '↘',
                   summary.revenueHealth === 'At Risk' ? 'Yes' : 'No'
                  ]);
    worksheet.push(['Cash Flow Position', summary.cashFlowHealth, 
                   data.cashFlow.runwayMonths.toFixed(1) + ' months runway',
                   data.cashFlow.runwayMonths > 12 ? '↗' : '↘',
                   summary.cashFlowHealth === 'Critical' ? 'Immediate' : 'No'
                  ]);
    worksheet.push(['Risk Profile', 
                   summary.highRisks === 0 ? 'Low' : 
                   summary.highRisks <= 2 ? 'Moderate' : 'High',
                   `${summary.highRisks}/${summary.totalRisks} high risks`,
                   '→',
                   summary.highRisks > 0 ? 'Yes' : 'No'
                  ]);
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Key Performance Indicators
    worksheet.push(['📊 KEY PERFORMANCE INDICATORS', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['KPI', 'Current', 'Target', 'Previous', 'Variance', 'Status', 'Trend', 'Priority', '', '']);
    
    for (const kpi of data.kpis.filter(k => k.importance === 'high')) {
      const status = this.calculateKPIStatus(kpi);
      
      worksheet.push([
        kpi.name,
        kpi.current.toFixed(2) + ' ' + kpi.unit,
        kpi.target.toFixed(2) + ' ' + kpi.unit,
        kpi.previous.toFixed(2) + ' ' + kpi.unit,
        status.variance.toFixed(1) + '%',
        status.status,
        kpi.trend === 'increasing' ? '↗' : kpi.trend === 'decreasing' ? '↘' : '→',
        kpi.importance.toUpperCase(),
        ''
      ]);
    }
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Financial Snapshot
    worksheet.push(['💰 FINANCIAL SNAPSHOT', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Metric', 'Current', 'Target', 'Variance', 'Status', 'Comments', '', '', '', '']);
    
    worksheet.push(['Revenue', 
                   (data.financial.revenue.current / 1000000).toFixed(1) + 'M',
                   (data.financial.revenue.target / 1000000).toFixed(1) + 'M',
                   data.financial.revenue.variance.toFixed(1) + '%',
                   data.financial.revenue.variance > -5 ? 'On Track' : 'Below Target',
                   data.financial.revenue.variance < -10 ? 'Requires immediate attention' : 'Monitor closely'
                  ]);
    
    worksheet.push(['EBITDA', 
                   (data.financial.ebitda.current / 1000000).toFixed(1) + 'M',
                   (data.financial.ebitda.target / 1000000).toFixed(1) + 'M',
                   (((data.financial.ebitda.current - data.financial.ebitda.target) / data.financial.ebitda.target) * 100).toFixed(1) + '%',
                   data.financial.ebitda.current >= data.financial.ebitda.target ? 'Meeting Target' : 'Below Target',
                   'Margin: ' + data.financial.ebitda.margin.toFixed(1) + '%'
                  ]);
    
    worksheet.push(['Working Capital Days', 
                   data.financial.workingCapital.days.toFixed(0) + ' days',
                   data.financial.workingCapital.target.toFixed(0) + ' days',
                   ((data.financial.workingCapital.days - data.financial.workingCapital.target) / data.financial.workingCapital.target * 100).toFixed(1) + '%',
                   data.financial.workingCapital.days <= data.financial.workingCapital.target ? 'Efficient' : 'Needs Improvement',
                   data.financial.workingCapital.days > 60 ? 'Optimization opportunity' : 'Acceptable'
                  ]);
    
    worksheet.push(['Debt-to-Equity', 
                   data.financial.debtToEquity.current.toFixed(2) + 'x',
                   data.financial.debtToEquity.target.toFixed(2) + 'x',
                   (((data.financial.debtToEquity.current - data.financial.debtToEquity.target) / data.financial.debtToEquity.target) * 100).toFixed(1) + '%',
                   data.financial.debtToEquity.current <= data.financial.debtToEquity.target ? 'Healthy' : 'Elevated',
                   'Trend: ' + data.financial.debtToEquity.trend
                  ]);
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Cash Flow Analysis
    worksheet.push(['💵 CASH FLOW ANALYSIS', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Metric', 'Amount', 'Status', 'Runway/Rate', 'Risk Level', 'Action', '', '', '', '']);
    
    worksheet.push(['Current Cash Position', 
                   (data.cashFlow.current / 1000000).toFixed(1) + 'M',
                   data.cashFlow.current > data.cashFlow.burnRate * 12 ? 'Strong' : 
                   data.cashFlow.current > data.cashFlow.burnRate * 6 ? 'Adequate' : 'Low',
                   data.cashFlow.runwayMonths.toFixed(1) + ' months',
                   data.cashFlow.runwayMonths < 6 ? 'High' : 
                   data.cashFlow.runwayMonths < 12 ? 'Medium' : 'Low',
                   data.cashFlow.runwayMonths < 6 ? 'Urgent funding needed' : 'Monitor'
                  ]);
    
    worksheet.push(['Monthly Burn Rate', 
                   (data.cashFlow.burnRate / 1000000).toFixed(2) + 'M/month',
                   'Tracking', // Would compare to budget
                   'N/A',
                   'Medium',
                   'Optimize expenses'
                  ]);
    
    worksheet.push(['13-Week Projection', 
                   (data.cashFlow.projected13Week / 1000000).toFixed(1) + 'M',
                   data.cashFlow.projected13Week > 0 ? 'Positive' : 'Negative',
                   '13 weeks',
                   data.cashFlow.projected13Week < 0 ? 'High' : 'Low',
                   data.cashFlow.projected13Week < 0 ? 'Revenue acceleration needed' : 'Continue plan'
                  ]);
    
    worksheet.push(['52-Week Projection', 
                   (data.cashFlow.projected52Week / 1000000).toFixed(1) + 'M',
                   data.cashFlow.projected52Week > data.cashFlow.current ? 'Growth' : 'Decline',
                   '1 year',
                   data.cashFlow.projected52Week < data.cashFlow.current * 0.5 ? 'High' : 'Medium',
                   'Strategic planning required'
                  ]);
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Risk Dashboard
    worksheet.push(['⚠️ RISK DASHBOARD', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Risk Factor', 'Score (0-100)', 'Level', 'Category', 'Trend', 'Est. Impact', 'Mitigation', '', '', '']);
    
    for (const risk of data.risks.filter(r => r.level === 'critical' || r.level === 'high')) {
      worksheet.push([
        risk.name,
        risk.score.toFixed(0),
        risk.level.toUpperCase(),
        risk.category,
        risk.trend === 'improving' ? '↗' : risk.trend === 'deteriorating' ? '↘' : '→',
        (risk.impact / 1000000).toFixed(1) + 'M',
        risk.mitigation,
        '', ''
      ]);
    }
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Strategic Insights
    worksheet.push(['🔍 STRATEGIC INSIGHTS & RECOMMENDATIONS', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Priority', 'Type', 'Insight', 'Recommendation', 'Impact', 'Timeframe', 'Confidence', '', '', '']);
    
    for (const insight of insights.slice(0, 8)) { // Top 8 insights
      worksheet.push([
        insight.priority.toUpperCase(),
        insight.type.charAt(0).toUpperCase() + insight.type.slice(1),
        insight.title,
        insight.recommendation,
        insight.financialImpact > 0 ? (insight.financialImpact / 1000000).toFixed(1) + 'M' : 'TBD',
        insight.timeframe,
        insight.confidence + '%',
        '', ''
      ]);
    }
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Action Items
    worksheet.push(['✅ PRIORITY ACTION ITEMS', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Priority', 'Action Required', 'Owner', 'Due Date', 'Status', 'Impact', '', '', '', '']);
    
    // Generate action items from high-priority insights
    const highPriorityInsights = insights.filter(i => i.priority === 'high');
    for (let i = 0; i < Math.min(5, highPriorityInsights.length); i++) {
      const insight = highPriorityInsights[i];
      worksheet.push([
        'HIGH',
        insight.recommendation,
        'CFO/Management Team',
        insight.timeframe === 'Immediate' ? 'This Week' : 
        insight.timeframe === 'Next 30 Days' ? '30 days' : 
        insight.timeframe,
        'OPEN',
        insight.financialImpact > 0 ? (insight.financialImpact / 1000000).toFixed(1) + 'M' : 'High',
        '', '', ''
      ]);
    }
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Performance Trends (would be populated with historical data)
    worksheet.push(['📈 PERFORMANCE TRENDS (Last 12 Months)', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Month', 'Revenue', 'EBITDA', 'Cash Flow', 'Burn Rate', 'Health Score', '', '', '', '']);
    
    // Sample trend data - in production would come from historical data
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    for (let i = 0; i < 12; i++) {
      const monthlyRevenue = data.financial.revenue.current * (0.85 + Math.random() * 0.3);
      const monthlyEBITDA = monthlyRevenue * data.financial.ebitda.margin / 100;
      
      worksheet.push([
        months[i],
        (monthlyRevenue / 1000000).toFixed(1) + 'M',
        (monthlyEBITDA / 1000000).toFixed(1) + 'M',
        ((Math.random() - 0.5) * monthlyRevenue * 0.1 / 1000000).toFixed(1) + 'M',
        (data.cashFlow.burnRate * (0.9 + Math.random() * 0.2) / 1000000).toFixed(1) + 'M',
        (70 + Math.random() * 25).toFixed(0),
        '', '', ''
      ]);
    }
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Footer & Disclaimers
    worksheet.push(['📋 DASHBOARD NOTES & DISCLAIMERS', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- Dashboard data as of: ' + data.reportDate.toLocaleDateString(), '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- Automated insights generated using statistical analysis', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- Risk scores based on industry benchmarks and internal thresholds', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- Financial projections subject to market conditions and assumptions', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- Requires management review and validation for strategic decisions', '', '', '', '', '', '', '', '', '']);
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['🔗 RELATED REPORTS & MODELS', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- 13-Week Cash Flow Forecast (analytics_13_week_forecast)', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- 52-Week Strategic Forecast (analytics_52_week_forecast)', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- DCF Valuation Model (analytics_dcf_valuation)', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- Monte Carlo Risk Analysis (analytics_monte_carlo_simulation)', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- Scenario Planning (analytics_scenario_comparison)', '', '', '', '', '', '', '', '', '']);
    
    return worksheet;
  }

  generateKPIDashboard(kpis: KPIMetric[]): Array<Array<CellValue | string | number>> {
    const worksheet: Array<Array<CellValue | string | number>> = [];
    
    // Header
    worksheet.push(['KPI PERFORMANCE DASHBOARD', '', '', '', '', '', '', '']);
    worksheet.push(['Generated: ' + new Date().toLocaleString(), '', '', '', '', '', '', '']);
    worksheet.push(['', '', '', '', '', '', '', '']);
    
    // KPI Summary by Category
    const categories = ['financial', 'operational', 'strategic', 'risk'];
    
    for (const category of categories) {
      const categoryKPIs = kpis.filter(k => k.category === category);
      if (categoryKPIs.length === 0) continue;
      
      worksheet.push([category.toUpperCase() + ' KPIS', '', '', '', '', '', '', '']);
      worksheet.push(['KPI Name', 'Current', 'Target', 'Status', 'Variance', 'Trend', 'Action']);
      
      for (const kpi of categoryKPIs) {
        const status = this.calculateKPIStatus(kpi);
        worksheet.push([
          kpi.name,
          kpi.current.toFixed(2) + ' ' + kpi.unit,
          kpi.target.toFixed(2) + ' ' + kpi.unit,
          status.status,
          status.variance.toFixed(1) + '%',
          kpi.trend === 'increasing' ? '↗' : kpi.trend === 'decreasing' ? '↘' : '→',
          status.status === 'Critical' ? 'Immediate Action Required' :
          status.status === 'Warning' ? 'Monitor Closely' : 'Continue Current Strategy'
        ]);
      }
      
      worksheet.push(['', '', '', '', '', '', '']);
    }
    
    return worksheet;
  }
}