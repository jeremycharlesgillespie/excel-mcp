// Chart Intelligence Engine for Excel MCP Server
// Automatically recommends the best chart type based on data characteristics and business context

export interface DataCharacteristics {
  dataType: 'time-series' | 'categorical' | 'comparative' | 'composition' | 'correlation' | 'distribution';
  valueCount: number;
  seriesCount: number;
  hasTimeComponent: boolean;
  hasNegativeValues: boolean;
  businessContext?: string;
  dataPattern?: 'trend' | 'comparison' | 'breakdown' | 'relationship' | 'flow';
}

export interface ChartRecommendation {
  primaryChart: 'column' | 'bar' | 'line' | 'area' | 'pie' | 'scatter' | 'radar';
  confidence: number; // 0-100
  reasoning: string;
  alternatives: Array<{
    chartType: 'column' | 'bar' | 'line' | 'area' | 'pie' | 'scatter' | 'radar';
    reasoning: string;
    confidence: number;
  }>;
  warnings?: string[];
}

export class ChartIntelligence {
  
  /**
   * Analyze data and recommend the best chart type
   */
  static recommendChart(
    dataDescription: string,
    categories: string[],
    seriesNames: string[],
    sampleData?: number[][]
  ): ChartRecommendation {
    
    const characteristics = this.analyzeDataCharacteristics(
      dataDescription,
      categories,
      seriesNames,
      sampleData
    );
    
    return this.selectOptimalChart(characteristics, dataDescription);
  }

  /**
   * Analyze the characteristics of the data
   */
  private static analyzeDataCharacteristics(
    description: string,
    categories: string[],
    seriesNames: string[],
    sampleData?: number[][]
  ): DataCharacteristics {
    
    const desc = description.toLowerCase();
    const categoryText = categories.join(' ').toLowerCase();
    const seriesText = seriesNames.join(' ').toLowerCase();
    
    // Detect time-based data
    const timeKeywords = ['month', 'quarter', 'year', 'week', 'day', 'time', 'period', 'jan', 'feb', 'mar', 'q1', 'q2', 'q3', 'q4'];
    const hasTimeComponent = timeKeywords.some(keyword => 
      categoryText.includes(keyword) || desc.includes(keyword)
    );
    
    // Detect negative values
    let hasNegativeValues = false;
    if (sampleData) {
      hasNegativeValues = sampleData.flat().some(value => value < 0);
    }
    
    // Analyze business context keywords
    const businessContext = this.detectBusinessContext(desc + ' ' + seriesText);
    const dataPattern = this.detectDataPattern(desc, hasTimeComponent, categories.length, seriesNames.length);
    
    // Determine primary data type
    let dataType: DataCharacteristics['dataType'];
    
    if (hasTimeComponent) {
      dataType = 'time-series';
    } else if (this.isCompositionData(desc, seriesText)) {
      dataType = 'composition';
    } else if (this.isCorrelationData(desc, seriesText)) {
      dataType = 'correlation';
    } else if (categories.length > seriesNames.length) {
      dataType = 'categorical';
    } else {
      dataType = 'comparative';
    }
    
    return {
      dataType,
      valueCount: categories.length,
      seriesCount: seriesNames.length,
      hasTimeComponent,
      hasNegativeValues,
      businessContext,
      dataPattern
    };
  }

  /**
   * Select the optimal chart based on characteristics
   */
  private static selectOptimalChart(
    characteristics: DataCharacteristics,
    originalDescription: string
  ): ChartRecommendation {
    
    const { dataType } = characteristics;
    // @ts-ignore - characteristics used in future enhancement
    void characteristics;
    
    // Rule-based chart selection with confidence scoring
    if (dataType === 'time-series') {
      return this.recommendTimeSeriesChart(characteristics, originalDescription);
    }
    
    if (dataType === 'composition') {
      return this.recommendCompositionChart(characteristics);
    }
    
    if (dataType === 'correlation') {
      return this.recommendCorrelationChart(characteristics);
    }
    
    if (dataType === 'comparative') {
      return this.recommendComparativeChart(characteristics);
    }
    
    if (dataType === 'categorical') {
      return this.recommendCategoricalChart(characteristics);
    }
    
    // Default fallback
    return {
      primaryChart: 'column',
      confidence: 50,
      reasoning: 'Default column chart selected for general data comparison',
      alternatives: [
        { chartType: 'bar', reasoning: 'Horizontal comparison option', confidence: 40 },
        { chartType: 'line', reasoning: 'If data has sequential relationship', confidence: 30 }
      ]
    };
  }

  /**
   * Recommend chart for time-series data
   */
  private static recommendTimeSeriesChart(
    characteristics: DataCharacteristics,
    description: string
  ): ChartRecommendation {
    
    const { seriesCount, hasNegativeValues, businessContext } = characteristics;
    const desc = description.toLowerCase();
    
    // Cash flow and financial flows - prefer area charts
    if (desc.includes('cash flow') || desc.includes('cashflow') || 
        desc.includes('cash') || businessContext === 'cash-management') {
      return {
        primaryChart: 'area',
        confidence: 90,
        reasoning: 'Area chart best shows cash flow accumulation over time and handles negative values well',
        alternatives: [
          { chartType: 'line', reasoning: 'Line chart for trend focus without cumulative effect', confidence: 75 },
          { chartType: 'column', reasoning: 'Column chart for period-by-period comparison', confidence: 60 }
        ],
        warnings: hasNegativeValues ? [] : ['Consider showing cumulative values for better cash flow insight']
      };
    }
    
    // Revenue, sales trends - prefer line charts
    if (desc.includes('revenue') || desc.includes('sales') || desc.includes('income') ||
        desc.includes('profit') || businessContext === 'revenue-analysis') {
      return {
        primaryChart: 'line',
        confidence: 85,
        reasoning: 'Line chart effectively shows revenue trends and growth patterns over time',
        alternatives: [
          { chartType: 'area', reasoning: 'Area chart to emphasize cumulative growth', confidence: 70 },
          { chartType: 'column', reasoning: 'Column chart for period comparisons', confidence: 60 }
        ]
      };
    }
    
    // Stock prices, portfolio performance
    if (desc.includes('stock') || desc.includes('portfolio') || desc.includes('price') ||
        businessContext === 'investment-analysis') {
      return {
        primaryChart: 'line',
        confidence: 95,
        reasoning: 'Line chart is the standard for financial price movements and portfolio tracking',
        alternatives: [
          { chartType: 'area', reasoning: 'Area chart to show value accumulation', confidence: 70 }
        ]
      };
    }
    
    // Multiple series time data
    if (seriesCount > 3) {
      return {
        primaryChart: 'line',
        confidence: 80,
        reasoning: 'Line chart handles multiple time series without clutter',
        alternatives: [
          { chartType: 'area', reasoning: 'Stacked area if series are additive', confidence: 60 }
        ],
        warnings: ['Consider limiting to 5 or fewer series for readability']
      };
    }
    
    // Default time-series
    return {
      primaryChart: 'line',
      confidence: 75,
      reasoning: 'Line chart is optimal for showing trends and changes over time',
      alternatives: [
        { chartType: 'area', reasoning: 'Area chart for cumulative visualization', confidence: 65 },
        { chartType: 'column', reasoning: 'Column chart for discrete time periods', confidence: 55 }
      ]
    };
  }

  /**
   * Recommend chart for composition/breakdown data
   */
  private static recommendCompositionChart(characteristics: DataCharacteristics): ChartRecommendation {
    const { valueCount, hasNegativeValues } = characteristics;
    
    if (hasNegativeValues) {
      return {
        primaryChart: 'bar',
        confidence: 85,
        reasoning: 'Bar chart handles negative values better than pie chart for composition data',
        alternatives: [
          { chartType: 'column', reasoning: 'Column chart alternative for negative values', confidence: 75 }
        ],
        warnings: ['Pie charts cannot effectively display negative values']
      };
    }
    
    if (valueCount > 7) {
      return {
        primaryChart: 'bar',
        confidence: 80,
        reasoning: 'Bar chart more readable than pie chart for many categories',
        alternatives: [
          { chartType: 'column', reasoning: 'Column chart for vertical comparison', confidence: 70 }
        ],
        warnings: ['Pie charts become cluttered with more than 7 categories']
      };
    }
    
    return {
      primaryChart: 'pie',
      confidence: 90,
      reasoning: 'Pie chart perfect for showing parts of a whole with few categories',
      alternatives: [
        { chartType: 'bar', reasoning: 'Bar chart for easier value comparison', confidence: 70 },
        { chartType: 'column', reasoning: 'Column chart alternative', confidence: 65 }
      ]
    };
  }

  /**
   * Recommend chart for correlation/relationship data
   */
  private static recommendCorrelationChart(characteristics: DataCharacteristics): ChartRecommendation {
    // @ts-ignore - characteristics used in future enhancement
    void characteristics;
    return {
      primaryChart: 'scatter',
      confidence: 95,
      reasoning: 'Scatter plot is the best choice for showing correlation between two variables',
      alternatives: [
        { chartType: 'line', reasoning: 'Line chart if relationship is clearly sequential', confidence: 40 }
      ]
    };
  }

  /**
   * Recommend chart for comparative data
   */
  private static recommendComparativeChart(characteristics: DataCharacteristics): ChartRecommendation {
    const { valueCount, seriesCount } = characteristics;
    
    if (valueCount > 10) {
      return {
        primaryChart: 'bar',
        confidence: 80,
        reasoning: 'Bar chart more readable for many categories',
        alternatives: [
          { chartType: 'column', reasoning: 'Column chart if labels are short', confidence: 60 }
        ]
      };
    }
    
    if (seriesCount === 1) {
      return {
        primaryChart: 'column',
        confidence: 75,
        reasoning: 'Column chart ideal for comparing single series across categories',
        alternatives: [
          { chartType: 'bar', reasoning: 'Bar chart for long category names', confidence: 70 }
        ]
      };
    }
    
    return {
      primaryChart: 'column',
      confidence: 70,
      reasoning: 'Column chart good for multi-series comparison',
      alternatives: [
        { chartType: 'bar', reasoning: 'Bar chart for horizontal comparison', confidence: 65 },
        { chartType: 'line', reasoning: 'Line chart if categories have natural order', confidence: 50 }
      ]
    };
  }

  /**
   * Recommend chart for categorical data
   */
  private static recommendCategoricalChart(characteristics: DataCharacteristics): ChartRecommendation {
    const { valueCount } = characteristics;
    // @ts-ignore - characteristics used in future enhancement
    void characteristics;
    
    if (valueCount > 8) {
      return {
        primaryChart: 'bar',
        confidence: 85,
        reasoning: 'Bar chart handles many categories better than column chart',
        alternatives: [
          { chartType: 'column', reasoning: 'Column chart if space permits', confidence: 60 }
        ]
      };
    }
    
    return {
      primaryChart: 'column',
      confidence: 75,
      reasoning: 'Column chart effective for categorical comparison',
      alternatives: [
        { chartType: 'bar', reasoning: 'Bar chart for horizontal layout', confidence: 70 }
      ]
    };
  }

  /**
   * Detect business context from description
   */
  private static detectBusinessContext(text: string): string {
    const contexts = {
      'cash-management': ['cash', 'flow', 'liquidity', 'working capital'],
      'revenue-analysis': ['revenue', 'sales', 'income', 'earnings'],
      'expense-analysis': ['expense', 'cost', 'spending', 'budget'],
      'investment-analysis': ['portfolio', 'stock', 'investment', 'return', 'performance'],
      'profitability': ['profit', 'margin', 'profitability', 'net income'],
      'operational': ['production', 'efficiency', 'utilization', 'capacity']
    };
    
    for (const [context, keywords] of Object.entries(contexts)) {
      if (keywords.some(keyword => text.includes(keyword))) {
        return context;
      }
    }
    
    return 'general';
  }

  /**
   * Detect data pattern
   */
  private static detectDataPattern(
    description: string,
    hasTimeComponent: boolean,
    categoryCount: number,
    seriesCount: number
  ): 'trend' | 'comparison' | 'breakdown' | 'relationship' | 'flow' {
    
    const desc = description.toLowerCase();
    // @ts-ignore - parameters used in future enhancement
    void categoryCount;
    // @ts-ignore - parameters used in future enhancement  
    void seriesCount;
    
    if (desc.includes('trend') || desc.includes('growth') || hasTimeComponent) {
      return 'trend';
    }
    
    if (desc.includes('breakdown') || desc.includes('composition') || desc.includes('split')) {
      return 'breakdown';
    }
    
    if (desc.includes('correlation') || desc.includes('relationship') || desc.includes('vs')) {
      return 'relationship';
    }
    
    if (desc.includes('flow') || desc.includes('movement')) {
      return 'flow';
    }
    
    return 'comparison';
  }

  /**
   * Check if data represents composition/parts of whole
   */
  private static isCompositionData(description: string, seriesText: string): boolean {
    const compositionKeywords = ['breakdown', 'split', 'composition', 'allocation', 'distribution', 'share', 'percentage', 'portion'];
    const text = (description + ' ' + seriesText).toLowerCase();
    return compositionKeywords.some(keyword => text.includes(keyword));
  }

  /**
   * Check if data represents correlation/relationship
   */
  private static isCorrelationData(description: string, seriesText: string): boolean {
    const correlationKeywords = ['correlation', 'relationship', 'vs', 'against', 'scatter', 'plot'];
    const text = (description + ' ' + seriesText).toLowerCase();
    return correlationKeywords.some(keyword => text.includes(keyword));
  }

  /**
   * Get chart recommendation explanation for users
   */
  static explainRecommendation(recommendation: ChartRecommendation): string {
    let explanation = `Recommended: ${recommendation.primaryChart.toUpperCase()} chart (${recommendation.confidence}% confidence)\n`;
    explanation += `Reason: ${recommendation.reasoning}\n`;
    
    if (recommendation.alternatives.length > 0) {
      explanation += '\nAlternative options:\n';
      recommendation.alternatives.forEach(alt => {
        explanation += `- ${alt.chartType.toUpperCase()}: ${alt.reasoning} (${alt.confidence}% confidence)\n`;
      });
    }
    
    if (recommendation.warnings && recommendation.warnings.length > 0) {
      explanation += '\nWarnings:\n';
      recommendation.warnings.forEach(warning => {
        explanation += `⚠️  ${warning}\n`;
      });
    }
    
    return explanation;
  }
}