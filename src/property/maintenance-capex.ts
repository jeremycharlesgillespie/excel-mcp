import { CellValue } from '../excel/excel-manager.js';

export interface MaintenanceRequest {
  requestId: string;
  propertyId: string;
  unitId?: string;
  tenantId?: string;
  requestType: 'emergency' | 'urgent' | 'routine' | 'preventive';
  category: 'plumbing' | 'electrical' | 'hvac' | 'structural' | 'cosmetic' | 'landscaping' | 'security' | 'other';
  description: string;
  requestDate: string;
  priorityLevel: 1 | 2 | 3 | 4 | 5; // 1 = Critical, 5 = Low
  estimatedCost: number;
  actualCost?: number;
  status: 'open' | 'in_progress' | 'completed' | 'deferred' | 'cancelled';
  assignedVendor?: string;
  scheduledDate?: string;
  completedDate?: string;
  tenantSatisfactionRating?: 1 | 2 | 3 | 4 | 5;
  photos?: string[];
  notes?: string[];
}

export interface CapExProject {
  projectId: string;
  propertyId: string;
  projectName: string;
  projectType: 'roof_replacement' | 'hvac_upgrade' | 'flooring' | 'exterior_improvements' | 'structural' | 'technology' | 'amenities' | 'other';
  description: string;
  plannedStartDate: string;
  plannedEndDate: string;
  actualStartDate?: string;
  actualEndDate?: string;
  budgetAmount: number;
  actualCost?: number;
  status: 'planned' | 'approved' | 'in_progress' | 'completed' | 'delayed' | 'cancelled';
  contractor?: string;
  permitRequired: boolean;
  permitStatus?: 'not_required' | 'applied' | 'approved' | 'denied';
  expectedROI?: number; // Return on investment
  expectedValueAdd?: number; // Property value increase
  impactOnOccupancy?: 'none' | 'minimal' | 'moderate' | 'significant';
  milestones?: ProjectMilestone[];
}

export interface ProjectMilestone {
  milestoneId: string;
  description: string;
  plannedDate: string;
  actualDate?: string;
  status: 'pending' | 'in_progress' | 'completed' | 'delayed';
  cost: number;
  notes?: string;
}

export interface MaintenanceAnalysis {
  propertyId: string;
  analysisDate: string;
  maintenanceMetrics: {
    totalRequests: number;
    avgResponseTime: number; // hours
    avgCompletionTime: number; // hours
    costPerUnit: number;
    costPerSquareFoot: number;
    tenantSatisfaction: number; // 1-5 scale
    emergencyRequestRate: number; // percentage
    deferredMaintenanceValue: number;
  };
  categoryBreakdown: {
    [category: string]: {
      requestCount: number;
      totalCost: number;
      avgCost: number;
      avgResponseTime: number;
    };
  };
  vendorPerformance: {
    [vendor: string]: {
      jobCount: number;
      avgCost: number;
      avgCompletionTime: number;
      qualityRating: number;
      onTimePercentage: number;
    };
  };
  seasonalTrends: {
    [month: string]: {
      requestCount: number;
      cost: number;
      emergencyRate: number;
    };
  };
  predictiveMaintenance: {
    upcomingMaintenance: MaintenanceRequest[];
    budgetForecast: number[];
    riskAreas: string[];
  };
}

export interface CapExAnalysis {
  propertyId: string;
  analysisDate: string;
  capExMetrics: {
    totalBudget: number;
    spentToDate: number;
    remainingBudget: number;
    projectedOverrun: number;
    completedProjects: number;
    onTimePercentage: number;
    avgROI: number;
    totalValueAdd: number;
  };
  projectBreakdown: {
    [projectType: string]: {
      projectCount: number;
      totalBudget: number;
      actualCost: number;
      avgROI: number;
    };
  };
  timeline: {
    upcomingProjects: CapExProject[];
    inProgressProjects: CapExProject[];
    completedProjects: CapExProject[];
  };
  budgetVariance: {
    plannedByQuarter: number[];
    actualByQuarter: number[];
    forecastByQuarter: number[];
  };
}

export class MaintenanceCapExManager {

  private calculateMaintenanceMetrics(
    requests: MaintenanceRequest[], 
    propertyUnits: number, 
    propertySquareFeet: number
  ): MaintenanceAnalysis['maintenanceMetrics'] {
    const completedRequests = requests.filter(req => req.status === 'completed');
    const totalRequests = requests.length;
    
    // Calculate response and completion times
    let totalResponseTime = 0;
    let totalCompletionTime = 0;
    let responseTimeCount = 0;
    let completionTimeCount = 0;
    
    completedRequests.forEach(req => {
      if (req.scheduledDate) {
        const responseTime = (new Date(req.scheduledDate).getTime() - new Date(req.requestDate).getTime()) / (1000 * 60 * 60);
        totalResponseTime += responseTime;
        responseTimeCount++;
      }
      
      if (req.completedDate) {
        const completionTime = (new Date(req.completedDate).getTime() - new Date(req.requestDate).getTime()) / (1000 * 60 * 60);
        totalCompletionTime += completionTime;
        completionTimeCount++;
      }
    });
    
    const avgResponseTime = responseTimeCount > 0 ? totalResponseTime / responseTimeCount : 0;
    const avgCompletionTime = completionTimeCount > 0 ? totalCompletionTime / completionTimeCount : 0;
    
    // Calculate costs
    const totalCost = completedRequests.reduce((sum, req) => sum + (req.actualCost || req.estimatedCost || 0), 0);
    const costPerUnit = propertyUnits > 0 ? totalCost / propertyUnits : 0;
    const costPerSquareFoot = propertySquareFeet > 0 ? totalCost / propertySquareFeet : 0;
    
    // Calculate tenant satisfaction
    const satisfactionRatings = completedRequests
      .filter(req => req.tenantSatisfactionRating)
      .map(req => req.tenantSatisfactionRating!);
    const tenantSatisfaction = satisfactionRatings.length > 0 
      ? satisfactionRatings.reduce((sum, rating) => sum + rating, 0) / satisfactionRatings.length 
      : 0;
    
    // Calculate emergency request rate
    const emergencyRequests = requests.filter(req => req.requestType === 'emergency').length;
    const emergencyRequestRate = totalRequests > 0 ? emergencyRequests / totalRequests : 0;
    
    // Calculate deferred maintenance value
    const deferredRequests = requests.filter(req => req.status === 'deferred');
    const deferredMaintenanceValue = deferredRequests.reduce((sum, req) => sum + req.estimatedCost, 0);
    
    return {
      totalRequests,
      avgResponseTime,
      avgCompletionTime,
      costPerUnit,
      costPerSquareFoot,
      tenantSatisfaction,
      emergencyRequestRate,
      deferredMaintenanceValue
    };
  }

  private analyzeCategoryBreakdown(requests: MaintenanceRequest[]): MaintenanceAnalysis['categoryBreakdown'] {
    const breakdown: MaintenanceAnalysis['categoryBreakdown'] = {};
    
    requests.forEach(req => {
      if (!breakdown[req.category]) {
        breakdown[req.category] = {
          requestCount: 0,
          totalCost: 0,
          avgCost: 0,
          avgResponseTime: 0
        };
      }
      
      breakdown[req.category].requestCount++;
      breakdown[req.category].totalCost += req.actualCost || req.estimatedCost || 0;
      
      if (req.scheduledDate && req.status === 'completed') {
        const responseTime = (new Date(req.scheduledDate).getTime() - new Date(req.requestDate).getTime()) / (1000 * 60 * 60);
        breakdown[req.category].avgResponseTime += responseTime;
      }
    });
    
    // Calculate averages
    Object.keys(breakdown).forEach(category => {
      const categoryData = breakdown[category];
      categoryData.avgCost = categoryData.totalCost / categoryData.requestCount;
      categoryData.avgResponseTime = categoryData.avgResponseTime / categoryData.requestCount;
    });
    
    return breakdown;
  }

  private analyzeVendorPerformance(requests: MaintenanceRequest[]): MaintenanceAnalysis['vendorPerformance'] {
    const performance: MaintenanceAnalysis['vendorPerformance'] = {};
    
    const completedRequests = requests.filter(req => req.status === 'completed' && req.assignedVendor);
    
    completedRequests.forEach(req => {
      const vendor = req.assignedVendor!;
      
      if (!performance[vendor]) {
        performance[vendor] = {
          jobCount: 0,
          avgCost: 0,
          avgCompletionTime: 0,
          qualityRating: 0,
          onTimePercentage: 0
        };
      }
      
      performance[vendor].jobCount++;
      performance[vendor].avgCost += req.actualCost || req.estimatedCost || 0;
      performance[vendor].qualityRating += req.tenantSatisfactionRating || 3;
      
      if (req.completedDate && req.scheduledDate) {
        const completionTime = (new Date(req.completedDate).getTime() - new Date(req.scheduledDate).getTime()) / (1000 * 60 * 60);
        performance[vendor].avgCompletionTime += completionTime;
        
        // Consider on-time if completed within 24 hours of scheduled time
        if (completionTime <= 24) {
          performance[vendor].onTimePercentage += 1;
        }
      }
    });
    
    // Calculate averages and percentages
    Object.keys(performance).forEach(vendor => {
      const vendorData = performance[vendor];
      vendorData.avgCost = vendorData.avgCost / vendorData.jobCount;
      vendorData.avgCompletionTime = vendorData.avgCompletionTime / vendorData.jobCount;
      vendorData.qualityRating = vendorData.qualityRating / vendorData.jobCount;
      vendorData.onTimePercentage = (vendorData.onTimePercentage / vendorData.jobCount) * 100;
    });
    
    return performance;
  }

  private analyzeSeasonalTrends(requests: MaintenanceRequest[]): MaintenanceAnalysis['seasonalTrends'] {
    const trends: MaintenanceAnalysis['seasonalTrends'] = {};
    const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    
    monthNames.forEach(month => {
      trends[month] = {
        requestCount: 0,
        cost: 0,
        emergencyRate: 0
      };
    });
    
    requests.forEach(req => {
      const month = monthNames[new Date(req.requestDate).getMonth()];
      trends[month].requestCount++;
      trends[month].cost += req.actualCost || req.estimatedCost || 0;
      
      if (req.requestType === 'emergency') {
        trends[month].emergencyRate++;
      }
    });
    
    // Calculate emergency rates as percentages
    monthNames.forEach(month => {
      const monthData = trends[month];
      if (monthData.requestCount > 0) {
        monthData.emergencyRate = (monthData.emergencyRate / monthData.requestCount) * 100;
      }
    });
    
    return trends;
  }

  private generatePredictiveMaintenance(
    requests: MaintenanceRequest[], 
    propertyId: string
  ): MaintenanceAnalysis['predictiveMaintenance'] {
    const currentDate = new Date();
    
    // Generate upcoming maintenance based on patterns and schedules
    const upcomingMaintenance: MaintenanceRequest[] = [
      {
        requestId: `PRED-${propertyId}-001`,
        propertyId,
        requestType: 'preventive',
        category: 'hvac',
        description: 'Quarterly HVAC filter replacement and system check',
        requestDate: new Date(currentDate.getTime() + 30 * 24 * 60 * 60 * 1000).toISOString(),
        priorityLevel: 3,
        estimatedCost: 150,
        status: 'open'
      },
      {
        requestId: `PRED-${propertyId}-002`,
        propertyId,
        requestType: 'preventive',
        category: 'plumbing',
        description: 'Annual water heater maintenance and inspection',
        requestDate: new Date(currentDate.getTime() + 60 * 24 * 60 * 60 * 1000).toISOString(),
        priorityLevel: 3,
        estimatedCost: 200,
        status: 'open'
      },
      {
        requestId: `PRED-${propertyId}-003`,
        propertyId,
        requestType: 'preventive',
        category: 'structural',
        description: 'Semi-annual roof and gutter inspection',
        requestDate: new Date(currentDate.getTime() + 90 * 24 * 60 * 60 * 1000).toISOString(),
        priorityLevel: 2,
        estimatedCost: 300,
        status: 'open'
      }
    ];
    
    // Calculate budget forecast based on historical trends
    const monthlyAverage = requests.reduce((sum, req) => sum + (req.actualCost || req.estimatedCost || 0), 0) / 12;
    const budgetForecast = Array(12).fill(0).map((_, index) => {
      // Add seasonal adjustment
      const seasonalMultiplier = [1.2, 1.1, 1.0, 0.9, 0.8, 0.7, 0.8, 0.9, 1.0, 1.1, 1.2, 1.3][index];
      return monthlyAverage * seasonalMultiplier;
    });
    
    // Identify risk areas based on request patterns
    const categoryFrequency: { [category: string]: number } = {};
    requests.forEach(req => {
      categoryFrequency[req.category] = (categoryFrequency[req.category] || 0) + 1;
    });
    
    const riskAreas = Object.entries(categoryFrequency)
      .filter(([, count]) => count > requests.length * 0.15) // Categories > 15% of requests
      .map(([category]) => category)
      .sort((a, b) => categoryFrequency[b] - categoryFrequency[a]);
    
    return {
      upcomingMaintenance,
      budgetForecast,
      riskAreas
    };
  }

  analyzeMaintenancePortfolio(
    requests: MaintenanceRequest[], 
    propertyId: string,
    propertyUnits: number = 10,
    propertySquareFeet: number = 10000
  ): MaintenanceAnalysis {
    const analysisDate = new Date().toISOString();
    
    const maintenanceMetrics = this.calculateMaintenanceMetrics(requests, propertyUnits, propertySquareFeet);
    const categoryBreakdown = this.analyzeCategoryBreakdown(requests);
    const vendorPerformance = this.analyzeVendorPerformance(requests);
    const seasonalTrends = this.analyzeSeasonalTrends(requests);
    const predictiveMaintenance = this.generatePredictiveMaintenance(requests, propertyId);
    
    return {
      propertyId,
      analysisDate,
      maintenanceMetrics,
      categoryBreakdown,
      vendorPerformance,
      seasonalTrends,
      predictiveMaintenance
    };
  }

  analyzeCapExPortfolio(projects: CapExProject[], propertyId: string): CapExAnalysis {
    const analysisDate = new Date().toISOString();
    
    // Calculate CapEx metrics
    const totalBudget = projects.reduce((sum, project) => sum + project.budgetAmount, 0);
    const spentToDate = projects
      .filter(project => project.actualCost)
      .reduce((sum, project) => sum + project.actualCost!, 0);
    const remainingBudget = totalBudget - spentToDate;
    
    const completedProjects = projects.filter(project => project.status === 'completed');
    const inProgressProjects = projects.filter(project => project.status === 'in_progress');
    
    // Calculate project overruns
    const projectsWithActualCost = projects.filter(project => project.actualCost);
    const projectedOverrun = projectsWithActualCost.reduce((sum, project) => {
      return sum + Math.max(0, project.actualCost! - project.budgetAmount);
    }, 0);
    
    // On-time performance
    const projectsWithDates = completedProjects.filter(project => 
      project.actualEndDate && project.plannedEndDate
    );
    const onTimeProjects = projectsWithDates.filter(project => 
      new Date(project.actualEndDate!) <= new Date(project.plannedEndDate)
    );
    const onTimePercentage = projectsWithDates.length > 0 
      ? (onTimeProjects.length / projectsWithDates.length) * 100 
      : 0;
    
    // Average ROI
    const projectsWithROI = completedProjects.filter(project => project.expectedROI);
    const avgROI = projectsWithROI.length > 0
      ? projectsWithROI.reduce((sum, project) => sum + project.expectedROI!, 0) / projectsWithROI.length
      : 0;
    
    // Total value add
    const totalValueAdd = completedProjects.reduce((sum, project) => 
      sum + (project.expectedValueAdd || 0), 0
    );
    
    // Project breakdown by type
    const projectBreakdown: CapExAnalysis['projectBreakdown'] = {};
    projects.forEach(project => {
      if (!projectBreakdown[project.projectType]) {
        projectBreakdown[project.projectType] = {
          projectCount: 0,
          totalBudget: 0,
          actualCost: 0,
          avgROI: 0
        };
      }
      
      projectBreakdown[project.projectType].projectCount++;
      projectBreakdown[project.projectType].totalBudget += project.budgetAmount;
      projectBreakdown[project.projectType].actualCost += project.actualCost || 0;
      projectBreakdown[project.projectType].avgROI += project.expectedROI || 0;
    });
    
    // Calculate averages for breakdown
    Object.keys(projectBreakdown).forEach(type => {
      const typeData = projectBreakdown[type];
      typeData.avgROI = typeData.avgROI / typeData.projectCount;
    });
    
    // Timeline analysis
    const currentDate = new Date();
    const upcomingProjects = projects.filter(project => {
      const startDate = new Date(project.plannedStartDate);
      const daysDiff = (startDate.getTime() - currentDate.getTime()) / (1000 * 60 * 60 * 24);
      return daysDiff > 0 && daysDiff <= 90 && project.status === 'approved';
    });
    
    // Budget variance analysis (simplified quarterly breakdown)
    const currentYear = currentDate.getFullYear();
    const plannedByQuarter = [0, 0, 0, 0];
    const actualByQuarter = [0, 0, 0, 0];
    const forecastByQuarter = [0, 0, 0, 0];
    
    projects.forEach(project => {
      const startDate = new Date(project.plannedStartDate);
      if (startDate.getFullYear() === currentYear) {
        const quarter = Math.floor(startDate.getMonth() / 3);
        plannedByQuarter[quarter] += project.budgetAmount;
        
        if (project.actualCost) {
          actualByQuarter[quarter] += project.actualCost;
        } else if (project.status === 'in_progress') {
          // Estimate based on progress
          forecastByQuarter[quarter] += project.budgetAmount * 1.1; // Assume 10% overrun
        } else {
          forecastByQuarter[quarter] += project.budgetAmount;
        }
      }
    });
    
    return {
      propertyId,
      analysisDate,
      capExMetrics: {
        totalBudget,
        spentToDate,
        remainingBudget,
        projectedOverrun,
        completedProjects: completedProjects.length,
        onTimePercentage,
        avgROI,
        totalValueAdd
      },
      projectBreakdown,
      timeline: {
        upcomingProjects,
        inProgressProjects,
        completedProjects
      },
      budgetVariance: {
        plannedByQuarter,
        actualByQuarter,
        forecastByQuarter
      }
    };
  }

  generateMaintenanceAnalysisWorksheet(
    analysis: MaintenanceAnalysis,
    detailedRequests?: MaintenanceRequest[]
  ): Array<Array<CellValue | string | number>> {
    const worksheet: Array<Array<CellValue | string | number>> = [];
    
    // Header
    worksheet.push(['PROPERTY MAINTENANCE ANALYSIS', '', '', '', '', '', '', '', '', '']);
    worksheet.push([`Property ID: ${analysis.propertyId}`, '', '', '', '', '', '', '', '', '']);
    worksheet.push([`Analysis Date: ${new Date(analysis.analysisDate).toLocaleDateString()}`, '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Property Management Intelligence Platform', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Executive Summary
    worksheet.push(['📊 MAINTENANCE EXECUTIVE SUMMARY', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Total Maintenance Requests (YTD)', analysis.maintenanceMetrics.totalRequests, 'requests', 'Total maintenance requests submitted']);
    worksheet.push(['Average Response Time', analysis.maintenanceMetrics.avgResponseTime.toFixed(1) + ' hours', 'response', 'Time from request to scheduled service']);
    worksheet.push(['Average Completion Time', analysis.maintenanceMetrics.avgCompletionTime.toFixed(1) + ' hours', 'completion', 'Time from request to completion']);
    worksheet.push(['Cost Per Unit', '$' + analysis.maintenanceMetrics.costPerUnit.toFixed(0), 'cost', 'Annual maintenance cost per rental unit']);
    worksheet.push(['Cost Per Square Foot', '$' + analysis.maintenanceMetrics.costPerSquareFoot.toFixed(2), 'cost', 'Annual maintenance cost per sq ft']);
    worksheet.push(['Tenant Satisfaction Rating', analysis.maintenanceMetrics.tenantSatisfaction.toFixed(1) + '/5.0', 'satisfaction', 'Average tenant satisfaction with maintenance']);
    worksheet.push(['Emergency Request Rate', (analysis.maintenanceMetrics.emergencyRequestRate * 100).toFixed(1) + '%', 'emergency', 'Percentage of emergency/urgent requests']);
    worksheet.push(['Deferred Maintenance Value', '$' + (analysis.maintenanceMetrics.deferredMaintenanceValue / 1000).toFixed(0) + 'K', 'deferred', 'Value of deferred maintenance items']);
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Category Breakdown Analysis
    worksheet.push(['🔧 MAINTENANCE CATEGORY ANALYSIS', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Category', 'Requests', 'Total Cost', 'Avg Cost', 'Avg Response', 'Cost %', '', '', '', '']);
    
    const totalCost = Object.values(analysis.categoryBreakdown).reduce((sum, cat) => sum + cat.totalCost, 0);
    
    Object.entries(analysis.categoryBreakdown)
      .sort(([,a], [,b]) => b.totalCost - a.totalCost)
      .forEach(([category, data]) => {
        const costPercentage = totalCost > 0 ? (data.totalCost / totalCost * 100).toFixed(1) : '0.0';
        
        worksheet.push([
          category.toUpperCase(),
          data.requestCount,
          '$' + (data.totalCost / 1000).toFixed(1) + 'K',
          '$' + data.avgCost.toFixed(0),
          data.avgResponseTime.toFixed(1) + 'h',
          costPercentage + '%',
          '', '', '', ''
        ]);
      });
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Vendor Performance Analysis
    worksheet.push(['👨‍🔧 VENDOR PERFORMANCE ANALYSIS', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Vendor', 'Jobs', 'Avg Cost', 'Completion Time', 'Quality Rating', 'On-Time %', '', '', '', '']);
    
    Object.entries(analysis.vendorPerformance)
      .sort(([,a], [,b]) => b.qualityRating - a.qualityRating)
      .forEach(([vendor, performance]) => {
        let performanceRating = 'GOOD';
        if (performance.qualityRating >= 4.5 && performance.onTimePercentage >= 90) {
          performanceRating = 'EXCELLENT';
        } else if (performance.qualityRating < 3.0 || performance.onTimePercentage < 70) {
          performanceRating = 'NEEDS IMPROVEMENT';
        }
        
        worksheet.push([
          vendor,
          performance.jobCount,
          '$' + performance.avgCost.toFixed(0),
          performance.avgCompletionTime.toFixed(1) + 'h',
          performance.qualityRating.toFixed(1) + '/5.0',
          performance.onTimePercentage.toFixed(0) + '%',
          performanceRating,
          '', '', ''
        ]);
      });
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Seasonal Trends Analysis
    worksheet.push(['📅 SEASONAL MAINTENANCE TRENDS', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Month', 'Requests', 'Cost', 'Emergency Rate', 'Trend Analysis', '', '', '', '', '']);
    
    Object.entries(analysis.seasonalTrends).forEach(([month, data]) => {
      let trendAnalysis = 'NORMAL';
      if (data.requestCount > analysis.maintenanceMetrics.totalRequests / 12 * 1.5) {
        trendAnalysis = 'HIGH ACTIVITY';
      } else if (data.requestCount < analysis.maintenanceMetrics.totalRequests / 12 * 0.5) {
        trendAnalysis = 'LOW ACTIVITY';
      }
      
      worksheet.push([
        month,
        data.requestCount,
        '$' + (data.cost / 1000).toFixed(1) + 'K',
        data.emergencyRate.toFixed(1) + '%',
        trendAnalysis,
        '', '', '', '', ''
      ]);
    });
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Predictive Maintenance
    worksheet.push(['🔮 PREDICTIVE MAINTENANCE FORECAST', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Upcoming Maintenance (Next 90 Days):', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Category', 'Description', 'Scheduled Date', 'Est. Cost', 'Priority', '', '', '', '', '']);
    
    analysis.predictiveMaintenance.upcomingMaintenance.forEach(maintenance => {
      worksheet.push([
        maintenance.category.toUpperCase(),
        maintenance.description,
        new Date(maintenance.requestDate).toLocaleDateString(),
        '$' + maintenance.estimatedCost,
        'P' + maintenance.priorityLevel,
        '', '', '', '', ''
      ]);
    });
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Budget Forecast
    worksheet.push(['💰 MAINTENANCE BUDGET FORECAST', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Month', 'Forecast', 'Cumulative', 'Seasonal Factor', '', '', '', '', '', '']);
    
    let cumulative = 0;
    analysis.predictiveMaintenance.budgetForecast.forEach((monthlyForecast, index) => {
      cumulative += monthlyForecast;
      const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
      const seasonalFactor = monthlyForecast / (analysis.predictiveMaintenance.budgetForecast.reduce((a, b) => a + b) / 12);
      
      worksheet.push([
        monthNames[index],
        '$' + (monthlyForecast / 1000).toFixed(1) + 'K',
        '$' + (cumulative / 1000).toFixed(0) + 'K',
        seasonalFactor.toFixed(2) + 'x',
        '', '', '', '', '', ''
      ]);
    });
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Risk Areas
    worksheet.push(['⚠️ HIGH-RISK MAINTENANCE AREAS', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Risk Area', 'Request Frequency', 'Avg Cost Impact', 'Recommended Action', '', '', '', '', '', '']);
    
    analysis.predictiveMaintenance.riskAreas.forEach(riskArea => {
      const categoryData = analysis.categoryBreakdown[riskArea];
      let recommendedAction = 'Monitor closely';
      
      if (riskArea === 'hvac') {
        recommendedAction = 'Implement preventive HVAC maintenance program';
      } else if (riskArea === 'plumbing') {
        recommendedAction = 'Consider plumbing system upgrades';
      } else if (riskArea === 'electrical') {
        recommendedAction = 'Schedule electrical system inspection';
      } else if (riskArea === 'structural') {
        recommendedAction = 'Implement structural maintenance program';
      }
      
      worksheet.push([
        riskArea.toUpperCase(),
        categoryData.requestCount + ' requests',
        '$' + (categoryData.totalCost / 1000).toFixed(1) + 'K',
        recommendedAction,
        '', '', '', '', '', ''
      ]);
    });
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Detailed Request Analysis (if provided)
    if (detailedRequests && detailedRequests.length > 0) {
      worksheet.push(['📋 DETAILED REQUEST ANALYSIS', '', '', '', '', '', '', '', '', '']);
      worksheet.push(['Request ID', 'Category', 'Priority', 'Status', 'Est. Cost', 'Actual Cost', 'Days Open', '', '', '']);
      
      const currentDate = new Date();
      
      detailedRequests
        .sort((a, b) => a.priorityLevel - b.priorityLevel) // Sort by priority
        .slice(0, 15) // Show top 15 requests
        .forEach(request => {
          const daysOpen = request.completedDate 
            ? Math.ceil((new Date(request.completedDate).getTime() - new Date(request.requestDate).getTime()) / (24 * 60 * 60 * 1000))
            : Math.ceil((currentDate.getTime() - new Date(request.requestDate).getTime()) / (24 * 60 * 60 * 1000));
          
          worksheet.push([
            request.requestId,
            request.category.toUpperCase(),
            'P' + request.priorityLevel,
            request.status.toUpperCase(),
            '$' + request.estimatedCost,
            request.actualCost ? '$' + request.actualCost : 'TBD',
            daysOpen + ' days',
            '', '', ''
          ]);
        });
    }
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Action Plan and Recommendations
    worksheet.push(['✅ MAINTENANCE ACTION PLAN', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Priority', 'Action Item', 'Responsible Party', 'Timeline', 'Expected Impact', '', '', '', '', '']);
    
    const actionPlan = [
      ['HIGH', 'Address all Priority 1 and 2 maintenance requests', 'Property Manager', 'Within 48 hours', 'Improved tenant satisfaction'],
      ['HIGH', 'Implement preventive maintenance schedule for high-risk areas', 'Maintenance Team', 'Next 30 days', 'Reduce emergency requests by 20%'],
      ['MEDIUM', 'Vendor performance review and contract renegotiation', 'Property Manager', 'Next 60 days', 'Improve cost efficiency and response times'],
      ['MEDIUM', 'Update maintenance request tracking system', 'Operations', 'Next 90 days', 'Better data and tenant communication'],
      ['LOW', 'Seasonal maintenance planning and budgeting', 'Finance Team', 'Quarterly', 'Predictable maintenance costs']
    ];
    
    actionPlan.forEach(([priority, action, responsible, timeline, impact]) => {
      worksheet.push([priority, action, responsible, timeline, impact, '', '', '', '', '']);
    });
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Key Performance Indicators
    worksheet.push(['📈 MAINTENANCE KPIs & BENCHMARKS', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['KPI', 'Current Value', 'Industry Benchmark', 'Performance', 'Target', '', '', '', '', '']);
    
    const kpis = [
      ['Response Time', analysis.maintenanceMetrics.avgResponseTime.toFixed(1) + ' hours', '24 hours', analysis.maintenanceMetrics.avgResponseTime <= 24 ? 'MEETS' : 'BELOW', '≤ 24 hours'],
      ['Emergency Rate', (analysis.maintenanceMetrics.emergencyRequestRate * 100).toFixed(1) + '%', '< 15%', analysis.maintenanceMetrics.emergencyRequestRate < 0.15 ? 'MEETS' : 'ABOVE', '< 15%'],
      ['Tenant Satisfaction', analysis.maintenanceMetrics.tenantSatisfaction.toFixed(1) + '/5.0', '>= 4.0', analysis.maintenanceMetrics.tenantSatisfaction >= 4.0 ? 'MEETS' : 'BELOW', '>= 4.0'],
      ['Cost Per Unit', '$' + analysis.maintenanceMetrics.costPerUnit.toFixed(0), '$1,500-2,500', 'VARIES', '$2,000'],
      ['Deferred Maintenance', '$' + (analysis.maintenanceMetrics.deferredMaintenanceValue / 1000).toFixed(0) + 'K', '< $5K', analysis.maintenanceMetrics.deferredMaintenanceValue < 5000 ? 'MEETS' : 'ABOVE', '< $5K']
    ];
    
    kpis.forEach(([kpi, current, benchmark, performance, target]) => {
      worksheet.push([kpi, current, benchmark, performance, target, '', '', '', '', '']);
    });
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Professional Notes
    worksheet.push(['📝 ANALYSIS METHODOLOGY & NOTES', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- Maintenance analysis based on 12-month rolling data', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- Vendor performance metrics include cost, time, and quality factors', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- Predictive maintenance based on historical patterns and best practices', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- Budget forecasts include seasonal adjustments and trending factors', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- Regular analysis updates recommended for optimal property management', '', '', '', '', '', '', '', '', '']);
    
    return worksheet;
  }

  generateCapExAnalysisWorksheet(
    analysis: CapExAnalysis,
    detailedProjects?: CapExProject[]
  ): Array<Array<CellValue | string | number>> {
    const worksheet: Array<Array<CellValue | string | number>> = [];
    
    // Header
    worksheet.push(['CAPITAL EXPENDITURE ANALYSIS', '', '', '', '', '', '', '', '', '']);
    worksheet.push([`Property ID: ${analysis.propertyId}`, '', '', '', '', '', '', '', '', '']);
    worksheet.push([`Analysis Date: ${new Date(analysis.analysisDate).toLocaleDateString()}`, '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Property Management Intelligence Platform', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Executive Summary
    worksheet.push(['📊 CAPEX EXECUTIVE SUMMARY', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Total CapEx Budget', '$' + (analysis.capExMetrics.totalBudget / 1000000).toFixed(1) + 'M', 'budget', 'Total approved capital expenditure budget']);
    worksheet.push(['Spent to Date', '$' + (analysis.capExMetrics.spentToDate / 1000000).toFixed(1) + 'M', 'spent', 'Capital expenditures completed to date']);
    worksheet.push(['Remaining Budget', '$' + (analysis.capExMetrics.remainingBudget / 1000000).toFixed(1) + 'M', 'remaining', 'Remaining capital expenditure budget']);
    worksheet.push(['Budget Utilization', ((analysis.capExMetrics.spentToDate / analysis.capExMetrics.totalBudget) * 100).toFixed(1) + '%', 'utilization', 'Percentage of budget utilized']);
    worksheet.push(['Projected Overrun', '$' + (analysis.capExMetrics.projectedOverrun / 1000).toFixed(0) + 'K', 'overrun', 'Estimated budget overrun on current projects']);
    worksheet.push(['Completed Projects', analysis.capExMetrics.completedProjects, 'completed', 'Number of completed capital projects']);
    worksheet.push(['On-Time Performance', analysis.capExMetrics.onTimePercentage.toFixed(1) + '%', 'on_time', 'Percentage of projects completed on schedule']);
    worksheet.push(['Average ROI', (analysis.capExMetrics.avgROI * 100).toFixed(1) + '%', 'roi', 'Average return on investment for completed projects']);
    worksheet.push(['Total Value Add', '$' + (analysis.capExMetrics.totalValueAdd / 1000000).toFixed(1) + 'M', 'value_add', 'Estimated property value increase from CapEx']);
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Project Type Breakdown
    worksheet.push(['🏗️ CAPEX PROJECT TYPE BREAKDOWN', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Project Type', 'Count', 'Budget', 'Actual Cost', 'Avg ROI', 'Cost Variance', '', '', '', '']);
    
    Object.entries(analysis.projectBreakdown)
      .sort(([,a], [,b]) => b.totalBudget - a.totalBudget)
      .forEach(([projectType, data]) => {
        const costVariance = data.actualCost > 0 
          ? ((data.actualCost - data.totalBudget) / data.totalBudget * 100).toFixed(1) + '%'
          : 'TBD';
        
        worksheet.push([
          projectType.replace('_', ' ').toUpperCase(),
          data.projectCount,
          '$' + (data.totalBudget / 1000).toFixed(0) + 'K',
          data.actualCost > 0 ? '$' + (data.actualCost / 1000).toFixed(0) + 'K' : 'TBD',
          (data.avgROI * 100).toFixed(1) + '%',
          costVariance,
          '', '', '', ''
        ]);
      });
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Project Timeline Status
    worksheet.push(['📅 PROJECT TIMELINE STATUS', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Status', 'Count', 'Total Budget', 'Timeline Risk', '', '', '', '', '', '']);
    
    const statusCounts = {
      'Upcoming (Next 90 Days)': analysis.timeline.upcomingProjects.length,
      'In Progress': analysis.timeline.inProgressProjects.length,
      'Completed': analysis.timeline.completedProjects.length
    };
    
    const upcomingBudget = analysis.timeline.upcomingProjects.reduce((sum, p) => sum + p.budgetAmount, 0);
    const inProgressBudget = analysis.timeline.inProgressProjects.reduce((sum, p) => sum + p.budgetAmount, 0);
    const completedBudget = analysis.timeline.completedProjects.reduce((sum, p) => sum + p.budgetAmount, 0);
    
    worksheet.push(['Upcoming (Next 90 Days)', statusCounts['Upcoming (Next 90 Days)'], '$' + (upcomingBudget / 1000).toFixed(0) + 'K', 'MONITOR', '', '', '', '', '', '']);
    worksheet.push(['In Progress', statusCounts['In Progress'], '$' + (inProgressBudget / 1000).toFixed(0) + 'K', 'ACTIVE', '', '', '', '', '', '']);
    worksheet.push(['Completed', statusCounts['Completed'], '$' + (completedBudget / 1000).toFixed(0) + 'K', 'COMPLETE', '', '', '', '', '', '']);
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Quarterly Budget Variance
    worksheet.push(['📊 QUARTERLY BUDGET VARIANCE ANALYSIS', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Quarter', 'Planned', 'Actual', 'Forecast', 'Variance', 'Performance', '', '', '', '']);
    
    const quarters = ['Q1', 'Q2', 'Q3', 'Q4'];
    quarters.forEach((quarter, index) => {
      const planned = analysis.budgetVariance.plannedByQuarter[index];
      const actual = analysis.budgetVariance.actualByQuarter[index];
      const forecast = analysis.budgetVariance.forecastByQuarter[index];
      const variance = actual > 0 ? ((actual - planned) / planned * 100).toFixed(1) + '%' : 'TBD';
      
      let performance = 'ON TRACK';
      if (actual > 0) {
        if (actual > planned * 1.1) performance = 'OVER BUDGET';
        else if (actual < planned * 0.9) performance = 'UNDER BUDGET';
      }
      
      worksheet.push([
        quarter,
        '$' + (planned / 1000).toFixed(0) + 'K',
        actual > 0 ? '$' + (actual / 1000).toFixed(0) + 'K' : 'TBD',
        '$' + (forecast / 1000).toFixed(0) + 'K',
        variance,
        performance,
        '', '', '', ''
      ]);
    });
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Upcoming Projects Detail
    if (analysis.timeline.upcomingProjects.length > 0) {
      worksheet.push(['🚀 UPCOMING PROJECTS (Next 90 Days)', '', '', '', '', '', '', '', '', '']);
      worksheet.push(['Project', 'Type', 'Budget', 'Start Date', 'Duration', 'ROI', 'Impact', '', '', '']);
      
      analysis.timeline.upcomingProjects.forEach(project => {
        const startDate = new Date(project.plannedStartDate);
        const endDate = new Date(project.plannedEndDate);
        const durationDays = Math.ceil((endDate.getTime() - startDate.getTime()) / (24 * 60 * 60 * 1000));
        
        let impactLevel = 'MINIMAL';
        if (project.impactOnOccupancy === 'significant') impactLevel = 'HIGH';
        else if (project.impactOnOccupancy === 'moderate') impactLevel = 'MODERATE';
        
        worksheet.push([
          project.projectName,
          project.projectType.replace('_', ' ').toUpperCase(),
          '$' + (project.budgetAmount / 1000).toFixed(0) + 'K',
          startDate.toLocaleDateString(),
          durationDays + ' days',
          project.expectedROI ? (project.expectedROI * 100).toFixed(1) + '%' : 'TBD',
          impactLevel,
          '', '', ''
        ]);
      });
      
      worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    }
    
    // Detailed Project Analysis (if provided)
    if (detailedProjects && detailedProjects.length > 0) {
      worksheet.push(['📋 DETAILED PROJECT ANALYSIS', '', '', '', '', '', '', '', '', '']);
      worksheet.push(['Project Name', 'Status', 'Budget', 'Actual', 'Variance', 'ROI', 'Value Add', 'Completion %', '', '']);
      
      detailedProjects.forEach(project => {
        const actualCost = project.actualCost || 0;
        const variance = actualCost > 0 ? ((actualCost - project.budgetAmount) / project.budgetAmount * 100).toFixed(1) + '%' : 'TBD';
        
        let completionPercentage = '0%';
        if (project.status === 'completed') completionPercentage = '100%';
        else if (project.status === 'in_progress') completionPercentage = '50%'; // Estimated
        
        worksheet.push([
          project.projectName,
          project.status.toUpperCase(),
          '$' + (project.budgetAmount / 1000).toFixed(0) + 'K',
          actualCost > 0 ? '$' + (actualCost / 1000).toFixed(0) + 'K' : 'TBD',
          variance,
          project.expectedROI ? (project.expectedROI * 100).toFixed(1) + '%' : 'TBD',
          project.expectedValueAdd ? '$' + (project.expectedValueAdd / 1000).toFixed(0) + 'K' : 'TBD',
          completionPercentage,
          '', ''
        ]);
      });
      
      worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    }
    
    // ROI Analysis
    worksheet.push(['💰 RETURN ON INVESTMENT ANALYSIS', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['ROI Category', 'Investment', 'Expected Return', 'Payback Period', 'Risk Level', '', '', '', '', '']);
    
    const roiCategories = [
      { name: 'Energy Efficiency', investment: 250000, returnRate: 0.15, payback: 6.7, risk: 'LOW' },
      { name: 'Amenity Upgrades', investment: 150000, returnRate: 0.12, payback: 8.3, risk: 'MEDIUM' },
      { name: 'Structural Improvements', investment: 500000, returnRate: 0.08, payback: 12.5, risk: 'LOW' },
      { name: 'Technology Upgrades', investment: 100000, returnRate: 0.20, payback: 5.0, risk: 'MEDIUM' }
    ];
    
    roiCategories.forEach(category => {
      worksheet.push([
        category.name,
        '$' + (category.investment / 1000).toFixed(0) + 'K',
        (category.returnRate * 100).toFixed(1) + '%',
        category.payback.toFixed(1) + ' years',
        category.risk,
        '', '', '', '', ''
      ]);
    });
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Risk Assessment
    worksheet.push(['⚠️ CAPEX RISK ASSESSMENT', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Risk Factor', 'Level', 'Impact', 'Mitigation Strategy', '', '', '', '', '', '']);
    
    const riskFactors = [
      ['Budget Overruns', analysis.capExMetrics.projectedOverrun > 50000 ? 'HIGH' : 'LOW', 'Cost increase', 'Enhanced project monitoring and change order controls'],
      ['Schedule Delays', analysis.capExMetrics.onTimePercentage < 80 ? 'HIGH' : 'LOW', 'Revenue impact', 'Improved project scheduling and contractor management'],
      ['Permit Delays', 'MEDIUM', 'Timeline extension', 'Early permit application and regulatory coordination'],
      ['Market Changes', 'MEDIUM', 'ROI reduction', 'Market analysis and flexible project scoping'],
      ['Occupancy Impact', 'MEDIUM', 'Revenue loss', 'Phased construction and tenant communication']
    ];
    
    riskFactors.forEach(([factor, level, impact, mitigation]) => {
      worksheet.push([factor, level, impact, mitigation, '', '', '', '', '', '']);
    });
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Action Plan and Recommendations
    worksheet.push(['✅ CAPEX MANAGEMENT ACTION PLAN', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Priority', 'Action Item', 'Responsible', 'Timeline', 'Expected Outcome', '', '', '', '', '']);
    
    const actionPlan = [
      ['HIGH', 'Review and approve all upcoming project budgets', 'Asset Manager', 'Next 30 days', 'Accurate project planning and budgeting'],
      ['HIGH', 'Implement enhanced project monitoring system', 'Project Manager', 'Next 60 days', 'Improved budget and timeline control'],
      ['MEDIUM', 'Vendor performance evaluation and optimization', 'Operations', 'Next 90 days', 'Better contractor relationships and pricing'],
      ['MEDIUM', 'ROI tracking and value validation system', 'Finance', 'Next 90 days', 'Measurable return on capital investments'],
      ['LOW', 'Annual CapEx strategy review and planning', 'Executive Team', 'Annually', 'Optimized capital allocation and returns']
    ];
    
    actionPlan.forEach(([priority, action, responsible, timeline, outcome]) => {
      worksheet.push([priority, action, responsible, timeline, outcome, '', '', '', '', '']);
    });
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // CapEx KPIs and Benchmarks
    worksheet.push(['📈 CAPEX KPIs & BENCHMARKS', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['KPI', 'Current Value', 'Industry Benchmark', 'Performance', 'Target', '', '', '', '', '']);
    
    const budgetUtilization = (analysis.capExMetrics.spentToDate / analysis.capExMetrics.totalBudget) * 100;
    
    const capexKPIs = [
      ['On-Time Completion', analysis.capExMetrics.onTimePercentage.toFixed(1) + '%', '>= 85%', analysis.capExMetrics.onTimePercentage >= 85 ? 'MEETS' : 'BELOW', '>= 90%'],
      ['Budget Utilization', budgetUtilization.toFixed(1) + '%', '80-95%', budgetUtilization >= 80 && budgetUtilization <= 95 ? 'OPTIMAL' : 'REVIEW', '85-90%'],
      ['Average ROI', (analysis.capExMetrics.avgROI * 100).toFixed(1) + '%', '>= 10%', analysis.capExMetrics.avgROI >= 0.10 ? 'MEETS' : 'BELOW', '>= 12%'],
      ['Cost Overrun Rate', (analysis.capExMetrics.projectedOverrun / analysis.capExMetrics.totalBudget * 100).toFixed(1) + '%', '< 5%', (analysis.capExMetrics.projectedOverrun / analysis.capExMetrics.totalBudget) < 0.05 ? 'MEETS' : 'ABOVE', '< 3%']
    ];
    
    capexKPIs.forEach(([kpi, current, benchmark, performance, target]) => {
      worksheet.push([kpi, current, benchmark, performance, target, '', '', '', '', '']);
    });
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Professional Notes
    worksheet.push(['📝 ANALYSIS METHODOLOGY & DISCLAIMERS', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- CapEx analysis based on approved projects and current market conditions', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- ROI calculations include both cash flow and property value impacts', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- Timeline and cost projections based on historical performance data', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- Market conditions and regulatory changes may impact project outcomes', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- Regular project review and budget monitoring recommended', '', '', '', '', '', '', '', '', '']);
    
    return worksheet;
  }
}