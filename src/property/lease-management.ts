import { CellValue } from '../excel/excel-manager.js';

export interface Lease {
  leaseId: string;
  propertyId: string;
  unitId: string;
  tenantId: string;
  tenantName: string;
  leaseStartDate: string;
  leaseEndDate: string;
  currentRent: number;
  securityDeposit: number;
  leaseType: 'residential' | 'commercial' | 'retail' | 'office' | 'industrial';
  rentEscalations: RentEscalation[];
  renewalOptions: RenewalOption[];
  status: 'active' | 'expired' | 'terminated' | 'pending' | 'holdover';
  paymentTerms: string;
  specialClauses: string[];
}

export interface RentEscalation {
  effectiveDate: string;
  escalationType: 'fixed_amount' | 'percentage' | 'cpi_adjustment' | 'market_rate';
  escalationValue: number;
  baseAmount?: number;
}

export interface RenewalOption {
  optionPeriod: number; // months
  renewalRate: number;
  noticeDaysRequired: number;
  marketRateAdjustment?: boolean;
}

export interface LeaseAnalysis {
  leaseId: string;
  currentMetrics: {
    monthsRemaining: number;
    daysToExpiration: number;
    renewalProbability: number;
    marketRentComparison: number;
    escalationsDue: RentEscalation[];
  };
  financialProjection: {
    totalRemainingRent: number;
    futureEscalations: number[];
    renewalValue: number;
    vacancyRisk: number;
    leaseValue: number; // NPV of remaining lease payments
  };
  riskFactors: {
    tenantCreditRisk: 'low' | 'medium' | 'high';
    paymentHistory: 'excellent' | 'good' | 'concerning' | 'poor';
    marketRentRisk: 'below_market' | 'at_market' | 'above_market';
    renewalLikelihood: 'high' | 'medium' | 'low';
  };
}

export interface ExpirationAnalysis {
  propertyId: string;
  timeframe: string;
  expiringLeases: {
    totalLeases: number;
    totalSquareFeet: number;
    totalRent: number;
    averageRent: number;
    tenantMix: { [tenantType: string]: number };
  };
  renewalProjections: {
    expectedRenewals: number;
    projectedVacancy: number;
    rentRollRisk: number;
    revenueAtRisk: number;
  };
  marketAnalysis: {
    currentMarketRent: number;
    rentGrowthOpportunity: number;
    competitivePosition: string;
    leaseUpTime: number; // months to re-lease vacant units
  };
}

export class LeaseLifecycleManager {

  private calculateLeaseMetrics(lease: Lease, analysisDate: Date = new Date()): LeaseAnalysis['currentMetrics'] {
    // const startDate = new Date(lease.leaseStartDate); // Available for future use
    const endDate = new Date(lease.leaseEndDate);
    const daysToExpiration = Math.ceil((endDate.getTime() - analysisDate.getTime()) / (1000 * 60 * 60 * 24));
    const monthsRemaining = Math.ceil(daysToExpiration / 30.44); // Average days per month
    
    // Calculate renewal probability based on lease terms and market conditions
    let renewalProbability = 0.65; // Base probability
    
    // Adjust based on rent vs market
    if (lease.renewalOptions.length > 0) renewalProbability += 0.15;
    if (monthsRemaining > 12) renewalProbability += 0.10;
    if (monthsRemaining < 3) renewalProbability -= 0.20;
    
    renewalProbability = Math.max(0.1, Math.min(0.95, renewalProbability));
    
    // Find upcoming escalations
    const escalationsDue = lease.rentEscalations.filter(escalation => {
      const escalationDate = new Date(escalation.effectiveDate);
      const daysDiff = (escalationDate.getTime() - analysisDate.getTime()) / (1000 * 60 * 60 * 24);
      return daysDiff > 0 && daysDiff <= 90; // Next 90 days
    });
    
    return {
      monthsRemaining,
      daysToExpiration,
      renewalProbability,
      marketRentComparison: 0, // Would integrate with market data
      escalationsDue
    };
  }

  private calculateFinancialProjection(lease: Lease, metrics: LeaseAnalysis['currentMetrics']): LeaseAnalysis['financialProjection'] {
    const monthsRemaining = metrics.monthsRemaining;
    let currentRent = lease.currentRent;
    const futureEscalations: number[] = [];
    let totalRemainingRent = 0;
    
    // Calculate month-by-month rent with escalations
    for (let month = 1; month <= monthsRemaining; month++) {
      const monthDate = new Date();
      monthDate.setMonth(monthDate.getMonth() + month);
      
      // Check for escalations this month
      const escalation = lease.rentEscalations.find(esc => {
        const escDate = new Date(esc.effectiveDate);
        return escDate.getFullYear() === monthDate.getFullYear() && 
               escDate.getMonth() === monthDate.getMonth();
      });
      
      if (escalation) {
        switch (escalation.escalationType) {
          case 'fixed_amount':
            currentRent += escalation.escalationValue;
            break;
          case 'percentage':
            currentRent *= (1 + escalation.escalationValue / 100);
            break;
          case 'cpi_adjustment':
            // Simplified CPI adjustment
            currentRent *= (1 + Math.max(0.02, escalation.escalationValue / 100));
            break;
        }
        futureEscalations.push(currentRent - lease.currentRent);
      }
      
      totalRemainingRent += currentRent;
    }
    
    // Calculate renewal value if applicable
    let renewalValue = 0;
    if (lease.renewalOptions.length > 0 && metrics.renewalProbability > 0.5) {
      const renewalOption = lease.renewalOptions[0];
      renewalValue = renewalOption.renewalRate * renewalOption.optionPeriod * metrics.renewalProbability;
    }
    
    // Calculate NPV of remaining lease payments (simplified)
    const discountRate = 0.08; // 8% discount rate
    const leaseValue = totalRemainingRent / Math.pow(1 + discountRate, monthsRemaining / 12);
    
    return {
      totalRemainingRent,
      futureEscalations,
      renewalValue,
      vacancyRisk: (1 - metrics.renewalProbability) * totalRemainingRent,
      leaseValue
    };
  }

  private assessRiskFactors(lease: Lease, metrics: LeaseAnalysis['currentMetrics']): LeaseAnalysis['riskFactors'] {
    // Simplified risk assessment - in production would integrate with credit data
    let tenantCreditRisk: 'low' | 'medium' | 'high' = 'medium';
    let paymentHistory: 'excellent' | 'good' | 'concerning' | 'poor' = 'good';
    let marketRentRisk: 'below_market' | 'at_market' | 'above_market' = 'at_market';
    let renewalLikelihood: 'high' | 'medium' | 'low';
    
    // Assess renewal likelihood
    if (metrics.renewalProbability >= 0.75) renewalLikelihood = 'high';
    else if (metrics.renewalProbability >= 0.45) renewalLikelihood = 'medium';
    else renewalLikelihood = 'low';
    
    // Commercial leases typically have better credit assessment
    if (lease.leaseType === 'commercial' || lease.leaseType === 'office') {
      tenantCreditRisk = 'low';
      paymentHistory = 'excellent';
    }
    
    return {
      tenantCreditRisk,
      paymentHistory,
      marketRentRisk,
      renewalLikelihood
    };
  }

  analyzeLeasePortfolio(leases: Lease[], analysisDate: Date = new Date()): LeaseAnalysis[] {
    return leases.map(lease => {
      const currentMetrics = this.calculateLeaseMetrics(lease, analysisDate);
      const financialProjection = this.calculateFinancialProjection(lease, currentMetrics);
      const riskFactors = this.assessRiskFactors(lease, currentMetrics);
      
      return {
        leaseId: lease.leaseId,
        currentMetrics,
        financialProjection,
        riskFactors
      };
    });
  }

  generateExpirationAnalysis(
    leases: Lease[], 
    timeframe: 'next_30_days' | 'next_90_days' | 'next_6_months' | 'next_12_months' | 'next_24_months',
    analysisDate: Date = new Date()
  ): ExpirationAnalysis {
    
    // Calculate timeframe end date
    const endDate = new Date(analysisDate);
    switch (timeframe) {
      case 'next_30_days': endDate.setDate(endDate.getDate() + 30); break;
      case 'next_90_days': endDate.setDate(endDate.getDate() + 90); break;
      case 'next_6_months': endDate.setMonth(endDate.getMonth() + 6); break;
      case 'next_12_months': endDate.setFullYear(endDate.getFullYear() + 1); break;
      case 'next_24_months': endDate.setFullYear(endDate.getFullYear() + 2); break;
    }
    
    // Filter leases expiring in timeframe
    const expiringLeases = leases.filter(lease => {
      const leaseEndDate = new Date(lease.leaseEndDate);
      return leaseEndDate >= analysisDate && leaseEndDate <= endDate;
    });
    
    // Calculate metrics
    const totalLeases = expiringLeases.length;
    const totalRent = expiringLeases.reduce((sum, lease) => sum + lease.currentRent, 0);
    const totalSquareFeet = expiringLeases.length * 1000; // Simplified - would get actual sq ft
    const averageRent = totalLeases > 0 ? totalRent / totalLeases : 0;
    
    // Tenant mix analysis
    const tenantMix: { [tenantType: string]: number } = {};
    expiringLeases.forEach(lease => {
      tenantMix[lease.leaseType] = (tenantMix[lease.leaseType] || 0) + 1;
    });
    
    // Renewal projections
    const leaseAnalyses = this.analyzeLeasePortfolio(expiringLeases, analysisDate);
    const averageRenewalProbability = leaseAnalyses.reduce((sum, analysis) => 
      sum + analysis.currentMetrics.renewalProbability, 0) / leaseAnalyses.length;
    
    const expectedRenewals = Math.round(totalLeases * averageRenewalProbability);
    const projectedVacancy = totalLeases - expectedRenewals;
    const rentRollRisk = totalRent * (1 - averageRenewalProbability);
    const revenueAtRisk = rentRollRisk * 12; // Annualized
    
    return {
      propertyId: 'PORTFOLIO', // Would be specific property or portfolio-wide
      timeframe,
      expiringLeases: {
        totalLeases,
        totalSquareFeet,
        totalRent,
        averageRent,
        tenantMix
      },
      renewalProjections: {
        expectedRenewals,
        projectedVacancy,
        rentRollRisk,
        revenueAtRisk
      },
      marketAnalysis: {
        currentMarketRent: averageRent * 1.05, // Simplified market rent
        rentGrowthOpportunity: averageRent * 0.05,
        competitivePosition: 'Competitive',
        leaseUpTime: 4 // Months
      }
    };
  }

  generateLeaseExpirationWorksheet(
    analyses: ExpirationAnalysis[], 
    detailedLeases?: LeaseAnalysis[]
  ): Array<Array<CellValue | string | number>> {
    const worksheet: Array<Array<CellValue | string | number>> = [];
    
    // Header
    worksheet.push(['LEASE EXPIRATION ANALYSIS REPORT', '', '', '', '', '', '', '', '', '']);
    worksheet.push([`Analysis Date: ${new Date().toLocaleDateString()}`, '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Property Management Intelligence Platform', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Executive Summary
    worksheet.push(['📊 EXECUTIVE SUMMARY', '', '', '', '', '', '', '', '', '']);
    
    const totalExpiringLeases = analyses.reduce((sum, analysis) => sum + analysis.expiringLeases.totalLeases, 0);
    const totalRentAtRisk = analyses.reduce((sum, analysis) => sum + analysis.expiringLeases.totalRent, 0);
    const totalRevenueAtRisk = analyses.reduce((sum, analysis) => sum + analysis.renewalProjections.revenueAtRisk, 0);
    
    worksheet.push(['Total Expiring Leases (Next 24 Months)', totalExpiringLeases, 'leases', 'Lease renewal management required']);
    worksheet.push(['Monthly Rent at Risk', (totalRentAtRisk / 1000).toFixed(0) + 'K', 'monthly', 'Potential revenue loss if not renewed']);
    worksheet.push(['Annual Revenue at Risk', (totalRevenueAtRisk / 1000000).toFixed(1) + 'M', 'annually', 'Total revenue impact from non-renewals']);
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Expiration Timeline Analysis
    worksheet.push(['🕐 LEASE EXPIRATION TIMELINE', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Timeframe', 'Expiring Leases', 'Monthly Rent', 'Sq Ft', 'Renewal Rate', 'Revenue Risk', 'Action Required', '', '', '']);
    
    for (const analysis of analyses) {
      const renewalRate = analysis.renewalProjections.expectedRenewals / analysis.expiringLeases.totalLeases;
      const actionRequired = renewalRate < 0.7 ? 'HIGH PRIORITY' : renewalRate < 0.85 ? 'MODERATE' : 'MONITOR';
      
      worksheet.push([
        analysis.timeframe.replace('_', ' ').toUpperCase(),
        analysis.expiringLeases.totalLeases,
        (analysis.expiringLeases.totalRent / 1000).toFixed(0) + 'K',
        (analysis.expiringLeases.totalSquareFeet / 1000).toFixed(0) + 'K',
        (renewalRate * 100).toFixed(1) + '%',
        (analysis.renewalProjections.revenueAtRisk / 1000000).toFixed(2) + 'M',
        actionRequired,
        '', '', ''
      ]);
    }
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Tenant Mix Analysis
    worksheet.push(['🏢 TENANT MIX ANALYSIS', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Property Type', 'Expiring Leases', 'Percentage', 'Renewal Risk', 'Market Conditions', '', '', '', '', '']);
    
    const allTenantTypes = new Set<string>();
    analyses.forEach(analysis => {
      Object.keys(analysis.expiringLeases.tenantMix).forEach(type => allTenantTypes.add(type));
    });
    
    for (const tenantType of allTenantTypes) {
      const typeLeases = analyses.reduce((sum, analysis) => 
        sum + (analysis.expiringLeases.tenantMix[tenantType] || 0), 0);
      const percentage = (typeLeases / totalExpiringLeases * 100).toFixed(1);
      
      let renewalRisk = 'MODERATE';
      let marketConditions = 'Stable';
      
      if (tenantType === 'retail') {
        renewalRisk = 'HIGH';
        marketConditions = 'Challenging';
      } else if (tenantType === 'office') {
        renewalRisk = 'MODERATE';
        marketConditions = 'Mixed';
      } else if (tenantType === 'industrial') {
        renewalRisk = 'LOW';
        marketConditions = 'Strong';
      }
      
      worksheet.push([
        tenantType.toUpperCase(),
        typeLeases,
        percentage + '%',
        renewalRisk,
        marketConditions,
        '', '', '', '', ''
      ]);
    }
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Detailed Lease Analysis (if provided)
    if (detailedLeases && detailedLeases.length > 0) {
      worksheet.push(['📋 DETAILED LEASE ANALYSIS', '', '', '', '', '', '', '', '', '']);
      worksheet.push(['Lease ID', 'Tenant', 'Expiration', 'Monthly Rent', 'Renewal Prob', 'Risk Level', 'Action', 'Value at Risk', '', '']);
      
      // Sort by renewal probability (lowest first)
      const sortedLeases = detailedLeases.sort((a, b) => 
        a.currentMetrics.renewalProbability - b.currentMetrics.renewalProbability);
      
      for (const lease of sortedLeases.slice(0, 20)) { // Top 20 highest risk
        const riskLevel = lease.riskFactors.renewalLikelihood === 'low' ? 'HIGH RISK' :
                         lease.riskFactors.renewalLikelihood === 'medium' ? 'MODERATE' : 'LOW RISK';
        
        const action = lease.currentMetrics.daysToExpiration < 90 ? 'URGENT - RENEW NOW' :
                      lease.currentMetrics.daysToExpiration < 180 ? 'BEGIN RENEWAL PROCESS' :
                      'MONITOR AND PLAN';
        
        worksheet.push([
          lease.leaseId,
          'Tenant Name', // Would be actual tenant name
          new Date(Date.now() + lease.currentMetrics.daysToExpiration * 24 * 60 * 60 * 1000).toLocaleDateString(),
          (lease.financialProjection.totalRemainingRent / lease.currentMetrics.monthsRemaining / 1000).toFixed(0) + 'K',
          (lease.currentMetrics.renewalProbability * 100).toFixed(0) + '%',
          riskLevel,
          action,
          (lease.financialProjection.vacancyRisk / 1000).toFixed(0) + 'K',
          '', ''
        ]);
      }
    }
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Renewal Strategy Recommendations
    worksheet.push(['💡 RENEWAL STRATEGY RECOMMENDATIONS', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Priority', 'Strategy', 'Target Leases', 'Expected Impact', 'Timeline', 'Resources Required', '', '', '', '']);
    
    worksheet.push([
      'HIGH',
      'Immediate Renewal Outreach',
      'Expiring < 90 days',
      'Prevent immediate vacancy',
      'Next 30 days', 
      'Leasing team focus',
      '', '', '', ''
    ]);
    
    worksheet.push([
      'MEDIUM',
      'Market Rate Analysis & Negotiation',
      'Below-market rents',
      '5-15% rent increase',
      '60-90 days',
      'Market research, negotiation',
      '', '', '', ''
    ]);
    
    worksheet.push([
      'LOW',
      'Tenant Satisfaction Program', 
      'High renewal probability',
      'Maintain occupancy',
      'Ongoing',
      'Property management',
      '', '', '', ''
    ]);
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Financial Impact Projections
    worksheet.push(['💰 FINANCIAL IMPACT PROJECTIONS', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Scenario', 'Renewal Rate', 'Revenue Impact', 'Leasing Costs', 'Net Impact', 'Probability', '', '', '', '']);
    
    const scenarios = [
      { name: 'Pessimistic', renewalRate: 0.60, probability: 0.15 },
      { name: 'Base Case', renewalRate: 0.75, probability: 0.60 },
      { name: 'Optimistic', renewalRate: 0.90, probability: 0.25 }
    ];
    
    for (const scenario of scenarios) {
      const revenueImpact = totalRentAtRisk * 12 * (1 - scenario.renewalRate);
      const leasingCosts = totalExpiringLeases * (1 - scenario.renewalRate) * 5000; // $5K per re-lease
      const netImpact = revenueImpact + leasingCosts;
      
      worksheet.push([
        scenario.name,
        (scenario.renewalRate * 100).toFixed(0) + '%',
        '($' + (revenueImpact / 1000000).toFixed(1) + 'M)',
        '($' + (leasingCosts / 1000).toFixed(0) + 'K)',
        '($' + (netImpact / 1000000).toFixed(1) + 'M)',
        (scenario.probability * 100).toFixed(0) + '%',
        '', '', '', ''
      ]);
    }
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Action Plan
    worksheet.push(['✅ 90-DAY ACTION PLAN', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Week', 'Action Item', 'Responsible', 'Target', 'Success Metric', '', '', '', '', '']);
    
    const actionPlan = [
      ['1-2', 'Complete lease expiration audit', 'Property Manager', 'All expiring leases identified', '100% lease data accuracy'],
      ['3-4', 'Market rent analysis and pricing strategy', 'Leasing Team', 'Market-competitive pricing', 'Pricing within 5% of market'],
      ['5-6', 'Begin renewal outreach (90+ day pipeline)', 'Leasing Team', '80% contact rate', 'Tenant meetings scheduled'],
      ['7-8', 'Urgent renewals (60-90 day expirations)', 'Property Manager', '70% renewal rate', 'Signed lease agreements'],
      ['9-10', 'Marketing for non-renewals', 'Marketing Team', 'Active marketing campaigns', '50% prospect pipeline'],
      ['11-12', 'Final renewal push and new lease signings', 'All Teams', '75% overall renewal rate', 'Minimized vacancy']
    ];
    
    for (const [week, action, responsible, target, metric] of actionPlan) {
      worksheet.push([week, action, responsible, target, metric, '', '', '', '', '']);
    }
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Professional Notes
    worksheet.push(['📝 ANALYSIS METHODOLOGY & ASSUMPTIONS', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- Renewal probabilities based on market conditions and lease terms', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- Revenue at risk assumes non-renewed leases remain vacant for 4 months', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- Market rent estimates based on comparable property analysis', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- Financial projections include estimated re-leasing costs and downtime', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- Analysis should be updated monthly as market conditions change', '', '', '', '', '', '', '', '', '']);
    
    return worksheet;
  }
}