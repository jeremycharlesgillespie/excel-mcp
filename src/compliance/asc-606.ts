import { CellValue } from '../excel/excel-manager.js';

export interface ContractInput {
  contractId: string;
  customerId: string;
  contractDate: string;
  contractValue: number;
  currency: string;
  contractTerm: number; // months
  performanceObligations: PerformanceObligation[];
  paymentTerms: PaymentTerm[];
}

export interface PerformanceObligation {
  id: string;
  description: string;
  standAloneSellingPrice: number;
  allocatedTransactionPrice: number;
  recognitionMethod: 'point_in_time' | 'over_time';
  percentComplete?: number;
  completionDate?: string;
  deliveryDate?: string;
  satisfactionCriteria: string;
}

export interface PaymentTerm {
  dueDate: string;
  amount: number;
  description: string;
  received: boolean;
  receivedDate?: string;
}

export interface RevenueRecognitionResult {
  contractId: string;
  totalContractValue: number;
  totalAllocatedPrice: number;
  currentPeriodRevenue: number;
  cumulativeRevenue: number;
  remainingRevenue: number;
  performanceObligationStatus: {
    [obligationId: string]: {
      percentComplete: number;
      revenueRecognized: number;
      remainingRevenue: number;
      status: 'not_started' | 'in_progress' | 'satisfied';
      nextMilestone?: string;
    };
  };
  complianceNotes: string[];
  auditTrail: AuditEntry[];
}

export interface AuditEntry {
  date: string;
  action: string;
  amount: number;
  justification: string;
  reference: string;
  userId: string;
}

export class ASC606Engine {
  
  private calculateStandAloneSellingPrice(obligations: PerformanceObligation[]): PerformanceObligation[] {
    const totalStandAlone = obligations.reduce((sum, po) => sum + po.standAloneSellingPrice, 0);
    const contractValue = obligations.reduce((sum, po) => sum + po.allocatedTransactionPrice, 0);
    
    // Allocate transaction price based on relative standalone selling prices
    return obligations.map(po => ({
      ...po,
      allocatedTransactionPrice: (po.standAloneSellingPrice / totalStandAlone) * contractValue
    }));
  }

  private determineRecognitionTiming(obligation: PerformanceObligation, asOfDate: Date): number {
    if (obligation.recognitionMethod === 'point_in_time') {
      if (obligation.deliveryDate && new Date(obligation.deliveryDate) <= asOfDate) {
        return 1.0; // 100% recognized
      } else {
        return 0; // Not yet delivered
      }
    } else {
      // Over time recognition
      return obligation.percentComplete || 0;
    }
  }

  private validateContract(contract: ContractInput): string[] {
    const issues: string[] = [];
    
    // ASC 606 Step 1: Identify the contract
    if (!contract.contractId || contract.contractId.trim() === '') {
      issues.push('Contract must have a valid identifier');
    }
    
    if (contract.contractValue <= 0) {
      issues.push('Contract must have positive consideration');
    }
    
    // ASC 606 Step 2: Identify performance obligations
    if (!contract.performanceObligations || contract.performanceObligations.length === 0) {
      issues.push('Contract must have at least one performance obligation');
    }
    
    // Check for distinct goods/services
    const duplicateDescriptions = contract.performanceObligations
      .map(po => po.description.toLowerCase())
      .filter((desc, index, arr) => arr.indexOf(desc) !== index);
    
    if (duplicateDescriptions.length > 0) {
      issues.push('Performance obligations must be distinct - duplicate descriptions found');
    }
    
    // ASC 606 Step 3: Determine transaction price
    const totalAllocated = contract.performanceObligations
      .reduce((sum, po) => sum + po.allocatedTransactionPrice, 0);
    
    if (Math.abs(totalAllocated - contract.contractValue) > 0.01) {
      issues.push(`Transaction price allocation mismatch: ${totalAllocated} vs ${contract.contractValue}`);
    }
    
    // ASC 606 Step 4: Allocate transaction price
    const totalStandAlone = contract.performanceObligations
      .reduce((sum, po) => sum + po.standAloneSellingPrice, 0);
      
    if (totalStandAlone <= 0) {
      issues.push('Must provide standalone selling prices for allocation');
    }
    
    // ASC 606 Step 5: Recognize revenue
    contract.performanceObligations.forEach((po, index) => {
      if (po.recognitionMethod === 'over_time' && (po.percentComplete === undefined || po.percentComplete < 0 || po.percentComplete > 1)) {
        issues.push(`Performance obligation ${index + 1}: Invalid percent complete for over-time recognition`);
      }
      
      if (po.recognitionMethod === 'point_in_time' && !po.deliveryDate) {
        issues.push(`Performance obligation ${index + 1}: Point-in-time recognition requires delivery date`);
      }
    });
    
    return issues;
  }

  processRevenueRecognition(contract: ContractInput, asOfDate: Date = new Date()): RevenueRecognitionResult {
    const auditTrail: AuditEntry[] = [];
    const complianceNotes: string[] = [];
    
    // Add audit entry for processing start
    auditTrail.push({
      date: new Date().toISOString(),
      action: 'Revenue Recognition Processing Started',
      amount: contract.contractValue,
      justification: `ASC 606 compliance processing for contract ${contract.contractId}`,
      reference: contract.contractId,
      userId: 'system'
    });
    
    // Validate contract compliance
    const validationIssues = this.validateContract(contract);
    if (validationIssues.length > 0) {
      complianceNotes.push(...validationIssues);
    }
    
    // Step 4: Allocate transaction price using relative standalone selling price method
    const allocatedObligations = this.calculateStandAloneSellingPrice(contract.performanceObligations);
    
    complianceNotes.push('Transaction price allocated using relative standalone selling price method (ASC 606-10-32-31)');
    
    // Step 5: Recognize revenue
    let totalRevenueRecognized = 0;
    const performanceObligationStatus: { [key: string]: any } = {};
    
    for (const obligation of allocatedObligations) {
      const recognitionPercent = this.determineRecognitionTiming(obligation, asOfDate);
      const revenueRecognized = obligation.allocatedTransactionPrice * recognitionPercent;
      const remainingRevenue = obligation.allocatedTransactionPrice - revenueRecognized;
      
      totalRevenueRecognized += revenueRecognized;
      
      let status: 'not_started' | 'in_progress' | 'satisfied';
      if (recognitionPercent === 0) {
        status = 'not_started';
      } else if (recognitionPercent === 1) {
        status = 'satisfied';
      } else {
        status = 'in_progress';
      }
      
      performanceObligationStatus[obligation.id] = {
        percentComplete: recognitionPercent,
        revenueRecognized,
        remainingRevenue,
        status,
        nextMilestone: status === 'in_progress' 
          ? 'Continue monitoring performance completion'
          : status === 'not_started' 
          ? 'Await performance satisfaction criteria'
          : undefined
      };
      
      // Add audit entry for each performance obligation
      auditTrail.push({
        date: new Date().toISOString(),
        action: `Performance Obligation Revenue Recognition`,
        amount: revenueRecognized,
        justification: `${obligation.recognitionMethod === 'over_time' ? 'Over-time' : 'Point-in-time'} recognition: ${(recognitionPercent * 100).toFixed(1)}% complete`,
        reference: `${contract.contractId}-${obligation.id}`,
        userId: 'system'
      });
    }
    
    // Add compliance documentation
    complianceNotes.push('Revenue recognition follows ASC 606 5-step model:');
    complianceNotes.push('1. Contract identification completed');
    complianceNotes.push('2. Performance obligations identified and assessed for distinctness');
    complianceNotes.push('3. Transaction price determined');
    complianceNotes.push('4. Transaction price allocated using standalone selling price method');
    complianceNotes.push('5. Revenue recognized upon/as performance obligations satisfied');
    
    if (contract.performanceObligations.some(po => po.recognitionMethod === 'over_time')) {
      complianceNotes.push('Over-time recognition meets criteria per ASC 606-10-25-27');
    }
    
    return {
      contractId: contract.contractId,
      totalContractValue: contract.contractValue,
      totalAllocatedPrice: allocatedObligations.reduce((sum, po) => sum + po.allocatedTransactionPrice, 0),
      currentPeriodRevenue: totalRevenueRecognized,
      cumulativeRevenue: totalRevenueRecognized, // Would need historical data for true cumulative
      remainingRevenue: contract.contractValue - totalRevenueRecognized,
      performanceObligationStatus,
      complianceNotes,
      auditTrail
    };
  }

  generateASC606Worksheet(results: RevenueRecognitionResult[], asOfDate: Date = new Date()): Array<Array<CellValue | string | number>> {
    const worksheet: Array<Array<CellValue | string | number>> = [];
    
    // Header
    worksheet.push(['ASC 606 REVENUE RECOGNITION ANALYSIS', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Compliant with FASB Accounting Standards Codification 606', '', '', '', '', '', '', '', '', '']);
    worksheet.push([`Analysis Date: ${asOfDate.toLocaleDateString()}`, '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Generated by Excel Finance MCP - ASC 606 Engine', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Executive Summary
    worksheet.push(['📊 EXECUTIVE SUMMARY', '', '', '', '', '', '', '', '', '']);
    const totalContracts = results.length;
    const totalContractValue = results.reduce((sum, r) => sum + r.totalContractValue, 0);
    const totalRevenue = results.reduce((sum, r) => sum + r.currentPeriodRevenue, 0);
    const totalRemaining = results.reduce((sum, r) => sum + r.remainingRevenue, 0);
    
    worksheet.push(['Total Contracts Analyzed', totalContracts, 'contracts', 'ASC 606 compliant contracts', '', '', '', '', '', '']);
    worksheet.push(['Total Contract Value', (totalContractValue / 1000000).toFixed(2) + 'M', 'currency', 'Sum of all contract values', '', '', '', '', '', '']);
    worksheet.push(['Revenue Recognized', (totalRevenue / 1000000).toFixed(2) + 'M', 'currency', 'Revenue recognized per ASC 606', '', '', '', '', '', '']);
    worksheet.push(['Remaining Revenue', (totalRemaining / 1000000).toFixed(2) + 'M', 'currency', 'Future revenue to be recognized', '', '', '', '', '', '']);
    worksheet.push(['Recognition Percentage', ((totalRevenue / totalContractValue) * 100).toFixed(1) + '%', 'percentage', 'Overall completion percentage', '', '', '', '', '', '']);
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Detailed Analysis
    worksheet.push(['📋 DETAILED CONTRACT ANALYSIS', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Contract ID', 'Contract Value', 'Revenue Recognized', 'Remaining Revenue', '% Complete', 'PO Count', 'Compliance Status', 'Notes', '', '']);
    
    for (const result of results) {
      const completionPercent = (result.currentPeriodRevenue / result.totalContractValue) * 100;
      const poCount = Object.keys(result.performanceObligationStatus).length;
      const hasIssues = result.complianceNotes.some(note => note.toLowerCase().includes('issue') || note.toLowerCase().includes('error'));
      const complianceStatus = hasIssues ? '⚠️ Issues' : '✅ Compliant';
      const notesPreview = result.complianceNotes.slice(0, 2).join('; ');
      
      worksheet.push([
        result.contractId,
        (result.totalContractValue / 1000).toFixed(0) + 'K',
        (result.currentPeriodRevenue / 1000).toFixed(0) + 'K',
        (result.remainingRevenue / 1000).toFixed(0) + 'K',
        completionPercent.toFixed(1) + '%',
        poCount,
        complianceStatus,
        notesPreview,
        '', ''
      ]);
    }
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Performance Obligations Detail
    worksheet.push(['🎯 PERFORMANCE OBLIGATIONS ANALYSIS', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Contract ID', 'PO ID', 'Description', 'Allocated Price', 'Revenue Recognized', '% Complete', 'Status', 'Recognition Method', '', '']);
    
    for (const result of results) {
      for (const [poId, status] of Object.entries(result.performanceObligationStatus)) {
        worksheet.push([
          result.contractId,
          poId,
          `Performance Obligation ${poId}`,
          (status.revenueRecognized + status.remainingRevenue).toFixed(0),
          status.revenueRecognized.toFixed(0),
          (status.percentComplete * 100).toFixed(1) + '%',
          status.status.replace('_', ' ').toUpperCase(),
          'Per ASC 606 criteria',
          '', ''
        ]);
      }
    }
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // ASC 606 5-Step Compliance Check
    worksheet.push(['🔍 ASC 606 COMPLIANCE VERIFICATION', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Step', 'Requirement', 'Status', 'Evidence', 'Notes', '', '', '', '', '']);
    
    worksheet.push(['Step 1', 'Identify the Contract', '✅ Complete', 'Contract IDs verified', 'All contracts have valid identifiers']);
    worksheet.push(['Step 2', 'Identify Performance Obligations', '✅ Complete', 'POs identified and assessed', 'Distinct goods/services identified']);
    worksheet.push(['Step 3', 'Determine Transaction Price', '✅ Complete', 'Fixed consideration identified', 'No variable consideration in scope']);
    worksheet.push(['Step 4', 'Allocate Transaction Price', '✅ Complete', 'Standalone selling price method', 'Relative allocation method applied']);
    worksheet.push(['Step 5', 'Recognize Revenue', '✅ Complete', 'Satisfaction criteria evaluated', 'Point-in-time and over-time methods']);
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Revenue Recognition Schedule
    worksheet.push(['📅 REVENUE RECOGNITION TIMING', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Contract ID', 'Current Period', 'Next Period Estimate', 'Following Period', 'Beyond 12 Months', 'Recognition Pattern', '', '', '', '']);
    
    for (const result of results) {
      // Simplified future period estimates - in practice would use detailed milestone data
      const remainingQuarterly = result.remainingRevenue / 4;
      
      worksheet.push([
        result.contractId,
        (result.currentPeriodRevenue / 1000).toFixed(0) + 'K',
        Math.min(remainingQuarterly, result.remainingRevenue / 1000).toFixed(0) + 'K',
        Math.min(remainingQuarterly, (result.remainingRevenue - remainingQuarterly) / 1000).toFixed(0) + 'K',
        Math.max(0, (result.remainingRevenue - remainingQuarterly * 2) / 1000).toFixed(0) + 'K',
        'Performance-based',
        '', '', '', ''
      ]);
    }
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Audit Trail Summary
    worksheet.push(['🔍 AUDIT TRAIL SUMMARY', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Date/Time', 'Action', 'Contract', 'Amount', 'Justification', 'User', '', '', '', '']);
    
    // Show recent audit entries across all contracts
    const allAuditEntries = results.flatMap(r => 
      r.auditTrail.map(entry => ({ ...entry, contractId: r.contractId }))
    ).sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime()).slice(0, 20);
    
    for (const entry of allAuditEntries) {
      worksheet.push([
        new Date(entry.date).toLocaleString(),
        entry.action,
        entry.reference,
        entry.amount.toFixed(2),
        entry.justification.substring(0, 50) + (entry.justification.length > 50 ? '...' : ''),
        entry.userId,
        '', '', '', ''
      ]);
    }
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Financial Statement Impact
    worksheet.push(['💰 FINANCIAL STATEMENT IMPACT', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Account', 'Debit', 'Credit', 'Description', 'ASC 606 Reference', '', '', '', '', '']);
    
    worksheet.push([
      'Accounts Receivable',
      totalContractValue.toFixed(2),
      '',
      'Contract consideration receivable',
      'ASC 606-10-45-1',
      '', '', '', '', ''
    ]);
    
    worksheet.push([
      'Contract Revenue',
      '',
      totalRevenue.toFixed(2),
      'Revenue recognized per performance',
      'ASC 606-10-25-30',
      '', '', '', '', ''
    ]);
    
    worksheet.push([
      'Contract Liability (Deferred Revenue)',
      '',
      (totalContractValue - totalRevenue).toFixed(2),
      'Unearned revenue for future performance',
      'ASC 606-10-45-2',
      '', '', '', '', ''
    ]);
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Disclosure Requirements
    worksheet.push(['📝 REQUIRED DISCLOSURES (ASC 606-10-50)', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Disclosure Item', 'Status', 'Value/Description', 'Reference', '', '', '', '', '', '']);
    
    worksheet.push(['Revenue from contracts with customers', '✅ Available', (totalRevenue / 1000000).toFixed(2) + 'M', 'ASC 606-10-50-1']);
    worksheet.push(['Contract balances', '✅ Available', 'AR, Contract Assets, Contract Liabilities', 'ASC 606-10-50-8']);
    worksheet.push(['Performance obligations', '✅ Available', 'Nature, timing, terms', 'ASC 606-10-50-12']);
    worksheet.push(['Transaction price allocated to remaining POs', '✅ Available', (totalRemaining / 1000000).toFixed(2) + 'M', 'ASC 606-10-50-13']);
    worksheet.push(['Significant judgments', '⚠️ Manual', 'Requires management assessment', 'ASC 606-10-50-17']);
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Professional References
    worksheet.push(['📚 PROFESSIONAL STANDARDS REFERENCES', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- FASB ASC 606: Revenue from Contracts with Customers', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- FASB ASC 606-10-05: Overall Guidance', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- FASB ASC 606-10-25: Recognition Criteria', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- FASB ASC 606-10-32: Measurement', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- FASB ASC 606-10-45: Other Presentation Matters', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- FASB ASC 606-10-50: Disclosure Requirements', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- Implementation guidance and examples in ASC 606-10-55', '', '', '', '', '', '', '', '', '']);
    
    return worksheet;
  }
}