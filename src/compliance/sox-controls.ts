import { CellValue } from '../excel/excel-manager.js';

export interface SOXControl {
  controlId: string;
  controlName: string;
  controlObjective: string;
  controlType: 'preventive' | 'detective' | 'corrective';
  frequency: 'daily' | 'weekly' | 'monthly' | 'quarterly' | 'annually' | 'event_driven';
  owner: string;
  riskRating: 'high' | 'medium' | 'low';
  process: string;
  controlActivity: string;
  evidence: string[];
  lastTestDate?: string;
  nextTestDate?: string;
  status: 'effective' | 'ineffective' | 'not_tested' | 'requires_remediation';
}

export interface ControlTest {
  testId: string;
  controlId: string;
  testDate: string;
  tester: string;
  testProcedure: string;
  sampleSize: number;
  populationSize: number;
  deficiencies: ControlDeficiency[];
  conclusion: 'effective' | 'ineffective' | 'requires_remediation';
  managementResponse?: string;
  remediationPlan?: string;
  retestDate?: string;
}

export interface ControlDeficiency {
  deficiencyId: string;
  description: string;
  severity: 'significant_deficiency' | 'material_weakness' | 'minor_deficiency';
  rootCause: string;
  potentialImpact: string;
  managementResponse: string;
  remediationDeadline: string;
  status: 'open' | 'in_progress' | 'resolved';
}

export interface SOXComplianceResult {
  compliancePeriod: string;
  totalControls: number;
  testedControls: number;
  effectiveControls: number;
  deficientControls: number;
  controlsByProcess: { [process: string]: number };
  controlsByRisk: { [risk: string]: number };
  deficiencySummary: {
    materialWeaknesses: number;
    significantDeficiencies: number;
    minorDeficiencies: number;
  };
  overallAssessment: 'effective' | 'ineffective' | 'requires_improvement';
  executiveSummary: string[];
  recommendations: string[];
}

export class SOXControlsEngine {

  private calculateSampleSize(populationSize: number, confidenceLevel: number = 0.95): number {
    // Simplified sample size calculation for SOX testing
    // In practice, would use more sophisticated statistical sampling
    
    const zScore = confidenceLevel === 0.95 ? 1.96 : 
                   confidenceLevel === 0.90 ? 1.645 : 1.96;
    
    const margin = 0.05; // 5% margin of error
    const proportion = 0.5; // Conservative estimate
    
    const numerator = zScore * zScore * proportion * (1 - proportion);
    const denominator = margin * margin;
    
    let sampleSize = numerator / denominator;
    
    // Finite population correction
    if (populationSize < 10000) {
      sampleSize = sampleSize / (1 + (sampleSize - 1) / populationSize);
    }
    
    // Minimum sample sizes for SOX compliance
    if (populationSize <= 10) return Math.min(populationSize, 5);
    if (populationSize <= 25) return Math.min(populationSize, 10);
    if (populationSize <= 100) return Math.min(populationSize, 25);
    
    return Math.max(25, Math.ceil(sampleSize));
  }

  // Available for detailed control effectiveness assessment
  // @ts-ignore - Method available for future use
  private assessControlEffectiveness(_control: SOXControl, test: ControlTest): string {
    const deficiencyCount = test.deficiencies.length;
    const errorRate = deficiencyCount / test.sampleSize;
    
    // SOX effectiveness criteria
    if (deficiencyCount === 0) {
      return 'Control operating effectively';
    }
    
    if (errorRate <= 0.05 && !test.deficiencies.some(d => d.severity === 'material_weakness')) {
      return 'Control operating effectively with minor exceptions';
    }
    
    if (test.deficiencies.some(d => d.severity === 'material_weakness')) {
      return 'Material weakness identified - control ineffective';
    }
    
    if (test.deficiencies.some(d => d.severity === 'significant_deficiency') || errorRate > 0.10) {
      return 'Significant deficiency identified - control requires remediation';
    }
    
    return 'Control operating effectively with exceptions to be monitored';
  }

  private generateControlTestProcedures(control: SOXControl): string[] {
    const procedures: string[] = [];
    
    switch (control.controlType) {
      case 'preventive':
        procedures.push('1. Obtain understanding of the preventive control design');
        procedures.push('2. Test control implementation through inquiry and observation');
        procedures.push('3. Select sample of transactions/events to test control operation');
        procedures.push('4. Verify control operates as designed to prevent errors/fraud');
        procedures.push('5. Test exception handling and override procedures');
        break;
        
      case 'detective':
        procedures.push('1. Understand the detective control design and timing');
        procedures.push('2. Test the control\'s ability to detect errors or irregularities');
        procedures.push('3. Verify timely identification and reporting of exceptions');
        procedures.push('4. Test follow-up procedures for detected issues');
        procedures.push('5. Validate control evidence and documentation');
        break;
        
      case 'corrective':
        procedures.push('1. Test the identification and escalation of control failures');
        procedures.push('2. Verify timely corrective actions are implemented');
        procedures.push('3. Test root cause analysis procedures');
        procedures.push('4. Validate remediation effectiveness');
        procedures.push('5. Test prevention of similar future occurrences');
        break;
    }
    
    // Add specific procedures based on control frequency
    if (control.frequency === 'daily' || control.frequency === 'weekly') {
      procedures.push('6. Test multiple periods to ensure consistent operation');
    }
    
    if (control.riskRating === 'high') {
      procedures.push('7. Perform additional substantive testing due to high risk rating');
      procedures.push('8. Test control environment factors supporting this control');
    }
    
    procedures.push('9. Document all testing procedures and results');
    procedures.push('10. Evaluate overall control effectiveness and communicate results');
    
    return procedures;
  }

  testSOXControl(control: SOXControl, testParameters: {
    tester: string;
    testDate?: string;
    sampleSize?: number;
    populationSize: number;
    customProcedures?: string[];
  }): ControlTest {
    
    const testDate = testParameters.testDate || new Date().toISOString().split('T')[0];
    const sampleSize = testParameters.sampleSize || this.calculateSampleSize(testParameters.populationSize);
    
    // Generate standard test procedures
    const procedures = testParameters.customProcedures || this.generateControlTestProcedures(control);
    
    // Simulate control testing (in production, this would interface with actual testing)
    const simulatedDeficiencies: ControlDeficiency[] = [];
    
    // Simulate finding deficiencies based on control risk and effectiveness
    const deficiencyProbability = control.riskRating === 'high' ? 0.15 : 
                                 control.riskRating === 'medium' ? 0.08 : 0.03;
    
    if (Math.random() < deficiencyProbability) {
      const severity = control.riskRating === 'high' && Math.random() < 0.3 ? 'material_weakness' :
                      control.riskRating === 'high' && Math.random() < 0.6 ? 'significant_deficiency' :
                      'minor_deficiency';
      
      simulatedDeficiencies.push({
        deficiencyId: `DEF-${control.controlId}-${Date.now()}`,
        description: `Control testing identified exception in ${control.controlActivity}`,
        severity,
        rootCause: 'Control not performed as designed',
        potentialImpact: severity === 'material_weakness' ? 
          'Could result in material misstatement in financial statements' :
          'Could result in errors that may not be detected timely',
        managementResponse: 'Management will implement corrective action plan',
        remediationDeadline: new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toISOString().split('T')[0],
        status: 'open'
      });
    }
    
    const conclusion = simulatedDeficiencies.some(d => d.severity === 'material_weakness') ? 'ineffective' :
                      simulatedDeficiencies.some(d => d.severity === 'significant_deficiency') ? 'requires_remediation' :
                      'effective';
    
    return {
      testId: `TEST-${control.controlId}-${Date.now()}`,
      controlId: control.controlId,
      testDate,
      tester: testParameters.tester,
      testProcedure: procedures.join('\n'),
      sampleSize,
      populationSize: testParameters.populationSize,
      deficiencies: simulatedDeficiencies,
      conclusion,
      managementResponse: conclusion !== 'effective' ? 
        'Management acknowledges the testing results and will implement corrective actions' : undefined,
      remediationPlan: conclusion !== 'effective' ?
        'Detailed remediation plan to be developed within 15 days' : undefined,
      retestDate: conclusion !== 'effective' ?
        new Date(Date.now() + 60 * 24 * 60 * 60 * 1000).toISOString().split('T')[0] : undefined
    };
  }

  assessSOXCompliance(controls: SOXControl[], tests: ControlTest[]): SOXComplianceResult {
    const totalControls = controls.length;
    const testedControls = tests.length;
    const effectiveControls = tests.filter(t => t.conclusion === 'effective').length;
    const deficientControls = tests.filter(t => t.conclusion !== 'effective').length;
    
    // Aggregate by process
    const controlsByProcess: { [process: string]: number } = {};
    controls.forEach(control => {
      controlsByProcess[control.process] = (controlsByProcess[control.process] || 0) + 1;
    });
    
    // Aggregate by risk
    const controlsByRisk: { [risk: string]: number } = {};
    controls.forEach(control => {
      controlsByRisk[control.riskRating] = (controlsByRisk[control.riskRating] || 0) + 1;
    });
    
    // Count deficiencies by severity
    const allDeficiencies = tests.flatMap(t => t.deficiencies);
    const deficiencySummary = {
      materialWeaknesses: allDeficiencies.filter(d => d.severity === 'material_weakness').length,
      significantDeficiencies: allDeficiencies.filter(d => d.severity === 'significant_deficiency').length,
      minorDeficiencies: allDeficiencies.filter(d => d.severity === 'minor_deficiency').length
    };
    
    // Overall assessment
    let overallAssessment: 'effective' | 'ineffective' | 'requires_improvement';
    if (deficiencySummary.materialWeaknesses > 0) {
      overallAssessment = 'ineffective';
    } else if (deficiencySummary.significantDeficiencies > 0 || deficientControls / testedControls > 0.1) {
      overallAssessment = 'requires_improvement';
    } else {
      overallAssessment = 'effective';
    }
    
    // Generate executive summary
    const executiveSummary: string[] = [];
    executiveSummary.push(`SOX compliance assessment completed for ${totalControls} controls across ${Object.keys(controlsByProcess).length} processes`);
    executiveSummary.push(`${testedControls} controls tested (${((testedControls / totalControls) * 100).toFixed(1)}% coverage)`);
    executiveSummary.push(`${effectiveControls} controls operating effectively (${((effectiveControls / testedControls) * 100).toFixed(1)}% effectiveness rate)`);
    
    if (deficiencySummary.materialWeaknesses > 0) {
      executiveSummary.push(`⚠️ CRITICAL: ${deficiencySummary.materialWeaknesses} material weakness(es) identified requiring immediate remediation`);
    }
    
    if (deficiencySummary.significantDeficiencies > 0) {
      executiveSummary.push(`${deficiencySummary.significantDeficiencies} significant deficiency(ies) identified requiring management attention`);
    }
    
    // Generate recommendations
    const recommendations: string[] = [];
    
    if (overallAssessment === 'ineffective') {
      recommendations.push('Immediate remediation required for material weaknesses before financial statement certification');
      recommendations.push('Enhanced monitoring and testing required for all high-risk controls');
      recommendations.push('Consider engaging external SOX specialists for remediation support');
    }
    
    if (overallAssessment === 'requires_improvement') {
      recommendations.push('Develop comprehensive corrective action plans for significant deficiencies');
      recommendations.push('Increase frequency of control monitoring and testing');
      recommendations.push('Consider process improvements to address root causes');
    }
    
    if (testedControls / totalControls < 0.8) {
      recommendations.push('Increase control testing coverage to ensure comprehensive assessment');
    }
    
    const highRiskUntested = controls.filter(c => 
      c.riskRating === 'high' && !tests.some(t => t.controlId === c.controlId)
    ).length;
    
    if (highRiskUntested > 0) {
      recommendations.push(`Priority testing needed for ${highRiskUntested} untested high-risk controls`);
    }
    
    recommendations.push('Continue ongoing monitoring and periodic reassessment of control effectiveness');
    
    return {
      compliancePeriod: new Date().getFullYear().toString(),
      totalControls,
      testedControls,
      effectiveControls,
      deficientControls,
      controlsByProcess,
      controlsByRisk,
      deficiencySummary,
      overallAssessment,
      executiveSummary,
      recommendations
    };
  }

  generateSOXComplianceWorksheet(
    controls: SOXControl[], 
    tests: ControlTest[], 
    assessment: SOXComplianceResult
  ): Array<Array<CellValue | string | number>> {
    const worksheet: Array<Array<CellValue | string | number>> = [];
    
    // Header
    worksheet.push(['SOX CONTROLS COMPLIANCE ASSESSMENT', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Sarbanes-Oxley Act Section 404 - Internal Control Assessment', '', '', '', '', '', '', '', '', '']);
    worksheet.push([`Assessment Period: ${assessment.compliancePeriod}`, '', '', '', '', '', '', '', '', '']);
    worksheet.push([`Generated: ${new Date().toLocaleDateString()}`, '', '', '', '', '', '', '', '', '']);
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Executive Summary
    worksheet.push(['🎯 EXECUTIVE SUMMARY', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Metric', 'Value', 'Status', 'Benchmark', 'Assessment', '', '', '', '', '']);
    
    worksheet.push(['Overall SOX Compliance', 
                   assessment.overallAssessment.toUpperCase(),
                   assessment.overallAssessment === 'effective' ? '✅ Pass' : '❌ Fail',
                   'Effective required',
                   assessment.overallAssessment === 'effective' ? 'Meets requirements' : 'Requires action'
                  ]);
    
    worksheet.push(['Total Controls', 
                   assessment.totalControls,
                   'Complete',
                   'Full population',
                   'Control population documented'
                  ]);
    
    worksheet.push(['Controls Tested', 
                   `${assessment.testedControls}/${assessment.totalControls}`,
                   assessment.testedControls / assessment.totalControls >= 0.8 ? '✅ Adequate' : '⚠️ Low',
                   '≥80% coverage',
                   ((assessment.testedControls / assessment.totalControls) * 100).toFixed(1) + '% coverage'
                  ]);
    
    worksheet.push(['Effective Controls', 
                   `${assessment.effectiveControls}/${assessment.testedControls}`,
                   assessment.effectiveControls / assessment.testedControls >= 0.95 ? '✅ Strong' : '⚠️ Weak',
                   '≥95% effectiveness',
                   ((assessment.effectiveControls / assessment.testedControls) * 100).toFixed(1) + '% effective'
                  ]);
    
    worksheet.push(['Material Weaknesses', 
                   assessment.deficiencySummary.materialWeaknesses,
                   assessment.deficiencySummary.materialWeaknesses === 0 ? '✅ None' : '❌ Found',
                   'Zero required',
                   assessment.deficiencySummary.materialWeaknesses === 0 ? 'Compliant' : 'Non-compliant'
                  ]);
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Control Population Analysis
    worksheet.push(['📊 CONTROL POPULATION ANALYSIS', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Process', 'Control Count', 'Tested', 'Effective', 'Deficient', 'Effectiveness %', '', '', '', '']);
    
    for (const [process, count] of Object.entries(assessment.controlsByProcess)) {
      const processControls = controls.filter(c => c.process === process);
      const processTests = tests.filter(t => processControls.some(c => c.controlId === t.controlId));
      const processEffective = processTests.filter(t => t.conclusion === 'effective').length;
      
      worksheet.push([
        process,
        count,
        processTests.length,
        processEffective,
        processTests.length - processEffective,
        processTests.length > 0 ? ((processEffective / processTests.length) * 100).toFixed(1) + '%' : 'N/A',
        '', '', '', ''
      ]);
    }
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Risk Analysis
    worksheet.push(['⚠️ RISK ANALYSIS', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Risk Rating', 'Control Count', 'Tested', 'Effective', 'Test Priority', 'Status', '', '', '', '']);
    
    for (const [risk, count] of Object.entries(assessment.controlsByRisk)) {
      const riskControls = controls.filter(c => c.riskRating === risk);
      const riskTests = tests.filter(t => riskControls.some(c => c.controlId === t.controlId));
      const riskEffective = riskTests.filter(t => t.conclusion === 'effective').length;
      
      const priority = risk === 'high' ? 'Critical' : risk === 'medium' ? 'Important' : 'Standard';
      const status = risk === 'high' && riskTests.length / count < 0.9 ? '⚠️ Insufficient Testing' :
                    riskTests.length > 0 && riskEffective / riskTests.length >= 0.95 ? '✅ Effective' :
                    '⚠️ Requires Attention';
      
      worksheet.push([
        risk.toUpperCase(),
        count,
        riskTests.length,
        riskEffective,
        priority,
        status,
        '', '', '', ''
      ]);
    }
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Detailed Control Testing Results
    worksheet.push(['🔍 DETAILED CONTROL TESTING RESULTS', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Control ID', 'Control Name', 'Process', 'Risk', 'Test Date', 'Tester', 'Result', 'Deficiencies', '', '']);
    
    for (const test of tests) {
      const control = controls.find(c => c.controlId === test.controlId);
      
      worksheet.push([
        test.controlId,
        control?.controlName || 'Unknown',
        control?.process || 'Unknown',
        control?.riskRating.toUpperCase() || 'Unknown',
        test.testDate,
        test.tester,
        test.conclusion.replace('_', ' ').toUpperCase(),
        test.deficiencies.length.toString(),
        '', ''
      ]);
    }
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Deficiency Summary
    worksheet.push(['❌ CONTROL DEFICIENCIES SUMMARY', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Deficiency ID', 'Control ID', 'Severity', 'Description', 'Root Cause', 'Remediation Deadline', 'Status', '', '', '']);
    
    const allDeficiencies = tests.flatMap(test => 
      test.deficiencies.map(def => ({ ...def, testId: test.testId, controlId: test.controlId }))
    );
    
    for (const deficiency of allDeficiencies) {
      worksheet.push([
        deficiency.deficiencyId,
        deficiency.controlId,
        deficiency.severity.replace('_', ' ').toUpperCase(),
        deficiency.description.substring(0, 50) + (deficiency.description.length > 50 ? '...' : ''),
        deficiency.rootCause.substring(0, 40) + (deficiency.rootCause.length > 40 ? '...' : ''),
        deficiency.remediationDeadline,
        deficiency.status.toUpperCase(),
        '', '', ''
      ]);
    }
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Management Representation
    worksheet.push(['📝 MANAGEMENT ASSESSMENT & REPRESENTATIONS', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Area', 'Management Assessment', 'Supporting Evidence', 'Conclusion', '', '', '', '', '', '']);
    
    worksheet.push(['Control Environment', 
                   'Management maintains effective control environment',
                   'Tone at top, organizational structure, HR policies',
                   'Adequate based on testing'
                  ]);
    
    worksheet.push(['Risk Assessment', 
                   'Company identifies and assesses financial reporting risks',
                   'Risk identification process, fraud risk assessment',
                   'Processes in place and operating'
                  ]);
    
    worksheet.push(['Control Activities', 
                   'Control activities are designed and operating effectively',
                   `${assessment.effectiveControls}/${assessment.testedControls} controls effective`,
                   assessment.overallAssessment === 'effective' ? 'Effective' : 'Needs improvement'
                  ]);
    
    worksheet.push(['Information & Communication', 
                   'Relevant financial reporting information is communicated',
                   'Financial reporting systems, policies, procedures',
                   'Systems support effective reporting'
                  ]);
    
    worksheet.push(['Monitoring Activities', 
                   'Management monitors internal control effectiveness',
                   'Ongoing monitoring, separate evaluations',
                   'Monitoring processes established'
                  ]);
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Recommendations & Action Plan
    worksheet.push(['💡 RECOMMENDATIONS & ACTION PLAN', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['Priority', 'Recommendation', 'Responsible Party', 'Target Date', 'Status', '', '', '', '', '']);
    
    let priority = 1;
    for (const recommendation of assessment.recommendations) {
      const priorityLevel = recommendation.includes('Immediate') || recommendation.includes('material weakness') ? 'HIGH' :
                           recommendation.includes('significant defic') ? 'MEDIUM' : 'LOW';
      
      const targetDate = priorityLevel === 'HIGH' ? 
        new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toISOString().split('T')[0] :
        new Date(Date.now() + 90 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];
      
      worksheet.push([
        priorityLevel,
        recommendation,
        'Management/Internal Audit',
        targetDate,
        'OPEN',
        '', '', '', '', ''
      ]);
      priority++;
    }
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Regulatory Requirements
    worksheet.push(['📋 REGULATORY COMPLIANCE CHECKLIST', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['SOX Requirement', 'Status', 'Evidence', 'Notes', '', '', '', '', '', '']);
    
    worksheet.push(['Section 302 - CEO/CFO Certification', 
                   assessment.deficiencySummary.materialWeaknesses === 0 ? '✅ Ready' : '❌ Not Ready',
                   'Management assessment completed',
                   assessment.deficiencySummary.materialWeaknesses === 0 ? 
                     'No material weaknesses identified' : 
                     `${assessment.deficiencySummary.materialWeaknesses} material weakness(es) must be remediated`
                  ]);
    
    worksheet.push(['Section 404(a) - Management Assessment', 
                   '✅ Complete',
                   'Internal control assessment documented',
                   'Management assessment of ICFR effectiveness completed'
                  ]);
    
    worksheet.push(['Section 404(b) - Auditor Attestation', 
                   assessment.overallAssessment === 'effective' ? '✅ Ready' : '⚠️ At Risk',
                   'Control testing and deficiency analysis',
                   'Auditor testing will validate management assessment'
                  ]);
    
    worksheet.push(['Disclosure Controls', 
                   '✅ Assessed',
                   'Control testing results',
                   'Disclosure controls included in assessment scope'
                  ]);
    
    worksheet.push(['', '', '', '', '', '', '', '', '', '']);
    
    // Professional Standards References
    worksheet.push(['📚 PROFESSIONAL STANDARDS REFERENCES', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- Sarbanes-Oxley Act of 2002, Sections 302 and 404', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- SEC Rules 13a-15 and 15d-15 (Disclosure Controls)', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- PCAOB AS 2201 - Audit of Internal Control', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- COSO Internal Control - Integrated Framework', '', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- SEC Guidance on Management Assessment of ICFR', '', '', '', '', '', '', '', '', '', '']);
    worksheet.push(['- PCAOB Staff Questions and Answers on AS 2201', '', '', '', '', '', '', '', '', '', '']);
    
    return worksheet;
  }
}