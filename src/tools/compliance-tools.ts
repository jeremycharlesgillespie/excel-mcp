import { Tool, ToolResult } from '../types/index.js';
import { ExcelManager } from '../excel/excel-manager.js';
import { ASC606Engine, ContractInput, RevenueRecognitionResult } from '../compliance/asc-606.js';
import { SOXControlsEngine, SOXControl, ControlTest } from '../compliance/sox-controls.js';

const excelManager = new ExcelManager();
const asc606Engine = new ASC606Engine();
const soxEngine = new SOXControlsEngine();

export const complianceTools: Tool[] = [
  {
    name: "compliance_asc606_revenue_recognition",
    description: "Process revenue recognition under ASC 606 with full compliance documentation and audit trail",
    inputSchema: {
      type: "object",
      properties: {
        contracts: {
          type: "array",
          items: {
            type: "object",
            properties: {
              contractId: { type: "string", description: "Unique contract identifier" },
              customerId: { type: "string", description: "Customer identifier" },
              contractDate: { type: "string", description: "Contract date (YYYY-MM-DD)" },
              contractValue: { type: "number", description: "Total contract value in currency" },
              currency: { type: "string", default: "USD", description: "Contract currency code" },
              contractTerm: { type: "number", description: "Contract term in months" },
              performanceObligations: {
                type: "array",
                items: {
                  type: "object",
                  properties: {
                    id: { type: "string", description: "Performance obligation ID" },
                    description: { type: "string", description: "Description of goods/services" },
                    standAloneSellingPrice: { type: "number", description: "Standalone selling price" },
                    allocatedTransactionPrice: { type: "number", description: "Allocated transaction price" },
                    recognitionMethod: { 
                      type: "string", 
                      enum: ["point_in_time", "over_time"],
                      description: "Revenue recognition method"
                    },
                    percentComplete: { type: "number", description: "Percent complete (0-1) for over_time recognition" },
                    completionDate: { type: "string", description: "Expected completion date (YYYY-MM-DD)" },
                    deliveryDate: { type: "string", description: "Delivery date for point_in_time (YYYY-MM-DD)" },
                    satisfactionCriteria: { type: "string", description: "Criteria for satisfaction of obligation" }
                  },
                  required: ["id", "description", "standAloneSellingPrice", "allocatedTransactionPrice", "recognitionMethod", "satisfactionCriteria"]
                }
              },
              paymentTerms: {
                type: "array",
                items: {
                  type: "object",
                  properties: {
                    dueDate: { type: "string", description: "Payment due date (YYYY-MM-DD)" },
                    amount: { type: "number", description: "Payment amount" },
                    description: { type: "string", description: "Payment description" },
                    received: { type: "boolean", default: false, description: "Whether payment received" },
                    receivedDate: { type: "string", description: "Date payment received (YYYY-MM-DD)" }
                  },
                  required: ["dueDate", "amount", "description", "received"]
                }
              }
            },
            required: ["contractId", "customerId", "contractDate", "contractValue", "contractTerm", "performanceObligations", "paymentTerms"]
          }
        },
        asOfDate: { type: "string", default: "today", description: "Recognition as-of date (YYYY-MM-DD)" },
        worksheetName: { type: "string", default: "ASC 606 Analysis" }
      },
      required: ["contracts"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const asOfDate = args.asOfDate && args.asOfDate !== "today" 
          ? new Date(args.asOfDate) 
          : new Date();
        
        const contracts: ContractInput[] = args.contracts;
        const results: RevenueRecognitionResult[] = [];
        
        // Process each contract
        for (const contract of contracts) {
          const result = asc606Engine.processRevenueRecognition(contract, asOfDate);
          results.push(result);
        }
        
        // Generate comprehensive worksheet
        const worksheetData = asc606Engine.generateASC606Worksheet(results, asOfDate);
        const worksheetName = args.worksheetName || "ASC 606 Analysis";
        
        await excelManager.addWorksheet(worksheetName);
        await excelManager.writeWorksheet(worksheetName, worksheetData);
        await excelManager.autoFitColumnWidths(worksheetName, { minWidth: 100, maxWidth: 300 });
        
        // Calculate summary statistics
        const totalContractValue = results.reduce((sum, r) => sum + r.totalContractValue, 0);
        const totalRevenue = results.reduce((sum, r) => sum + r.currentPeriodRevenue, 0);
        const totalRemaining = results.reduce((sum, r) => sum + r.remainingRevenue, 0);
        const recognitionRate = (totalRevenue / totalContractValue) * 100;
        
        // Count compliance issues
        const contractsWithIssues = results.filter(r => 
          r.complianceNotes.some(note => 
            note.toLowerCase().includes('issue') || 
            note.toLowerCase().includes('error') ||
            note.toLowerCase().includes('mismatch')
          )
        ).length;
        
        return {
          success: true,
          message: `ASC 606 revenue recognition analysis completed for ${contracts.length} contracts`,
          data: {
            worksheetName,
            summary: {
              contractsAnalyzed: contracts.length,
              totalContractValue: (totalContractValue / 1000000).toFixed(2) + "M",
              revenueRecognized: (totalRevenue / 1000000).toFixed(2) + "M",
              remainingRevenue: (totalRemaining / 1000000).toFixed(2) + "M",
              recognitionRate: recognitionRate.toFixed(1) + "%",
              complianceStatus: contractsWithIssues === 0 ? "✅ Fully Compliant" : `⚠️ ${contractsWithIssues} contracts with issues`
            },
            compliance: {
              asc606Compliant: contractsWithIssues === 0,
              totalAuditEntries: results.reduce((sum, r) => sum + r.auditTrail.length, 0),
              avgPerformanceObligations: (results.reduce((sum, r) => sum + Object.keys(r.performanceObligationStatus).length, 0) / results.length).toFixed(1)
            }
          }
        };
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "compliance_sox_controls_test",
    description: "Test SOX controls for compliance with detailed testing procedures and deficiency tracking",
    inputSchema: {
      type: "object",
      properties: {
        controls: {
          type: "array",
          items: {
            type: "object",
            properties: {
              controlId: { type: "string", description: "Unique control identifier" },
              controlName: { type: "string", description: "Control name/title" },
              controlObjective: { type: "string", description: "Control objective statement" },
              controlType: { 
                type: "string", 
                enum: ["preventive", "detective", "corrective"],
                description: "Type of control"
              },
              frequency: { 
                type: "string", 
                enum: ["daily", "weekly", "monthly", "quarterly", "annually", "event_driven"],
                description: "Control frequency"
              },
              owner: { type: "string", description: "Control owner/responsible party" },
              riskRating: { 
                type: "string", 
                enum: ["high", "medium", "low"],
                description: "Risk rating of the control"
              },
              process: { type: "string", description: "Business process (e.g., 'Revenue', 'Procurement')" },
              controlActivity: { type: "string", description: "Description of control activity" },
              evidence: { 
                type: "array", 
                items: { type: "string" },
                description: "Types of evidence for control operation"
              },
              lastTestDate: { type: "string", description: "Last test date (YYYY-MM-DD)" },
              nextTestDate: { type: "string", description: "Next scheduled test date (YYYY-MM-DD)" },
              status: { 
                type: "string", 
                enum: ["effective", "ineffective", "not_tested", "requires_remediation"],
                description: "Current control status"
              }
            },
            required: ["controlId", "controlName", "controlObjective", "controlType", "frequency", "owner", "riskRating", "process", "controlActivity", "evidence"]
          }
        },
        testParameters: {
          type: "array",
          items: {
            type: "object",
            properties: {
              controlId: { type: "string", description: "Control ID to test" },
              tester: { type: "string", description: "Person performing the test" },
              testDate: { type: "string", description: "Test date (YYYY-MM-DD)" },
              populationSize: { type: "number", description: "Total population size for sampling" },
              sampleSize: { type: "number", description: "Custom sample size (optional)" },
              customProcedures: { 
                type: "array", 
                items: { type: "string" },
                description: "Custom test procedures (optional)"
              }
            },
            required: ["controlId", "tester", "populationSize"]
          }
        },
        worksheetName: { type: "string", default: "SOX Controls Testing" }
      },
      required: ["controls", "testParameters"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const controls: SOXControl[] = args.controls;
        const testParameters = args.testParameters;
        
        const tests: ControlTest[] = [];
        
        // Perform testing for each specified control
        for (const testParam of testParameters) {
          const control = controls.find(c => c.controlId === testParam.controlId);
          if (!control) {
            throw new Error(`Control ${testParam.controlId} not found in control population`);
          }
          
          const test = soxEngine.testSOXControl(control, {
            tester: testParam.tester,
            testDate: testParam.testDate,
            sampleSize: testParam.sampleSize,
            populationSize: testParam.populationSize,
            customProcedures: testParam.customProcedures
          });
          
          tests.push(test);
        }
        
        // Perform overall compliance assessment
        const assessment = soxEngine.assessSOXCompliance(controls, tests);
        
        // Generate comprehensive compliance worksheet
        const worksheetData = soxEngine.generateSOXComplianceWorksheet(controls, tests, assessment);
        const worksheetName = args.worksheetName || "SOX Controls Testing";
        
        await excelManager.addWorksheet(worksheetName);
        await excelManager.writeWorksheet(worksheetName, worksheetData);
        await excelManager.autoFitColumnWidths(worksheetName, { minWidth: 90, maxWidth: 250 });
        
        // Count critical findings
        const allDeficiencies = tests.flatMap(t => t.deficiencies);
        const materialWeaknesses = allDeficiencies.filter(d => d.severity === 'material_weakness').length;
        const significantDeficiencies = allDeficiencies.filter(d => d.severity === 'significant_deficiency').length;
        
        return {
          success: true,
          message: `SOX controls testing completed for ${tests.length} controls with ${assessment.overallAssessment} assessment`,
          data: {
            worksheetName,
            testingSummary: {
              controlsTested: tests.length,
              totalControls: controls.length,
              testCoverage: ((tests.length / controls.length) * 100).toFixed(1) + "%",
              effectiveControls: assessment.effectiveControls,
              effectivenessRate: ((assessment.effectiveControls / tests.length) * 100).toFixed(1) + "%"
            },
            complianceAssessment: {
              overallStatus: assessment.overallAssessment.toUpperCase(),
              materialWeaknesses,
              significantDeficiencies,
              readyForCertification: materialWeaknesses === 0,
              auditReadiness: assessment.overallAssessment === 'effective' ? "✅ Ready" : "⚠️ Remediation Required"
            },
            riskProfile: {
              highRiskControls: controls.filter(c => c.riskRating === 'high').length,
              highRiskTested: tests.filter(t => controls.find(c => c.controlId === t.controlId)?.riskRating === 'high').length,
              criticalGaps: materialWeaknesses + significantDeficiencies
            }
          }
        };
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "compliance_audit_preparation",
    description: "Generate comprehensive audit preparation package with all compliance documentation",
    inputSchema: {
      type: "object",
      properties: {
        auditType: { 
          type: "string", 
          enum: ["financial_audit", "sox_audit", "compliance_audit", "operational_audit"],
          description: "Type of audit being prepared for"
        },
        auditPeriod: { type: "string", description: "Audit period (e.g., 'FY 2024', 'Q4 2024')" },
        auditorFirm: { type: "string", description: "External auditor firm name" },
        keyAuditors: { 
          type: "array", 
          items: { type: "string" },
          description: "Names of key audit team members"
        },
        auditAreas: {
          type: "array",
          items: { type: "string" },
          description: "Key audit areas/processes to focus on"
        },
        controlsToDocument: { 
          type: "array", 
          items: { type: "string" },
          description: "Control IDs that require documentation"
        },
        complianceFrameworks: {
          type: "array",
          items: { 
            type: "string",
            enum: ["SOX", "ASC606", "ASC842", "GAAP", "IFRS", "COSO"]
          },
          description: "Compliance frameworks in scope"
        },
        worksheetName: { type: "string", default: "Audit Preparation" }
      },
      required: ["auditType", "auditPeriod", "auditAreas", "complianceFrameworks"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const worksheet: Array<Array<any>> = [];
        
        // Header
        worksheet.push(['AUDIT PREPARATION PACKAGE', '', '', '', '', '', '', '', '', '']);
        worksheet.push([`${args.auditType.replace('_', ' ').toUpperCase()} - ${args.auditPeriod}`, '', '', '', '', '', '', '', '', '']);
        worksheet.push([`Auditor: ${args.auditorFirm || 'TBD'}`, '', '', '', '', '', '', '', '', '']);
        worksheet.push([`Prepared: ${new Date().toLocaleDateString()}`, '', '', '', '', '', '', '', '', '']);
        worksheet.push(['', '', '', '', '', '', '', '', '', '']);
        
        // Audit Scope Summary
        worksheet.push(['🎯 AUDIT SCOPE & OBJECTIVES', '', '', '', '', '', '', '', '', '']);
        worksheet.push(['Area', 'In Scope', 'Key Risks', 'Control Framework', 'Documentation Status', '', '', '', '', '']);
        
        for (const area of args.auditAreas) {
          const framework = args.complianceFrameworks.includes('SOX') ? 'SOX/COSO' :
                           args.complianceFrameworks.includes('ASC606') ? 'ASC 606' :
                           args.complianceFrameworks.join(', ');
          
          worksheet.push([
            area,
            '✅ Yes',
            area.toLowerCase().includes('revenue') ? 'Revenue recognition, cut-off' :
            area.toLowerCase().includes('expense') ? 'Completeness, accuracy' :
            area.toLowerCase().includes('cash') ? 'Existence, valuation' :
            'Accuracy, completeness, valuation',
            framework,
            '📋 Ready for Review',
            '', '', '', '', ''
          ]);
        }
        
        worksheet.push(['', '', '', '', '', '', '', '', '', '']);
        
        // Pre-Audit Checklist
        worksheet.push(['✅ PRE-AUDIT CHECKLIST', '', '', '', '', '', '', '', '', '']);
        worksheet.push(['Category', 'Requirement', 'Status', 'Due Date', 'Responsible Party', 'Notes', '', '', '', '']);
        
        const checklistItems = [
          ['Financial Statements', 'Draft financial statements prepared', '✅ Complete', '2 weeks before audit', 'Accounting Team', 'Ready for review'],
          ['Trial Balance', 'Final trial balance with all adjustments', '✅ Complete', '2 weeks before audit', 'Accounting Team', 'Reconciled and reviewed'],
          ['Account Reconciliations', 'All balance sheet reconciliations current', '⚠️ In Progress', '1 week before audit', 'Accounting Team', 'Monthly recs complete'],
          ['Supporting Documentation', 'Organize supporting documentation', '📋 Planned', '1 week before audit', 'All Teams', 'Document request pending'],
          ['Control Testing', 'Complete internal control testing', args.complianceFrameworks.includes('SOX') ? '✅ Complete' : 'N/A', 'Before fieldwork', 'Internal Audit', 'SOX testing completed'],
          ['Management Letter', 'Prepare management representation letter', '📝 Draft', 'End of fieldwork', 'Management', 'Template prepared'],
          ['Audit Committee', 'Schedule audit committee meetings', '📅 Scheduled', 'Per audit timeline', 'Board Secretary', 'Meetings confirmed'],
          ['PBC List', 'Prepare prepared-by-client list', '📋 In Progress', 'Before fieldwork', 'Audit Coordinator', 'Items being compiled']
        ];
        
        for (const [category, requirement, status, dueDate, responsible, notes] of checklistItems) {
          worksheet.push([category, requirement, status, dueDate, responsible, notes, '', '', '', '']);
        }
        
        worksheet.push(['', '', '', '', '', '', '', '', '', '']);
        
        // Document Index
        worksheet.push(['📁 AUDIT DOCUMENTATION INDEX', '', '', '', '', '', '', '', '', '']);
        worksheet.push(['Document Category', 'Document Type', 'Location/Reference', 'Last Updated', 'Prepared By', 'Status', '', '', '', '']);
        
        const documentCategories = [
          ['Financial Statements', 'Draft F/S Package', 'Accounting/Financial Statements/', new Date().toLocaleDateString(), 'Controller', '✅ Ready'],
          ['Account Reconciliations', 'Balance Sheet Reconciliations', 'Accounting/Reconciliations/', new Date().toLocaleDateString(), 'Staff Accountant', '✅ Ready'],
          ['General Ledger', 'Trial Balance & G/L Detail', 'Accounting/General Ledger/', new Date().toLocaleDateString(), 'Accounting Manager', '✅ Ready'],
          ['Cash', 'Bank Reconciliations & Statements', 'Treasury/Cash Management/', new Date().toLocaleDateString(), 'Treasury', '✅ Ready'],
          ['Accounts Receivable', 'A/R Aging & Confirmations', 'Accounting/AR/', new Date().toLocaleDateString(), 'AR Manager', '✅ Ready'],
          ['Accounts Payable', 'A/P Aging & Cut-off Testing', 'Accounting/AP/', new Date().toLocaleDateString(), 'AP Manager', '📋 Pending'],
          ['Inventory', 'Inventory Counts & Valuations', 'Operations/Inventory/', new Date().toLocaleDateString(), 'Operations Manager', '📋 Pending'],
          ['Fixed Assets', 'Asset Register & Depreciation', 'Accounting/Fixed Assets/', new Date().toLocaleDateString(), 'Fixed Asset Accountant', '✅ Ready'],
          ['Debt & Financing', 'Loan Agreements & Schedules', 'Treasury/Debt/', new Date().toLocaleDateString(), 'CFO', '✅ Ready'],
          ['Equity', 'Stock Records & Transactions', 'Legal/Equity/', new Date().toLocaleDateString(), 'Legal Counsel', '✅ Ready'],
          ['Revenue', 'Revenue Recognition Documentation', 'Accounting/Revenue/', new Date().toLocaleDateString(), 'Revenue Manager', args.complianceFrameworks.includes('ASC606') ? '✅ Ready' : '📋 TBD'],
          ['Internal Controls', 'SOX Documentation & Testing', 'Internal Audit/SOX/', new Date().toLocaleDateString(), 'Internal Audit', args.complianceFrameworks.includes('SOX') ? '✅ Ready' : 'N/A']
        ];
        
        for (const [category, docType, location, updated, preparedBy, status] of documentCategories) {
          worksheet.push([category, docType, location, updated, preparedBy, status, '', '', '', '']);
        }
        
        worksheet.push(['', '', '', '', '', '', '', '', '', '']);
        
        // Key Contacts
        worksheet.push(['👥 KEY AUDIT CONTACTS', '', '', '', '', '', '', '', '', '']);
        worksheet.push(['Role', 'Name', 'Title', 'Email', 'Phone', 'Backup', '', '', '', '']);
        
        const contacts = [
          ['Audit Committee Chair', 'TBD', 'Board Member', 'chair@company.com', '(555) 001-0001', 'Vice Chair'],
          ['CFO', 'TBD', 'Chief Financial Officer', 'cfo@company.com', '(555) 001-0002', 'Controller'],
          ['Controller', 'TBD', 'Corporate Controller', 'controller@company.com', '(555) 001-0003', 'Assistant Controller'],
          ['Internal Audit Director', 'TBD', 'Director, Internal Audit', 'ia_director@company.com', '(555) 001-0004', 'Senior Manager'],
          ['Legal Counsel', 'TBD', 'General Counsel', 'legal@company.com', '(555) 001-0005', 'Assistant Counsel'],
          ['IT Director', 'TBD', 'Director, Information Technology', 'it_director@company.com', '(555) 001-0006', 'IT Manager'],
          ['HR Director', 'TBD', 'Director, Human Resources', 'hr_director@company.com', '(555) 001-0007', 'HR Manager']
        ];
        
        for (const [role, name, title, email, phone, backup] of contacts) {
          worksheet.push([role, name, title, email, phone, backup, '', '', '', '']);
        }
        
        worksheet.push(['', '', '', '', '', '', '', '', '', '']);
        
        // Audit Timeline
        worksheet.push(['📅 AUDIT TIMELINE & MILESTONES', '', '', '', '', '', '', '', '', '']);
        worksheet.push(['Phase', 'Activity', 'Start Date', 'End Date', 'Responsible Party', 'Dependencies', '', '', '', '']);
        
        const today = new Date();
        const planningStart = new Date(today.getTime() + 7 * 24 * 60 * 60 * 1000);
        const fieldworkStart = new Date(today.getTime() + 21 * 24 * 60 * 60 * 1000);
        const completionDate = new Date(today.getTime() + 60 * 24 * 60 * 60 * 1000);
        
        const timeline = [
          ['Planning', 'Audit planning and risk assessment', planningStart.toLocaleDateString(), new Date(planningStart.getTime() + 7 * 24 * 60 * 60 * 1000).toLocaleDateString(), 'External Auditors', 'PBC list completion'],
          ['Interim Testing', 'Internal controls testing', new Date(planningStart.getTime() + 7 * 24 * 60 * 60 * 1000).toLocaleDateString(), new Date(fieldworkStart.getTime() - 7 * 24 * 60 * 60 * 1000).toLocaleDateString(), 'External Auditors', 'Control documentation'],
          ['Fieldwork', 'Substantive testing and procedures', fieldworkStart.toLocaleDateString(), new Date(fieldworkStart.getTime() + 21 * 24 * 60 * 60 * 1000).toLocaleDateString(), 'External Auditors', 'Financial statements ready'],
          ['Review', 'Partner review and quality control', new Date(fieldworkStart.getTime() + 21 * 24 * 60 * 60 * 1000).toLocaleDateString(), new Date(completionDate.getTime() - 7 * 24 * 60 * 60 * 1000).toLocaleDateString(), 'Audit Partner', 'Fieldwork completion'],
          ['Completion', 'Final procedures and report issuance', new Date(completionDate.getTime() - 7 * 24 * 60 * 60 * 1000).toLocaleDateString(), completionDate.toLocaleDateString(), 'External Auditors', 'Management letter'],
          ['Communication', 'Audit committee and management communication', completionDate.toLocaleDateString(), new Date(completionDate.getTime() + 7 * 24 * 60 * 60 * 1000).toLocaleDateString(), 'Management', 'Audit completion']
        ];
        
        for (const [phase, activity, startDate, endDate, responsible, dependencies] of timeline) {
          worksheet.push([phase, activity, startDate, endDate, responsible, dependencies, '', '', '', '']);
        }
        
        worksheet.push(['', '', '', '', '', '', '', '', '', '']);
        
        // Compliance Framework Summary
        worksheet.push(['📋 COMPLIANCE FRAMEWORKS IN SCOPE', '', '', '', '', '', '', '', '', '']);
        worksheet.push(['Framework', 'Applicable Standards', 'Key Requirements', 'Documentation Status', 'Readiness Level', '', '', '', '', '']);
        
        for (const framework of args.complianceFrameworks) {
          let standards, requirements, status;
          
          switch (framework) {
            case 'SOX':
              standards = 'Sections 302, 404, 906';
              requirements = 'ICFR assessment, CEO/CFO certification';
              status = '✅ Controls tested and documented';
              break;
            case 'ASC606':
              standards = 'Revenue from Contracts with Customers';
              requirements = '5-step model compliance, disclosure requirements';
              status = '✅ Implementation documented';
              break;
            case 'ASC842':
              standards = 'Leases';
              requirements = 'Lease accounting, ROU assets/liabilities';
              status = '📋 Assessment in progress';
              break;
            case 'GAAP':
              standards = 'US Generally Accepted Accounting Principles';
              requirements = 'Financial statement preparation and presentation';
              status = '✅ GAAP-compliant processes';
              break;
            case 'IFRS':
              standards = 'International Financial Reporting Standards';
              requirements = 'Global financial reporting consistency';
              status = '📋 Convergence analysis needed';
              break;
            case 'COSO':
              standards = 'Internal Control Framework';
              requirements = '5 components of internal control';
              status = '✅ Framework implemented';
              break;
            default:
              standards = 'Various applicable standards';
              requirements = 'Framework-specific requirements';
              status = '📋 Assessment needed';
          }
          
          worksheet.push([framework, standards, requirements, status, status.includes('✅') ? '🟢 Ready' : '🟡 In Progress', '', '', '', '', '']);
        }
        
        worksheet.push(['', '', '', '', '', '', '', '', '', '']);
        
        // Risk Assessment
        worksheet.push(['⚠️ AUDIT RISK ASSESSMENT', '', '', '', '', '', '', '', '', '']);
        worksheet.push(['Risk Area', 'Risk Level', 'Description', 'Mitigation Strategy', 'Owner', '', '', '', '', '']);
        
        const riskAreas = [
          ['Revenue Recognition', 'High', 'Complex contracts and ASC 606 compliance', 'Detailed contract review and testing', 'Revenue Manager'],
          ['Management Override', 'Medium', 'Risk of management override of controls', 'Enhanced journal entry testing', 'Internal Audit'],
          ['IT General Controls', 'Medium', 'System access and change management', 'ITGC testing and documentation', 'IT Director'],
          ['Related Party Transactions', 'Low', 'Identification and disclosure of RPTs', 'Related party questionnaires', 'Legal Counsel'],
          ['Going Concern', 'Low', 'Entity\'s ability to continue as going concern', 'Cash flow analysis and projections', 'CFO'],
          ['Fraud Risk', 'Medium', 'Risk of material misstatement due to fraud', 'Fraud risk assessment procedures', 'Management']
        ];
        
        for (const [area, level, description, mitigation, owner] of riskAreas) {
          worksheet.push([area, level, description, mitigation, owner, '', '', '', '', '']);
        }
        
        worksheet.push(['', '', '', '', '', '', '', '', '', '']);
        
        // Final Checklist
        worksheet.push(['🎯 FINAL READINESS CHECKLIST', '', '', '', '', '', '', '', '', '']);
        worksheet.push(['Item', 'Complete', 'Notes', 'Sign-off', '', '', '', '', '', '']);
        
        const finalChecklist = [
          ['All account reconciliations current and reviewed', '✅', 'Monthly reconciliations complete', 'Controller'],
          ['Supporting documentation organized and accessible', '📋', 'Document request list prepared', 'Audit Coordinator'],
          ['Internal control testing completed (if applicable)', args.complianceFrameworks.includes('SOX') ? '✅' : 'N/A', 'SOX testing documented', 'Internal Audit'],
          ['Management representation letter prepared', '📝', 'Draft prepared, final pending', 'Management'],
          ['Audit committee briefed on audit plan', '📅', 'Meeting scheduled', 'Board Secretary'],
          ['Key personnel availability confirmed', '✅', 'Schedules coordinated', 'HR Director'],
          ['Workspace and technology prepared for auditors', '🏢', 'Conference room reserved, IT access ready', 'Facilities'],
          ['Prior year audit issues addressed', '✅', 'No open items from prior year', 'Controller']
        ];
        
        for (const [item, complete, notes, signoff] of finalChecklist) {
          worksheet.push([item, complete, notes, signoff, '', '', '', '', '', '']);
        }
        
        const worksheetName = args.worksheetName || "Audit Preparation";
        
        await excelManager.addWorksheet(worksheetName);
        await excelManager.writeWorksheet(worksheetName, worksheet);
        await excelManager.autoFitColumnWidths(worksheetName, { minWidth: 120, maxWidth: 300 });
        
        // Calculate readiness metrics
        const readyItems = worksheet.filter(row => typeof row[1] === 'string' && row[1].includes('✅')).length;
        const totalItems = worksheet.filter(row => typeof row[0] === 'string' && (row[1] === '✅' || row[1] === '📋' || row[1] === '⚠️')).length;
        const readinessPercentage = totalItems > 0 ? (readyItems / totalItems * 100) : 0;
        
        return {
          success: true,
          message: `Audit preparation package generated for ${args.auditType} covering ${args.auditAreas.length} areas`,
          data: {
            worksheetName,
            auditDetails: {
              auditType: args.auditType,
              auditPeriod: args.auditPeriod,
              auditorFirm: args.auditorFirm || "TBD",
              scopeAreas: args.auditAreas.length,
              complianceFrameworks: args.complianceFrameworks.length
            },
            readinessAssessment: {
              overallReadiness: readinessPercentage.toFixed(1) + "%",
              readyItems,
              totalItems,
              status: readinessPercentage >= 90 ? "✅ Audit Ready" :
                     readinessPercentage >= 75 ? "🟡 Nearly Ready" : "🔴 Preparation Needed"
            },
            nextSteps: [
              "Review and complete all pending documentation",
              "Confirm audit team availability and schedules", 
              "Finalize prepared-by-client (PBC) list",
              "Brief key personnel on audit procedures and expectations"
            ]
          }
        };
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  }
];