import { Tool, ToolResult } from '../types/index.js';
import { LeaseLifecycleManager, Lease } from '../property/lease-management.js';
import { PropertyInvestmentAnalyzer, PropertyInvestment } from '../property/property-investment.js';
import { MaintenanceCapExManager, MaintenanceRequest, CapExProject } from '../property/maintenance-capex.js';

const leaseManager = new LeaseLifecycleManager();
const investmentAnalyzer = new PropertyInvestmentAnalyzer();
const maintenanceManager = new MaintenanceCapExManager();

export const propertyTools: Tool[] = [
  {
    name: "property_lease_expiration_analysis",
    description: "Comprehensive analysis of lease expirations and renewal risk assessment",
    inputSchema: {
      type: "object",
      properties: {
        leases: {
          type: "array",
          items: {
            type: "object",
            properties: {
              leaseId: { type: "string" },
              propertyId: { type: "string" },
              unitId: { type: "string" },
              tenantId: { type: "string" },
              tenantName: { type: "string" },
              leaseStartDate: { type: "string", format: "date" },
              leaseEndDate: { type: "string", format: "date" },
              currentRent: { type: "number" },
              securityDeposit: { type: "number" },
              leaseType: { type: "string", enum: ["residential", "commercial", "retail", "office", "industrial"] },
              status: { type: "string", enum: ["active", "expired", "terminated", "pending", "holdover"] }
            },
            required: ["leaseId", "propertyId", "tenantName", "leaseStartDate", "leaseEndDate", "currentRent", "leaseType"]
          }
        },
        timeframe: { 
          type: "string", 
          enum: ["next_30_days", "next_90_days", "next_6_months", "next_12_months", "next_24_months"],
          default: "next_12_months"
        }
      },
      required: ["leases"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const leases = args.leases.map((lease: any) => ({
          ...lease,
          rentEscalations: lease.rentEscalations || [],
          renewalOptions: lease.renewalOptions || [],
          paymentTerms: lease.paymentTerms || "Monthly",
          specialClauses: lease.specialClauses || []
        })) as Lease[];

        // Generate lease analysis
        const leaseAnalyses = leaseManager.analyzeLeasePortfolio(leases);
        
        // Generate expiration analysis
        const expirationAnalysis = leaseManager.generateExpirationAnalysis(leases, args.timeframe || "next_12_months");

        return {
          success: true,
          data: {
            expirationAnalysis,
            leaseAnalyses: leaseAnalyses.slice(0, 20), // Top 20 for response size
            summary: {
              totalExpiringLeases: expirationAnalysis.expiringLeases.totalLeases,
              totalRentAtRisk: expirationAnalysis.expiringLeases.totalRent,
              expectedRenewals: expirationAnalysis.renewalProjections.expectedRenewals,
              projectedVacancy: expirationAnalysis.renewalProjections.projectedVacancy,
              annualRevenueAtRisk: expirationAnalysis.renewalProjections.revenueAtRisk
            }
          },
          message: `Lease expiration analysis completed for ${args.timeframe.replace('_', ' ')}. Found ${expirationAnalysis.expiringLeases.totalLeases} expiring leases with ${expirationAnalysis.renewalProjections.revenueAtRisk.toLocaleString()} in annual revenue at risk.`
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
    name: "property_investment_analysis",
    description: "Comprehensive property investment performance analysis with ROI, cap rates, and cash flow projections",
    inputSchema: {
      type: "object",
      properties: {
        property: {
          type: "object",
          properties: {
            propertyId: { type: "string" },
            propertyName: { type: "string" },
            acquisitionPrice: { type: "number" },
            acquisitionDate: { type: "string", format: "date" },
            propertyType: { type: "string", enum: ["residential", "commercial", "retail", "office", "industrial", "mixed_use"] },
            squareFeet: { type: "number" },
            units: { type: "number" },
            currentMarketValue: { type: "number" },
            financingDetails: {
              type: "object",
              properties: {
                loanAmount: { type: "number" },
                interestRate: { type: "number" },
                loanTermYears: { type: "number" },
                monthlyPayment: { type: "number" },
                remainingBalance: { type: "number" }
              },
              required: ["loanAmount", "interestRate", "monthlyPayment", "remainingBalance"]
            },
            operatingMetrics: {
              type: "object",
              properties: {
                grossRentalIncome: { type: "number" },
                operatingExpenses: { type: "number" },
                netOperatingIncome: { type: "number" },
                occupancyRate: { type: "number" }
              },
              required: ["grossRentalIncome", "operatingExpenses", "netOperatingIncome", "occupancyRate"]
            }
          },
          required: ["propertyId", "propertyName", "acquisitionPrice", "acquisitionDate", "propertyType", "currentMarketValue", "financingDetails", "operatingMetrics"]
        }
      },
      required: ["property"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const property = {
          ...args.property,
          taxInformation: args.property.taxInformation || {
            annualPropertyTaxes: args.property.currentMarketValue * 0.012, // 1.2% default
            depreciation: args.property.acquisitionPrice * 0.0364, // 27.5 year straight line
            lastAppraisalValue: args.property.currentMarketValue,
            assessedValue: args.property.currentMarketValue * 0.9
          }
        } as PropertyInvestment;

        // Generate investment analysis
        const analysis = investmentAnalyzer.analyzePropertyInvestment(property);

        return {
          success: true,
          data: {
            analysis,
            keyMetrics: {
              currentCapRate: analysis.returnMetrics.currentCapRate,
              cashOnCashReturn: analysis.returnMetrics.cashOnCashReturn,
              totalReturn: analysis.returnMetrics.totalReturn,
              internalRateOfReturn: analysis.returnMetrics.internalRateOfReturn,
              monthlyAverageRevenue: analysis.returnMetrics.cashFlow,
              equityPosition: analysis.returnMetrics.equityPosition,
              marketRisk: analysis.riskMetrics.marketRisk,
              loanToValueRatio: analysis.riskMetrics.loanToValueRatio
            }
          },
          message: `Investment analysis completed for ${property.propertyName}. Current cap rate: ${(analysis.returnMetrics.currentCapRate * 100).toFixed(2)}%, IRR: ${(analysis.returnMetrics.internalRateOfReturn * 100).toFixed(2)}%, Market Risk: ${analysis.riskMetrics.marketRisk}`
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
    name: "property_maintenance_analysis",
    description: "Comprehensive maintenance request analysis with vendor performance and predictive insights",
    inputSchema: {
      type: "object",
      properties: {
        propertyId: { type: "string" },
        maintenanceRequests: {
          type: "array",
          items: {
            type: "object",
            properties: {
              requestId: { type: "string" },
              propertyId: { type: "string" },
              unitId: { type: "string" },
              requestType: { type: "string", enum: ["emergency", "urgent", "routine", "preventive"] },
              category: { type: "string", enum: ["plumbing", "electrical", "hvac", "structural", "cosmetic", "landscaping", "security", "other"] },
              description: { type: "string" },
              requestDate: { type: "string", format: "date" },
              priorityLevel: { type: "number", minimum: 1, maximum: 5 },
              estimatedCost: { type: "number" },
              actualCost: { type: "number" },
              status: { type: "string", enum: ["open", "in_progress", "completed", "deferred", "cancelled"] },
              assignedVendor: { type: "string" },
              scheduledDate: { type: "string", format: "date" },
              completedDate: { type: "string", format: "date" },
              tenantSatisfactionRating: { type: "number", minimum: 1, maximum: 5 }
            },
            required: ["requestId", "propertyId", "requestType", "category", "description", "requestDate", "priorityLevel", "estimatedCost", "status"]
          }
        },
        propertyUnits: { type: "number", default: 10 },
        propertySquareFeet: { type: "number", default: 10000 }
      },
      required: ["propertyId", "maintenanceRequests"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const maintenanceRequests = args.maintenanceRequests as MaintenanceRequest[];

        // Generate maintenance analysis
        const analysis = maintenanceManager.analyzeMaintenancePortfolio(
          maintenanceRequests, 
          args.propertyId,
          args.propertyUnits || 10,
          args.propertySquareFeet || 10000
        );

        return {
          success: true,
          data: {
            analysis,
            keyMetrics: {
              totalRequests: analysis.maintenanceMetrics.totalRequests,
              avgResponseTime: analysis.maintenanceMetrics.avgResponseTime,
              avgCompletionTime: analysis.maintenanceMetrics.avgCompletionTime,
              costPerUnit: analysis.maintenanceMetrics.costPerUnit,
              tenantSatisfaction: analysis.maintenanceMetrics.tenantSatisfaction,
              emergencyRequestRate: analysis.maintenanceMetrics.emergencyRequestRate,
              deferredMaintenanceValue: analysis.maintenanceMetrics.deferredMaintenanceValue,
              topRiskAreas: analysis.predictiveMaintenance.riskAreas.slice(0, 3)
            }
          },
          message: `Maintenance analysis completed for property ${args.propertyId}. ${analysis.maintenanceMetrics.totalRequests} requests analyzed with ${analysis.maintenanceMetrics.avgResponseTime.toFixed(1)} hour average response time and ${analysis.maintenanceMetrics.tenantSatisfaction.toFixed(1)}/5.0 tenant satisfaction.`
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
    name: "property_capex_analysis", 
    description: "Capital expenditure analysis with project tracking, ROI analysis, and budget variance reporting",
    inputSchema: {
      type: "object",
      properties: {
        propertyId: { type: "string" },
        capexProjects: {
          type: "array",
          items: {
            type: "object",
            properties: {
              projectId: { type: "string" },
              propertyId: { type: "string" },
              projectName: { type: "string" },
              projectType: { type: "string", enum: ["roof_replacement", "hvac_upgrade", "flooring", "exterior_improvements", "structural", "technology", "amenities", "other"] },
              description: { type: "string" },
              plannedStartDate: { type: "string", format: "date" },
              plannedEndDate: { type: "string", format: "date" },
              actualStartDate: { type: "string", format: "date" },
              actualEndDate: { type: "string", format: "date" },
              budgetAmount: { type: "number" },
              actualCost: { type: "number" },
              status: { type: "string", enum: ["planned", "approved", "in_progress", "completed", "delayed", "cancelled"] },
              contractor: { type: "string" },
              expectedROI: { type: "number" },
              expectedValueAdd: { type: "number" },
              impactOnOccupancy: { type: "string", enum: ["none", "minimal", "moderate", "significant"] }
            },
            required: ["projectId", "propertyId", "projectName", "projectType", "description", "plannedStartDate", "plannedEndDate", "budgetAmount", "status"]
          }
        }
      },
      required: ["propertyId", "capexProjects"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const capexProjects = args.capexProjects as CapExProject[];

        // Generate CapEx analysis
        const analysis = maintenanceManager.analyzeCapExPortfolio(capexProjects, args.propertyId);

        return {
          success: true,
          data: {
            analysis,
            keyMetrics: {
              totalBudget: analysis.capExMetrics.totalBudget,
              spentToDate: analysis.capExMetrics.spentToDate,
              remainingBudget: analysis.capExMetrics.remainingBudget,
              projectedOverrun: analysis.capExMetrics.projectedOverrun,
              completedProjects: analysis.capExMetrics.completedProjects,
              onTimePercentage: analysis.capExMetrics.onTimePercentage,
              avgROI: analysis.capExMetrics.avgROI,
              totalValueAdd: analysis.capExMetrics.totalValueAdd,
              upcomingProjectsCount: analysis.timeline.upcomingProjects.length
            }
          },
          message: `CapEx analysis completed for property ${args.propertyId}. Budget: ${(analysis.capExMetrics.totalBudget / 1000000).toFixed(1)}M, Spent: ${(analysis.capExMetrics.spentToDate / 1000000).toFixed(1)}M, On-time rate: ${analysis.capExMetrics.onTimePercentage.toFixed(1)}%, Avg ROI: ${(analysis.capExMetrics.avgROI * 100).toFixed(1)}%`
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