import { Tool, ToolResult } from '../types/index.js';
import { PythonBridge } from '../utils/python-bridge.js';

const pythonBridge = new PythonBridge();

export const rentalTools: Tool[] = [
  {
    name: "rental_generate_rent_roll",
    description: "Generate rent roll report for a property",
    inputSchema: {
      type: "object",
      properties: {
        propertyId: { type: "string" },
        asOfDate: { type: "string", format: "date", description: "Date in YYYY-MM-DD format" }
      },
      required: ["propertyId"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'rental_management',
          function: 'RentalPropertyManager.calculate_rent_roll',
          args: [args.propertyId],
          kwargs: args.asOfDate ? { as_of_date: args.asOfDate } : {}
        });
        return result;
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "rental_calculate_vacancy_rate",
    description: "Calculate vacancy rate and rental loss for a property",
    inputSchema: {
      type: "object",
      properties: {
        propertyId: { type: "string" },
        startDate: { type: "string", format: "date" },
        endDate: { type: "string", format: "date" }
      },
      required: ["propertyId", "startDate", "endDate"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'rental_management',
          function: 'RentalPropertyManager.calculate_vacancy_rate',
          args: [args.propertyId, args.startDate, args.endDate]
        });
        return result;
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "rental_lease_expiration_report",
    description: "Generate report of upcoming lease expirations",
    inputSchema: {
      type: "object",
      properties: {
        monthsAhead: { type: "number", default: 3, description: "Number of months to look ahead" }
      }
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'rental_management',
          function: 'RentalPropertyManager.generate_lease_expiration_report',
          args: [args.monthsAhead || 3]
        });
        return result;
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "rental_calculate_noi",
    description: "Calculate Net Operating Income for a property",
    inputSchema: {
      type: "object",
      properties: {
        propertyId: { type: "string" },
        year: { type: "number" }
      },
      required: ["propertyId", "year"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'rental_management',
          function: 'RentalPropertyManager.calculate_noi',
          args: [args.propertyId, args.year]
        });
        return result;
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "rental_calculate_cap_rate",
    description: "Calculate capitalization rate for a property",
    inputSchema: {
      type: "object",
      properties: {
        propertyId: { type: "string" },
        propertyValue: { type: "number" },
        year: { type: "number" }
      },
      required: ["propertyId", "propertyValue", "year"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'rental_management',
          function: 'RentalPropertyManager.calculate_cap_rate',
          args: [args.propertyId, args.propertyValue, args.year]
        });
        return result;
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "rental_project_cash_flow",
    description: "Project property cash flows with financing analysis",
    inputSchema: {
      type: "object",
      properties: {
        propertyId: { type: "string" },
        years: { type: "number" },
        initialInvestment: { type: "number" },
        loanAmount: { type: "number", default: 0 },
        loanRate: { type: "number", default: 0 },
        loanTermYears: { type: "number", default: 0 }
      },
      required: ["propertyId", "years", "initialInvestment"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'rental_management',
          function: 'RentalPropertyManager.project_cash_flow',
          args: [
            args.propertyId, 
            args.years, 
            args.initialInvestment,
            args.loanAmount || 0,
            args.loanRate || 0,
            args.loanTermYears || 0
          ]
        });
        return result;
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "rental_rent_comparables_analysis",
    description: "Analyze comparable rents for market analysis",
    inputSchema: {
      type: "object",
      properties: {
        unitId: { type: "string" },
        comparableProperties: { 
          type: "array", 
          items: { type: "string" },
          description: "List of comparable property IDs"
        }
      },
      required: ["unitId", "comparableProperties"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'rental_management',
          function: 'RentalPropertyManager.analyze_rent_comparables',
          args: [args.unitId, args.comparableProperties]
        });
        return result;
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  },

  {
    name: "rental_add_property",
    description: "Add a new rental property to the system",
    inputSchema: {
      type: "object",
      properties: {
        propertyId: { type: "string" },
        name: { type: "string" },
        address: { type: "string" },
        propertyType: { type: "string" },
        totalUnits: { type: "number" },
        yearBuilt: { type: "number" },
        amenities: { type: "array", items: { type: "string" } }
      },
      required: ["propertyId", "name", "address", "propertyType", "totalUnits"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        // This would interact with the rental management system
        // For now, return success - actual implementation would store the data
        return {
          success: true,
          data: { propertyId: args.propertyId },
          message: `Added property: ${args.name}`
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
    name: "rental_add_unit",
    description: "Add a rental unit to a property",
    inputSchema: {
      type: "object",
      properties: {
        unitId: { type: "string" },
        propertyId: { type: "string" },
        unitNumber: { type: "string" },
        squareFeet: { type: "number" },
        bedrooms: { type: "number" },
        bathrooms: { type: "number" },
        unitType: { type: "string" },
        amenities: { type: "array", items: { type: "string" } },
        marketRent: { type: "number" }
      },
      required: ["unitId", "propertyId", "unitNumber", "squareFeet", "bedrooms", "bathrooms", "marketRent"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        return {
          success: true,
          data: { unitId: args.unitId },
          message: `Added unit: ${args.unitNumber}`
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
    name: "rental_add_lease",
    description: "Add a new lease agreement",
    inputSchema: {
      type: "object",
      properties: {
        leaseId: { type: "string" },
        unitId: { type: "string" },
        tenantId: { type: "string" },
        startDate: { type: "string", format: "date" },
        endDate: { type: "string", format: "date" },
        monthlyRent: { type: "number" },
        securityDeposit: { type: "number" },
        escalationRate: { type: "number", default: 0 },
        escalationFrequency: { type: "string", default: "annually" }
      },
      required: ["leaseId", "unitId", "tenantId", "startDate", "endDate", "monthlyRent", "securityDeposit"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        return {
          success: true,
          data: { leaseId: args.leaseId },
          message: `Added lease: ${args.leaseId}`
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
    name: "rental_add_tenant",
    description: "Add a new tenant to the system",
    inputSchema: {
      type: "object",
      properties: {
        tenantId: { type: "string" },
        name: { type: "string" },
        contactInfo: { type: "object" },
        creditScore: { type: "number" },
        employmentInfo: { type: "object" }
      },
      required: ["tenantId", "name", "contactInfo"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        return {
          success: true,
          data: { tenantId: args.tenantId },
          message: `Added tenant: ${args.name}`
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