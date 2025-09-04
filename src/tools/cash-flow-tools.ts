import { Tool, ToolResult } from '../types/index.js';
import { PythonBridge } from '../utils/python-bridge.js';

const pythonBridge = new PythonBridge();

const cashFlowTools: Tool[] = [
  {
    name: "cash_flow_statement",
    description: "Generate cash flow statement for specified period",
    inputSchema: {
      type: "object",
      properties: {
        startDate: { type: "string", format: "date" },
        endDate: { type: "string", format: "date" }
      },
      required: ["startDate", "endDate"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'cash_flow_tools',
          function: 'generate_cash_flow_statement',
          args: [args.startDate, args.endDate]
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
    name: "cash_flow_forecast",
    description: "Forecast future cash flows with scenario analysis",
    inputSchema: {
      type: "object",
      properties: {
        monthsAhead: { type: "number", default: 12 },
        scenarios: {
          type: "object",
          properties: {
            base: { type: "number", default: 1.0 },
            conservative: { type: "number", default: 0.9 },
            optimistic: { type: "number", default: 1.1 }
          }
        }
      }
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'cash_flow_tools',
          function: 'forecast_cash_flow',
          args: [args.monthsAhead || 12, args.scenarios]
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
    name: "cash_burn_analysis",
    description: "Analyze cash burn rate and calculate runway",
    inputSchema: {
      type: "object",
      properties: {
        monthsBack: { type: "number", default: 6 }
      }
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'cash_flow_tools',
          function: 'cash_burn_analysis',
          args: [args.monthsBack || 6]
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
    name: "working_capital_analysis",
    description: "Analyze working capital changes and efficiency",
    inputSchema: {
      type: "object",
      properties: {
        startDate: { type: "string", format: "date" },
        endDate: { type: "string", format: "date" }
      },
      required: ["startDate", "endDate"]
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'cash_flow_tools',
          function: 'working_capital_analysis',
          args: [args.startDate, args.endDate]
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
    name: "cash_flow_at_risk",
    description: "Calculate Cash Flow at Risk (CFaR) metric",
    inputSchema: {
      type: "object",
      properties: {
        confidenceLevel: { type: "number", default: 0.95, minimum: 0.01, maximum: 0.99 }
      }
    },
    handler: async (args: any): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'cash_flow_tools',
          function: 'cash_flow_at_risk',
          args: [args.confidenceLevel || 0.95]
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
    name: "liquidity_analysis",
    description: "Analyze current liquidity position and requirements",
    inputSchema: {
      type: "object",
      properties: {}
    },
    handler: async (): Promise<ToolResult> => {
      try {
        const result = await pythonBridge.callPythonFunction({
          module: 'cash_flow_tools',
          function: 'liquidity_analysis',
          args: []
        });
        return result;
      } catch (error) {
        return {
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  }
];

export { cashFlowTools };