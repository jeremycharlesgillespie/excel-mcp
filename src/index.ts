import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";

import { excelTools } from "./tools/excel-tools.js";
import { financialTools } from "./tools/financial-tools.js";
import { rentalTools } from "./tools/rental-tools.js";
import { expenseTools } from "./tools/expense-tools.js";
import { reportingTools } from "./tools/reporting-tools.js";
import { cashFlowTools } from "./tools/cash-flow-tools.js";
import { taxTools } from "./tools/tax-tools.js";
import { analyticsTools } from "./tools/analytics-tools.js";
import { chartTools } from "./tools/chart-tools.js";
import { complianceTools } from "./tools/compliance-tools.js";
import { propertyTools } from "./tools/property-tools-simple.js";

const server = new Server(
  {
    name: "excel-finance-mcp",
    version: "1.0.0",
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

const allTools = [
  ...excelTools,
  ...financialTools,
  ...rentalTools,
  ...expenseTools,
  ...reportingTools,
  ...cashFlowTools,
  ...taxTools,
  ...analyticsTools,
  ...chartTools,
  ...complianceTools,
  ...propertyTools,
];

server.setRequestHandler(ListToolsRequestSchema, async () => ({
  tools: allTools.map(tool => ({
    name: tool.name,
    description: tool.description,
    inputSchema: tool.inputSchema,
  })),
}));

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;
  
  const tool = allTools.find(t => t.name === name);
  if (!tool) {
    throw new Error(`Unknown tool: ${name}`);
  }
  
  try {
    const result = await tool.handler(args);
    return {
      content: [
        {
          type: "text",
          text: JSON.stringify(result, null, 2),
        },
      ],
    };
  } catch (error) {
    return {
      content: [
        {
          type: "text",
          text: `Error: ${error instanceof Error ? error.message : String(error)}`,
        },
      ],
      isError: true,
    };
  }
});

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("Excel Finance MCP Server running on stdio");
}

main().catch(console.error);