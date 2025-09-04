export interface ToolResult {
  success: boolean;
  data?: any;
  error?: string;
  message?: string;
}

export interface FinancialCalculationParams {
  rate?: number;
  periods?: number;
  presentValue?: number;
  futureValue?: number;
  payment?: number;
  cashFlows?: number[];
  initialInvestment?: number;
}

export interface LoanParams {
  principal: number;
  annualRate: number;
  years: number;
  paymentFrequency?: 'monthly' | 'quarterly' | 'semi-annually' | 'annually';
}

export interface PropertyParams {
  propertyId: string;
  name?: string;
  address?: string;
  propertyType?: string;
  totalUnits?: number;
  yearBuilt?: number;
}

export interface UnitParams {
  unitId: string;
  propertyId: string;
  unitNumber: string;
  squareFeet: number;
  bedrooms: number;
  bathrooms: number;
  marketRent: number;
}

export interface LeaseParams {
  leaseId: string;
  unitId: string;
  tenantId: string;
  startDate: string;
  endDate: string;
  monthlyRent: number;
  securityDeposit: number;
  escalationRate?: number;
}

export interface TenantParams {
  tenantId: string;
  name: string;
  contactInfo: Record<string, string>;
  creditScore?: number;
}

export interface ExpenseParams {
  expenseId: string;
  date: string;
  vendorId: string;
  amount: number;
  category: string;
  subcategory?: string;
  description?: string;
  invoiceNumber?: string;
  costCenter?: string;
  projectId?: string;
}

export interface VendorParams {
  vendorId: string;
  name: string;
  contactInfo: Record<string, string>;
  taxId?: string;
  paymentTerms?: string;
  w9OnFile?: boolean;
}

export interface BudgetParams {
  budgetId: string;
  name: string;
  fiscalYear: number;
  period: 'annual' | 'quarterly' | 'monthly';
  categories: Record<string, number>;
}

export interface ExcelOperationParams {
  filePath?: string;
  worksheetName?: string;
  range?: string;
  data?: any[][];
  startRow?: number;
  startCol?: number;
}

export interface PythonBridgeParams {
  module: string;
  function: string;
  args: any[];
  kwargs?: Record<string, any>;
}

export interface Tool {
  name: string;
  description: string;
  inputSchema: {
    type: "object";
    properties: Record<string, any>;
    required?: string[];
  };
  handler: (args: any) => Promise<ToolResult>;
}