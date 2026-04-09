
export interface PlandayApiCredentials {
  clientId: string;
  refreshToken: string;
}

export interface Employee {
  id: number;
  firstName: string;
  lastName:string;
  salaryIdentifier: string | null;
}

export interface LeaveAccount {
  id: number;
  name: string;
  typeId: number;
  validityPeriod: {
    start: string | null;
    end: string | null;
  };
}

export interface LeaveAccountBalance {
    balance: number;
    unit: string;
}

export interface BalanceAdjustmentPayload {
    effectiveDate: string; // YYYY-MM-DD
    value: number;
    comment: string;
}

export interface AccountType {
    id: number;
    name: string;
    unit: string;
}

export interface TemplateDataRow {
    employeeId: number;
    salaryIdentifier: string | null;
    employeeName: string;
    accountId: number;
    accountName: string;
    validFrom: string;
    validTo: string;
    balanceDate: string;
    availableBalance: number;
    balanceUnit: string;
}

export interface AdjustmentReview {
    id: string;
    accountId: number;
    employeeName: string;
    accountName: string;
    availableBalance: number;
    newBalance: number;
    adjustment: number;
    unit?: string;
    timestamp?: string;
    effectiveDate: string; // YYYY-MM-DD
    validFrom?: string | null; // Constraint from file
    validTo?: string | null;   // Constraint from file
    comment: string;
    status?: 'pending' | 'success' | 'error';
    error?: string;
    isValidationError?: boolean; // Flag to distinguish pre-check errors from API errors
}