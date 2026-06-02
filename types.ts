
export interface PlandayApiCredentials {
  clientId: string;
  refreshToken: string;
}

export interface Employee {
  id: number;
  firstName: string;
  lastName:string;
  salaryIdentifier: string | null;
  departments?: number[] | any[];
  departmentId?: number;
  primaryDepartmentId?: number;
  employeeGroups?: number[] | any[];
  employeeGroupId?: number;
  employeeGroupIds?: number[];
  employeeType?: number | any;
  employeeTypeId?: number;
}

export interface Department {
  id: number | string;
  name: string;
}

export interface EmployeeGroup {
  id: number | string;
  name: string;
}

export interface EmployeeType {
  id: number | string;
  name: string;
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
    absenceType?: string;
    accruingRate?: {
        value: number;
        unit: {
            type: string;
        };
    };
}

export interface TemplateDataRow {
    employeeId: number;
    salaryIdentifier: string | null;
    employeeName: string;
    accountId: number;
    accountName: string;
    accountTypeCategory: 'FLEX/TOIL' | 'Fixed' | 'Accrued' | 'Unknown';
    validFrom: string;
    validTo: string;
    balanceDate: string;
    availableBalance: number;
    balanceUnit: string;
}

export interface AdjustmentReview {
    id: string;
    accountId: number;
    employeeId?: number;
    employeeName: string;
    salaryIdentifier?: string | null;
    accountName: string;
    accountTypeCategory?: 'FLEX/TOIL' | 'Fixed' | 'Accrued' | 'Unknown';
    balanceDate?: string;
    availableBalance: number | string;
    newBalance: number | string;
    adjustment: number;
    unit?: string;
    timestamp?: string;
    effectiveDate: string; // YYYY-MM-DD
    validFrom?: string | null; // Constraint from file
    validTo?: string | null;   // Constraint from file
    comment: string;
    status?: 'pending' | 'success' | 'error' | 'skipped';
    error?: string;
    isValidationError?: boolean; // Flag to distinguish pre-check errors from API errors
    postAdjustmentBalance?: number;
}