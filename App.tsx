
import React, { useState, useEffect, useRef, useMemo, useCallback } from 'react';
import type { PlandayApiCredentials, AccountType, TemplateDataRow, AdjustmentReview, Department, EmployeeGroup, EmployeeType } from './types';
import { initializeService, fetchEmployees, fetchLeaveAccounts, fetchAccountBalance, postBalanceAdjustment, postFlexBalanceAdjustment, fetchAccountTypes, fetchDepartments, fetchEmployeeGroups, fetchEmployeeTypes, fetchPortalInfo } from './services/plandayService';
import * as XLSX from 'xlsx-js-style';

// --- Custom Hooks ---
function useUndoableState<T>(initialValue: T, maxHistory: number = 5) {
    const [state, setState] = useState<{ past: T[]; present: T; future: T[] }>({
        past: [],
        present: initialValue,
        future: []
    });

    const set = useCallback((action: React.SetStateAction<T>) => {
        setState(current => {
            const nextPresent = typeof action === 'function' ? (action as any)(current.present) : action;
            if (current.present === nextPresent) return current;
            const newPast = [...current.past, current.present];
            if (newPast.length > maxHistory) newPast.shift();
            return { past: newPast, present: nextPresent, future: [] };
        });
    }, [maxHistory]);

    const undo = useCallback(() => {
        setState(current => {
            if (current.past.length === 0) return current;
            const previous = current.past[current.past.length - 1];
            const newPast = current.past.slice(0, current.past.length - 1);
            return { past: newPast, present: previous, future: [current.present, ...current.future] };
        });
    }, []);

    const redo = useCallback(() => {
        setState(current => {
            if (current.future.length === 0) return current;
            const next = current.future[0];
            const newFuture = current.future.slice(1);
            const newPast = [...current.past, current.present];
            if (newPast.length > maxHistory) newPast.shift();
            return { past: newPast, present: next, future: newFuture };
        });
    }, [maxHistory]);

    const clearHistory = useCallback((newPresent: T) => {
        setState({ past: [], present: newPresent, future: [] });
    }, []);

    return [state.present, set, undo, redo, state.past.length, state.future.length, clearHistory] as const;
}

// --- Utility Functions ---

// Formats a YYYY-MM-DD string into display format based on preference
const formatDateForDisplay = (dateString: string | null | undefined, format: 'EU' | 'US' = 'EU'): string => {
    if (!dateString) return 'N/A';
    try {
        // Safe check for YYYY-MM-DD
        if (!/^\d{4}-\d{2}-\d{2}$/.test(dateString)) return 'Invalid Date';
        
        const [year, month, day] = dateString.split('-');
        
        if (format === 'US') {
            return `${month}/${day}/${year}`;
        }
        return `${day}/${month}/${year}`;
    } catch { return 'Invalid Date'; }
};

const getAccountCategory = (typeId: number, accountTypes: AccountType[]): 'FLEX/TOIL' | 'Fixed' | 'Accrued' | 'Unknown' => {
    const type = accountTypes.find(t => t.id === typeId);
    if (!type) return 'Unknown';
    if (type.absenceType === 'Flextime') return 'FLEX/TOIL';
    if (type.accruingRate?.value === 0 && type.accruingRate?.unit?.type === 'Percent') return 'Fixed';
    return 'Accrued';
};

const formatDateToYYYYMMDD = (dateString: string): string => {
    if (!dateString || !/^\d{2}\/\d{2}\/\d{4}$/.test(dateString)) return dateString;
    const [day, month, year] = dateString.split('/');
    return `${year}-${month}-${day}`;
};

/**
 * Smart Date Parser
 * Parses input into YYYY-MM-DD string.
 * Strictly manages Year/Month/Day to avoid timezone shifts.
 */
const parseDateToIso = (input: any, formatPreference: 'EU' | 'US'): string | null => {
    if (!input) return null;
    
    // 1. Handle JS Date Objects (from Excel Cell Date types)
    if (input instanceof Date) {
         if (isNaN(input.getTime())) return null;
         // Always use local methods!
         const y = input.getFullYear();
         const m = String(input.getMonth() + 1).padStart(2, '0');
         const d = String(input.getDate()).padStart(2, '0');
         return `${y}-${m}-${d}`;
    }

    // 2. Catch raw serial numbers (General formatting)
    if (typeof input === 'number' || /^\d{5}$/.test(String(input))) {
         const serial = parseInt(input, 10);
         if (serial > 20000 && serial < 80000) { // Valid bounds for recent dates
             // 25569 is the difference in days between Jan 1 1900 and Jan 1 1970
             const dateObj = new Date(Math.round((serial - 25569) * 86400 * 1000));
             
             // Use UTC methods here because we calculated pure absolute time from the Unix Epoch
             const y = dateObj.getUTCFullYear();
             const m = String(dateObj.getUTCMonth() + 1).padStart(2, '0');
             const d = String(dateObj.getUTCDate()).padStart(2, '0');
             
             return `${y}-${m}-${d}`;
         }
    }

    const str = String(input).trim();
    if (!str) return null;

    // 3. Fallback string parsing for Text-formatted cells
    // Check for ISO format (YYYY-MM-DD) - Always prioritized
    if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
        // Simple validation
        const d = new Date(str);
        if (!isNaN(d.getTime())) return str;
    }

    // 4. Parse Text Formats (e.g., 1/4/25, 01-04-2025)
    // Regex: (1 or 2 digits) [separator] (1 or 2 digits) [separator] (2 or 4 digits)
    const match = str.match(/(^|\b)(\d{1,2})[\/\.\-](\d{1,2})[\/\.\-](\d{2}|\d{4})\b/);
    if (match) {
        let p1 = parseInt(match[2], 10);
        let p2 = parseInt(match[3], 10);
        let year = parseInt(match[4], 10);

        // Handle 2-digit years (Assume 2000s)
        if (year < 100) year += 2000;

        let day, month;

        if (formatPreference === 'US') {
            // MM/DD/YYYY
            month = p1;
            day = p2;
        } else {
            // EU: DD/MM/YYYY
            day = p1;
            month = p2;
        }

        // Basic logical validation
        if (month < 1 || month > 12) return null;
        if (day < 1 || day > 31) return null;

        // Strict Date validation (e.g. checks Feb 30th -> invalid)
        // We do NOT use ISOString() here to avoid timezone shifts.
        const dateObj = new Date(year, month - 1, day);
        if (dateObj.getFullYear() === year && dateObj.getMonth() === month - 1 && dateObj.getDate() === day) {
             return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        }
    }

    return null;
};

interface ColumnDetectionResult {
    format: 'EU' | 'US';
    source: 'detected' | 'fallback' | 'inherited' | 'empty'; // 'detected' means we found unambiguous dates, 'fallback' means ambiguous or empty, 'inherited' means copied from validFrom/To, 'empty' means no real dates found
    hasConflict: boolean;
    conflictDetails: string[];
    exampleRowIndex?: number; // 0-based index from the JSON array
    exampleReason?: string;
}

/**
 * Scans a list of rows to detect format per column.
 * - Detects conflicts (e.g. Row 2 is EU, Row 5 is US).
 * - Returns detailed result.
 */
const detectColumnFormat = (rows: any[], columnKey: string, fallbackFormat: 'EU' | 'US'): ColumnDetectionResult => {
    let unambiguousUS: { row: number, val: string, index: number }[] = [];
    let unambiguousEU: { row: number, val: string, index: number }[] = [];
    let ambiguousSample: { row: number, val: string, index: number } | null = null;
    let hasAnyDateLike = false;
    
    rows.forEach((row, index) => {
        const val = row[columnKey];
        if (val === undefined || val === null || val === '') return;
        
        // Treat checking 'N/A' as not a date at all.
        if (typeof val === 'string' && val.trim().toUpperCase() === 'N/A') return;

        if (val instanceof Date) { hasAnyDateLike = true; return; } // Unambiguous Date obj
        if (typeof val === 'number' || /^\d{5}$/.test(String(val))) { hasAnyDateLike = true; return; } // Unambiguous Serial

        const str = String(val).trim();
        
        // Skip ISO
        if (/^\d{4}-\d{2}-\d{2}$/.test(str)) { hasAnyDateLike = true; return; }

        const match = str.match(/(^|\b)(\d{1,2})[\/\.\-](\d{1,2})[\/\.\-](\d{2}|\d{4})\b/);
        if (match) {
            hasAnyDateLike = true;
            const p1 = parseInt(match[2], 10);
            const p2 = parseInt(match[3], 10);
            const year = parseInt(match[4], 10);
            
            // Invalid checks
            if (p1 > 31 && p2 > 31) return; // garbage
            if (p1 === 0 || p2 === 0) return;

            // Unambiguous US Check: Month First.
            // If p1 <= 12 AND p2 > 12. Example: 04/30/2025. 
            if (p1 <= 12 && p2 > 12) {
                unambiguousUS.push({ row: index + 2, val: str, index: index }); // +2 for Excel Row (0-based + header)
            }

            // Unambiguous EU Check: Day First.
            // If p1 > 12 AND p2 <= 12. Example: 30/04/2025.
            if (p1 > 12 && p2 <= 12) {
                unambiguousEU.push({ row: index + 2, val: str, index: index });
            }

            // Capture first ambiguous one for fallback example (e.g. 01/02/2025)
            if (!ambiguousSample && p1 <= 12 && p2 <= 12) {
                ambiguousSample = { row: index + 2, val: str, index: index };
            }
        }
    });

    // Check for Critical Conflict within the same column
    if (unambiguousUS.length > 0 && unambiguousEU.length > 0) {
        // Collect first few examples
        const examples = [
            ...unambiguousUS.slice(0, 2).map(i => `Row ${i.row}: ${i.val} (US Format)`),
            ...unambiguousEU.slice(0, 2).map(i => `Row ${i.row}: ${i.val} (EU Format)`)
        ];
        
        return {
            format: fallbackFormat, // Irrelevant, blocking error
            source: 'detected',
            hasConflict: true,
            conflictDetails: examples
        };
    }

    if (unambiguousUS.length > 0) return { 
        format: 'US', 
        source: 'detected', 
        hasConflict: false, 
        conflictDetails: [], 
        exampleRowIndex: unambiguousUS[0].index,
        exampleReason: `Detected US format (Month/Day/Year) in Row ${unambiguousUS[0].row}: "${unambiguousUS[0].val}"`
    };

    if (unambiguousEU.length > 0) return { 
        format: 'EU', 
        source: 'detected', 
        hasConflict: false, 
        conflictDetails: [], 
        exampleRowIndex: unambiguousEU[0].index,
        exampleReason: `Detected EU format (Day/Month/Year) in Row ${unambiguousEU[0].row}: "${unambiguousEU[0].val}"`
    };
    
    // If there were no dates found at all (e.g. all empty or 'N/A')
    if (!hasAnyDateLike) {
        return { 
            format: fallbackFormat, 
            source: 'empty', 
            hasConflict: false, 
            conflictDetails: []
        };
    }
    
    // If mixed (conflicting) or all ambiguous, use fallback
    return { 
        format: fallbackFormat, 
        source: 'fallback', 
        hasConflict: false, 
        conflictDetails: [],
        exampleRowIndex: ambiguousSample ? ambiguousSample.index : undefined,
        exampleReason: ambiguousSample ? `Date "${ambiguousSample.val}" in Row ${ambiguousSample.row} is ambiguous (could be EU or US). Defaulting to app setting (${fallbackFormat}).` : undefined
    };
};

const getTodayYYYYMMDD = () => {
    const local = new Date();
    const year = local.getFullYear();
    const month = String(local.getMonth() + 1).padStart(2, '0');
    const day = String(local.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
};

const EditableAdjustmentCell = ({ value, onChange, disabled, isBalance = false }: { value: number, onChange: (val: number) => void, disabled?: boolean, isBalance?: boolean }) => {
    const [localValue, setLocalValue] = useState(value.toFixed(2));
    const isError = isNaN(parseFloat(localValue)) && localValue !== '-' && localValue !== '';
    const numValue = parseFloat(localValue);
    const isZero = !isBalance && !isError && numValue === 0 && localValue !== '';
    const isNegative = !isError && numValue < 0;
    
    useEffect(() => {
        setLocalValue(value.toFixed(2));
    }, [value]);

    const titleText = isZero && !isBalance ? "An adjustment cannot be 0 and therefore, this adjustment will get skipped if you don't change it to a positive/negative value." : undefined;

    return (
        <input 
            type="text"
            title={titleText}
            className={`w-20 text-right bg-transparent border-b border-dashed focus:outline-none focus:border-blue-500 font-mono ${isError ? 'border-red-500 !text-red-500' : isZero ? 'border-orange-500 hover:border-orange-600 !text-orange-600' : ((isNegative && !isBalance) ? 'border-gray-300 hover:border-gray-400 !text-red-600' : (!isNegative && !isBalance && numValue !== 0 ? 'border-gray-300 hover:border-gray-400 !text-green-600' : 'border-gray-300 hover:border-gray-400 !text-gray-800'))}`}
            value={localValue}
            onChange={(e) => setLocalValue(e.target.value)}
            disabled={disabled}
            onBlur={() => {
                const parsed = parseFloat(localValue);
                if (!isNaN(parsed)) {
                    onChange(parsed);
                    setLocalValue(parsed.toFixed(2)); // Re-format
                } else {
                    onChange(0);
                    setLocalValue('0.00');
                }
            }}
        />
    )
};

// --- SVG Icons ---
const FilterIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg className={className} xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M3 4a1 1 0 011-1h16a1 1 0 011 1v2.586a1 1 0 01-.293.707l-6.414 6.414a1 1 0 00-.293.707V17l-4 4v-6.586a1 1 0 00-.293-.707L3.293 7.293A1 1 0 013 6.586V4z" /></svg>
);
const CheckIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg className={className} xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M5 13l4 4L19 7" /></svg>
);
const CopyIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg className={className} xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M8 16H6a2 2 0 01-2-2V6a2 2 0 012-2h8a2 2 0 012 2v2m-6 12h8a2 2 0 002-2v-8a2 2 0 00-2-2h-8a2 2 0 00-2 2v8a2 2 0 002 2z" /></svg>
);
const InfoIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg className={className} xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
    </svg>
);
const DownloadIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg className={className} xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
    </svg>
);
const ExclamationIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg className={className} xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
         <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" />
    </svg>
);
const CalendarIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg className={className} xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M8 7V3m8 4V3m-9 8h10M5 21h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z" />
    </svg>
);
const RefreshIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg className={className} xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15" />
    </svg>
);
const UploadCloudIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg className={className} xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
    </svg>
);
const TrashIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg className={className} xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" />
    </svg>
);
const SortAscIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg className={className} xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M3 4h13M3 8h9m-9 4h6m4 0l4-4m0 0l4 4m-4-4v12" />
    </svg>
);
const SortDescIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg className={className} xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M3 4h13M3 8h9m-9 4h5m4 0v12m0 0l4-4m-4 4l-4-4" />
    </svg>
);

// --- UI Components ---
const PageHeader: React.FC = () => (
    <div className="text-center">
        <h1 className="text-4xl font-bold text-gray-800 flex items-center justify-center gap-3">
            Planday Bulk Leave Adjustments
            <span className="bg-blue-500 text-white text-xs font-semibold px-2.5 py-0.5 rounded-full">BETA</span>
        </h1>
        <p className="mt-2 text-lg text-gray-500">Update Leave and Flex/TOIL balances in bulk from Excel files or Table</p>
    </div>
);

const Loader: React.FC<{ text: string }> = ({ text }) => (
    <div className="flex items-center text-gray-500"><svg className="animate-spin -ml-1 mr-3 h-5 w-5 text-blue-500" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24"><circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle><path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path></svg><span>{text}</span></div>
);

const ProgressBar: React.FC<{ progress: number; text: string }> = ({ progress, text }) => (
    <div className="w-full flex flex-col items-center">
         <div className="w-full bg-gray-200 rounded-full h-2.5 mb-2">
            <div className="bg-blue-600 h-2.5 rounded-full transition-all duration-300 ease-out" style={{ width: `${progress}%` }}></div>
        </div>
        <div className="flex justify-between w-full text-xs text-gray-500 font-medium">
            <span>{text}</span>
            <span>{Math.round(progress)}%</span>
        </div>
    </div>
);

const Stepper: React.FC<{ current: number; steps: { title: string; subtitle: string }[] }> = ({ current, steps }) => {
    return (
        <nav aria-label="Progress">
            <ol role="list" className="flex items-center">
                {steps.map((step, index) => {
                    const isResultsStep = index === steps.length - 1;
                    const isCompleted = index < current;
                    // If it is the results step and active, show it as completed/green check
                    const isResultsActive = isResultsStep && index === current;
                    const showCheck = isCompleted || isResultsActive;

                    return (
                        <li key={step.title} className={`relative ${index !== steps.length - 1 ? 'flex-1' : ''}`}>
                            <div className="flex items-center text-sm font-medium">
                                <span className={`flex h-10 w-10 flex-shrink-0 items-center justify-center rounded-full ${showCheck ? 'bg-green-600' : index === current ? 'bg-blue-600' : 'bg-gray-300'}`}>
                                    {showCheck ? <CheckIcon className="h-6 w-6 text-white" /> : <span className={index === current ? 'text-white' : 'text-gray-600'}>{index + 1}</span>}
                                </span>
                                <div className="ml-4 hidden md:block">
                                    <span className={`block text-sm font-semibold ${showCheck ? 'text-green-600' : index === current ? 'text-blue-600' : 'text-gray-500'}`}>{step.title}</span>
                                    <span className="block text-sm text-gray-500">{step.subtitle}</span>
                                </div>
                            </div>
                            {index !== steps.length - 1 && (
                                <div className={`absolute top-5 left-10 -ml-px mt-px h-0.5 w-full ${index < current ? 'bg-green-600' : 'bg-gray-300'}`} aria-hidden="true" />
                            )}
                        </li>
                    );
                })}
            </ol>
        </nav>
    );
};

// --- Main App & Step Components ---
type AppStep = 'auth' | 'configure' | 'upload' | 'review' | 'processing' | 'summary';
type ValidityMode = 'current' | 'current_future' | 'custom';
type UpdateMethod = 'excel' | 'editor';

const getStepConfig = (method: UpdateMethod) => ({
    labels: [
        { title: 'Authentication', subtitle: 'Connect to Planday' },
        method === 'excel' 
            ? { title: 'Configure', subtitle: 'Download template' }
            : { title: 'Select Accounts', subtitle: 'Choose accounts to update' },
        ...(method === 'excel' ? [{ title: 'Upload', subtitle: 'Upload Excel file' }] : []),
        { title: 'Review', subtitle: method === 'excel' ? 'Final review' : 'Make Adjustments' },
        { title: 'Update Process', subtitle: 'Adj. balances' },
        { title: 'Results', subtitle: 'View results' },
    ],
    order: method === 'excel'
        ? ['auth', 'configure', 'upload', 'review', 'processing', 'summary'] as const
        : ['auth', 'configure', 'review', 'processing', 'summary'] as const
});

interface DateReportExample {
    rowNumber: number;
    employee: string;
    account: string;
    columnName: string;
    rawValue: string;
    convertedValue: string;
    detectedFormat: string;
}

interface MultiSelectDropdownProps {
    label: string;
    pluralLabel?: string;
    options: { id: number | string; name: string }[];
    selectedIds: Set<string>;
    onChange: (selectedIds: Set<string>) => void;
}

const MultiSelectDropdown: React.FC<MultiSelectDropdownProps> = ({
    label,
    pluralLabel,
    options,
    selectedIds,
    onChange
}) => {
    const [isOpen, setIsOpen] = useState(false);
    const [searchTerm, setSearchTerm] = useState('');
    const dropdownRef = useRef<HTMLDivElement>(null);

    useEffect(() => {
        const handleClickOutside = (event: MouseEvent) => {
            if (dropdownRef.current && !dropdownRef.current.contains(event.target as Node)) {
                setIsOpen(false);
            }
        };
        document.addEventListener("mousedown", handleClickOutside);
        return () => document.removeEventListener("mousedown", handleClickOutside);
    }, []);

    const allChecked = selectedIds.size === options.length && options.length > 0;
    const isIndeterminate = selectedIds.size > 0 && selectedIds.size < options.length;

    const filteredOptions = options.filter(o => o.name.toLowerCase().includes(searchTerm.toLowerCase()));

    const isSearchActive = searchTerm.trim().length > 0;
    const allFilteredChecked = filteredOptions.length > 0 && filteredOptions.every(o => selectedIds.has(String(o.id)));
    const isFilteredIndeterminate = !allFilteredChecked && filteredOptions.some(o => selectedIds.has(String(o.id)));

    const handleSelectAll = () => {
        if (isSearchActive) {
            const newSet = new Set(selectedIds);
            if (allFilteredChecked) {
                filteredOptions.forEach(o => newSet.delete(String(o.id)));
            } else {
                filteredOptions.forEach(o => newSet.add(String(o.id)));
            }
            onChange(newSet);
        } else {
            if (allChecked) {
                onChange(new Set());
            } else {
                onChange(new Set(options.map(o => String(o.id))));
            }
        }
    };

    const handleToggle = (id: string) => {
        const newSet = new Set(selectedIds);
        if (newSet.has(id)) {
            newSet.delete(id);
        } else {
            newSet.add(id);
        }
        onChange(newSet);
    };

    const displayPlural = pluralLabel || label + 's';
    const displayValue = allChecked || selectedIds.size === 0 ? `All ${displayPlural}` : `${selectedIds.size} selected`;

    return (
        <div className="relative" ref={dropdownRef}>
            <div 
                className="w-full bg-white border border-gray-300 rounded-md shadow-sm py-2 px-3 flex justify-between items-center cursor-pointer focus:border-blue-500 focus:ring-blue-500 sm:text-sm h-[38px]"
                onClick={() => setIsOpen(!isOpen)}
            >
                <span className="truncate text-gray-700">{displayValue}</span>
                <span className="text-gray-500 text-xs ml-2">▼</span>
            </div>
            
            {isOpen && (
                <div className="absolute z-10 mx-0 mt-1 w-full bg-white border border-gray-300 rounded-md shadow-lg max-h-60 overflow-y-auto focus:outline-none">
                    <div className="p-2 border-b border-gray-200 sticky top-0 bg-white z-20">
                        <input 
                            type="text" 
                            className="w-full border border-gray-300 rounded px-2 py-1.5 text-sm focus:outline-none focus:border-blue-500"
                            placeholder="Search..."
                            value={searchTerm}
                            onChange={(e) => setSearchTerm(e.target.value)}
                            onClick={(e) => e.stopPropagation()}
                        />
                    </div>
                    
                    <div className="p-2 flex flex-col gap-1">
                        <label className="flex items-center space-x-2 px-2 py-1.5 hover:bg-gray-50 rounded cursor-pointer">
                            <input 
                                type="checkbox" 
                                checked={isSearchActive ? allFilteredChecked : allChecked} 
                                ref={input => { if(input) input.indeterminate = isSearchActive ? isFilteredIndeterminate : isIndeterminate; }}
                                onChange={handleSelectAll}
                                className="rounded text-blue-600 focus:ring-blue-500 h-4 w-4 border-gray-300"
                            />
                            <span className="text-sm font-medium text-gray-900">{isSearchActive ? "Select all" : "All"}</span>
                        </label>
                        
                        {filteredOptions.map((option) => (
                            <label key={option.id} className="flex items-center space-x-2 px-2 py-1.5 hover:bg-gray-50 rounded cursor-pointer">
                                <input 
                                    type="checkbox" 
                                    checked={selectedIds.has(String(option.id))}
                                    onChange={() => handleToggle(String(option.id))}
                                    className="rounded text-blue-600 focus:ring-blue-500 h-4 w-4 border-gray-300"
                                />
                                <span className="text-sm text-gray-700">{option.name}</span>
                            </label>
                        ))}
                        {filteredOptions.length === 0 && <div className="text-center text-gray-500 text-sm py-3">No results</div>}
                    </div>
                </div>
            )}
        </div>
    );
};

const App: React.FC = () => {
    const [currentStep, setCurrentStep] = useState<AppStep>('auth');
    const [updateMethod, setUpdateMethod] = useState<UpdateMethod>('excel');
    const [credentials, setCredentials] = useState<PlandayApiCredentials | undefined>();
    const [portalName, setPortalName] = useState<string | null>(null);
    const [error, setError] = useState<string | null>(null);
    const [isLoading, setIsLoading] = useState({ types: false, template: false, submitting: false });
    const [loadingText, setLoadingText] = useState('');
    const [progress, setProgress] = useState(0);
    const abortRef = useRef<boolean>(false);

    const handleAbort = () => {
        abortRef.current = true;
    };
    
    // Step 2 State
    const [accountTypes, setAccountTypes] = useState<AccountType[]>([]);
    const [accountTypesSearch, setAccountTypesSearch] = useState('');
    
    // NEW FILTER STATE
    const [departments, setDepartments] = useState<Department[]>([]);
    const [employeeGroups, setEmployeeGroups] = useState<EmployeeGroup[]>([]);
    const [employeeTypes, setEmployeeTypes] = useState<EmployeeType[]>([]);
    
    const [usePrimaryDepartmentOnly, setUsePrimaryDepartmentOnly] = useState(false);
    const [selectedDepartmentIds, setSelectedDepartmentIds] = useState<Set<string>>(new Set());
    const [selectedEmployeeGroupIds, setSelectedEmployeeGroupIds] = useState<Set<string>>(new Set());
    const [selectedEmployeeTypeIds, setSelectedEmployeeTypeIds] = useState<Set<string>>(new Set());
    
    const [selectedTypeIds, setSelectedTypeIds] = useState<Set<number>>(new Set());
    const [validityMode, setValidityMode] = useState<ValidityMode>('current');
    const [dateRange, setDateRange] = useState({ start: '', end: '' });
    // Default to empty strings/null to force user selection
    const [balanceDate, setBalanceDate] = useState('');
    const [includeBalance, setIncludeBalance] = useState<boolean | null>(null);
    const [balanceOptionError, setBalanceOptionError] = useState<string | null>(null);
    const [includeInactive, setIncludeInactive] = useState(false);
    const [downloadDateFormat, setDownloadDateFormat] = useState<'EU' | 'US'>('EU');

    const [allEmployees, setAllEmployees] = useState<Employee[]>([]);

    // Step 3 & 4 State
    const [fetchedTemplateData, setFetchedTemplateData] = useState<TemplateDataRow[]>([]);
    const [adjustmentsToReview, setAdjustmentsToReview, undoAdjustments, redoAdjustments, pastLength, futureLength, clearAdjustmentsHistory] = useUndoableState<AdjustmentReview[]>([], 5);
    // Review Step - Sort & Select State
    const [sortConfig, setSortConfig] = useState<{ key: keyof AdjustmentReview | 'status'; direction: 'asc' | 'desc' } | null>(null);
    const [bulkCustomDate, setBulkCustomDate] = useState<string>('');
    const [searchQuery, setSearchQuery] = useState<string>('');
    
    // Review Step - Filters
    const [selectedReviewIds, setSelectedReviewIds] = useState<Set<string>>(new Set());
    const [reviewDepartmentIds, setReviewDepartmentIds] = useState<Set<string>>(new Set());
    const [reviewEmployeeGroupIds, setReviewEmployeeGroupIds] = useState<Set<string>>(new Set());
    const [reviewEmployeeTypeIds, setReviewEmployeeTypeIds] = useState<Set<string>>(new Set());
    const [reviewLeaveAccountNames, setReviewLeaveAccountNames] = useState<Set<string>>(new Set());
    
    // Review Step - Pagination State
    const [reviewRowsPerPage, setReviewRowsPerPage] = useState<number | 'All'>(50);
    const [reviewPage, setReviewPage] = useState<number>(1);
    
    // Bulk Edit State
    const [bulkEditField, setBulkEditField] = useState<'adjustment' | 'newBalance' | 'effectiveDate' | 'comment' | ''>('');
    const [bulkEditValueAdjustment, setBulkEditValueAdjustment] = useState<string>('');
    const [bulkEditValueNewBalance, setBulkEditValueNewBalance] = useState<string>('');
    const [bulkEditEffectiveDateType, setBulkEditEffectiveDateType] = useState<'custom' | 'today' | 'start_date'>('today');
    const [bulkEditValueEffectiveDate, setBulkEditValueEffectiveDate] = useState<string>('');
    const [bulkEditValueComment, setBulkEditValueComment] = useState<string>('');
    
    const [detectedColumnFormats, setDetectedColumnFormats] = useState<{effective: ColumnDetectionResult, validFrom: ColumnDetectionResult, validTo: ColumnDetectionResult} | null>(null);
    const [uploadConflicts, setUploadConflicts] = useState<{column: string, details: string[]}[] | null>(null);
    const [uploadValidityErrors, setUploadValidityErrors] = useState<{row: number, details: string[]}[] | null>(null);
    const [dateReportExample, setDateReportExample] = useState<DateReportExample | null>(null);
    const [isDragging, setIsDragging] = useState(false);
    // Store uploaded file to allow re-processing with different settings
    const [lastUploadedFile, setLastUploadedFile] = useState<File | null>(null);
    
    // Floating Tooltip State
    const [activeTooltip, setActiveTooltip] = useState<{x: number, y: number, content: React.ReactNode} | null>(null);
    
    // Step 5 State
    const [updateSummary, setUpdateSummary] = useState<AdjustmentReview[]>([]);
    const [showConfirmModal, setShowConfirmModal] = useState(false);
    const [showAbortModal, setShowAbortModal] = useState(false);
    const [showRemoveConfirmModal, setShowRemoveConfirmModal] = useState(false);
    const [hasAttemptedSubmit, setHasAttemptedSubmit] = useState(false);
    const [showBackConfirmModal, setShowBackConfirmModal] = useState(false);
    const [showUpdateBalanceModal, setShowUpdateBalanceModal] = useState(false);
    const [showStopProcessModal, setShowStopProcessModal] = useState(false);
    const [isUpdatingBalances, setIsUpdatingBalances] = useState(false);
    
    useEffect(() => {
        try {
            const savedCreds = sessionStorage.getItem('plandayCredentials');
            if (savedCreds) handleAuthSuccess(JSON.parse(savedCreds));
        } catch (e) { console.error("Failed to parse saved credentials", e); }
    }, []);
    
    const loadConfigData = async () => {
        setIsLoading(prev => ({ ...prev, types: true }));
        try {
            const types = await fetchAccountTypes();
            setAccountTypes(types); // Important to set this to exit the loading loop dependencies
        } catch (err: any) { 
            handleApiError(err); 
        }

        try {
            const info = await fetchPortalInfo();
            setPortalName(info.name);
        } catch (err: any) {
            console.error("Failed to load portal info", err);
        }

        try {
            const depts = await fetchDepartments();
            setDepartments(depts);
            setSelectedDepartmentIds(new Set(depts.map(d => String(d.id))));
        } catch (err: any) {
            setError(prev => prev ? `${prev} | ${err.message}` : err.message);
        }

        try {
            const groups = await fetchEmployeeGroups();
            setEmployeeGroups(groups);
            setSelectedEmployeeGroupIds(new Set(groups.map(g => String(g.id))));
        } catch (err: any) {
            setError(prev => prev ? `${prev} | ${err.message}` : err.message);
        }

        try {
            const eTypes = await fetchEmployeeTypes();
            setEmployeeTypes(eTypes);
            setSelectedEmployeeTypeIds(new Set(eTypes.map(t => String(t.id))));
        } catch (err: any) {
            setError(prev => prev ? `${prev} | ${err.message}` : err.message);
        }

        try {
            const employees = await fetchEmployees();
            setAllEmployees(employees);
        } catch (err: any) {
            console.error("Failed to fetch employees for review filtering", err);
        }

        setIsLoading(prev => ({ ...prev, types: false }));
    };

    useEffect(() => {
        if (currentStep === 'configure' && accountTypes.length === 0) {
            loadConfigData();
        }
    }, [currentStep, accountTypes.length]);
    
    const handleApiError = (err: any) => {
        const message = err.message || 'An unknown error occurred.';
        setError(message);
        if (message.includes('re-enter them')) {
            setCurrentStep('auth');
            setCredentials(undefined);
            // Ensure types are cleared if we are forced to re-auth
            setAccountTypes([]);
            setSelectedTypeIds(new Set());
        }
    };

    const handleAuthSuccess = (creds: PlandayApiCredentials) => {
        sessionStorage.setItem('plandayCredentials', JSON.stringify(creds));
        initializeService(creds);
        setCredentials(creds);
        
        // Reset configuration state on new auth to force re-fetch of account types
        setAccountTypes([]);
        setSelectedTypeIds(new Set());

        setCurrentStep('configure');
        setError(null);
    };

    const handleLogout = () => {
        sessionStorage.removeItem('plandayCredentials');
        setCurrentStep('auth');
        setCredentials(undefined);
        
        // Clear configuration data
        setAccountTypes([]);
        setSelectedTypeIds(new Set());
    };
    
    const resetReviewSessionState = (data: AdjustmentReview[]) => {
        setSearchQuery('');
        setSortConfig(null);
        setSelectedReviewIds(new Set());
        setBulkEditField('');
        setBulkEditValueAdjustment('');
        setBulkEditValueNewBalance('');
        setBulkEditEffectiveDateType('today');
        setBulkEditValueEffectiveDate('');
        setBulkEditValueComment('');
        
        setReviewPage(1);
        setReviewRowsPerPage(50);
        
        setReviewDepartmentIds(new Set(departments.map(d => String(d.id))));
        setReviewEmployeeGroupIds(new Set(employeeGroups.map(g => String(g.id))));
        setReviewEmployeeTypeIds(new Set(employeeTypes.map(t => String(t.id))));
        if (data.length > 0) {
            setReviewLeaveAccountNames(new Set(Array.from(new Set(data.map(r => r.accountName)))));
        } else {
            setReviewLeaveAccountNames(new Set());
        }
    };

    const handleStartOver = () => {
        setBalanceOptionError(null);
        clearAdjustmentsHistory([]);
        setUpdateSummary([]);
        setFetchedTemplateData([]);
        setDetectedColumnFormats(null);
        setUploadConflicts(null);
        setUploadValidityErrors(null);
        setDateReportExample(null);
        setLastUploadedFile(null);
        setSortConfig(null);
        setSearchQuery('');
        setHasAttemptedSubmit(false);
        setCurrentStep('configure');
        setError(null);
    };

    const handleDownloadTemplate = async () => {
        if (selectedTypeIds.size === 0) { setError("Please select at least one account type."); return; }
        if (includeBalance === null) { 
            setBalanceOptionError("This option is required to generate the template.");
            return; 
        }
        // Validation: If date is missing, error.
        if (includeBalance && !balanceDate) { 
            setBalanceOptionError("Balance Date is required when including available balance."); 
            return; 
        }
        if (validityMode === 'custom' && (!dateRange.start || !dateRange.end)) { setError("Please select both start and end dates."); return; }
        
        setIsLoading(prev => ({ ...prev, template: true }));
        setProgress(0);
        setError(null);
        setBalanceOptionError(null);
        clearAdjustmentsHistory([]);
        setSortConfig(null);
        setSearchQuery('');
        setUploadConflicts(null);
        setUploadValidityErrors(null);
        abortRef.current = false;

        try {
            // OPTIMIZATION: Reduced batch size for stability
            const FETCH_BATCH_SIZE = 5; 

            // STEP 1: Fetch all employees
            setLoadingText('Fetching all employees...');
            setProgress(5);
            let employees = await fetchEmployees();
            
            // LOCAL FILTERING OF EMPLOYEES
            if (departments.length > 0 || employeeGroups.length > 0 || employeeTypes.length > 0) {
                employees = employees.filter(emp => {
                    let matches = true;
                    
                    if (departments.length > 0 && selectedDepartmentIds.size > 0 && selectedDepartmentIds.size < departments.length) {
                        let empMatchesDept = false;
                        if (usePrimaryDepartmentOnly) {
                            if (emp.primaryDepartmentId !== undefined && emp.primaryDepartmentId !== null) {
                                empMatchesDept = selectedDepartmentIds.has(String(emp.primaryDepartmentId));
                            } else {
                                empMatchesDept = false; // Exclude if no primary department
                            }
                        } else {
                            if (Array.isArray(emp.departments)) {
                                empMatchesDept = emp.departments.some(d => selectedDepartmentIds.has(String(d)) || selectedDepartmentIds.has(String(d.department)) || selectedDepartmentIds.has(String(d.id)));
                            } else if (emp.departmentId !== undefined) {
                                empMatchesDept = selectedDepartmentIds.has(String(emp.departmentId));
                            }
                        }
                        matches = matches && empMatchesDept;
                    }
                    
                    if (employeeGroups.length > 0 && selectedEmployeeGroupIds.size > 0 && selectedEmployeeGroupIds.size < employeeGroups.length) {
                        let empMatchesGroup = false;
                        if (Array.isArray(emp.employeeGroups)) {
                            empMatchesGroup = emp.employeeGroups.some(g => selectedEmployeeGroupIds.has(String(g)) || selectedEmployeeGroupIds.has(String(g.employeeGroup)) || selectedEmployeeGroupIds.has(String(g.id)));
                        } else if (Array.isArray(emp.employeeGroupIds)) {
                            empMatchesGroup = emp.employeeGroupIds.some(id => selectedEmployeeGroupIds.has(String(id)));
                        } else if (emp.employeeGroupId !== undefined) {
                            empMatchesGroup = selectedEmployeeGroupIds.has(String(emp.employeeGroupId));
                        }
                        matches = matches && empMatchesGroup;
                    }
                    
                    if (employeeTypes.length > 0 && selectedEmployeeTypeIds.size > 0 && selectedEmployeeTypeIds.size < employeeTypes.length) {
                        let empMatchesType = false;
                        if (emp.employeeTypeId !== undefined) {
                            empMatchesType = selectedEmployeeTypeIds.has(String(emp.employeeTypeId));
                        } else if (emp.employeeType !== undefined) {
                            empMatchesType = selectedEmployeeTypeIds.has(String(emp.employeeType)) || selectedEmployeeTypeIds.has(String(emp.employeeType?.id));
                        }
                        matches = matches && empMatchesType;
                    }
                    
                    // Also check if they deselected ALL options
                    if (departments.length > 0 && selectedDepartmentIds.size === 0) matches = false;
                    if (employeeGroups.length > 0 && selectedEmployeeGroupIds.size === 0) matches = false;
                    if (employeeTypes.length > 0 && selectedEmployeeTypeIds.size === 0) matches = false;

                    return matches;
                });
            }

            setProgress(10);
            
            if (employees.length === 0) {
                 throw new Error("No employees match the selected filters.");
            }

            // Determine API query params (for initial broad filtering)
            let apiStatusParam: string | undefined = undefined;
            let apiDateFilter = undefined;
            
            // If user wants custom dates, we can try to filter at API level if "Include Inactive" is NOT checked
            if (validityMode === 'custom' && !includeInactive) {
                apiStatusParam = 'Active';
                apiDateFilter = { start: dateRange.start, end: dateRange.end };
            } else if (validityMode === 'current') {
                apiStatusParam = 'Active';
            } else if (validityMode === 'current_future') {
                 apiStatusParam = 'Active';
            }

            // STEP 2: Fetch leave accounts for all employees in parallel batches
            setLoadingText(`Fetching accounts for ${employees.length} employees...`);
            const allAccountsWithEmployeeInfo: { emp: any; accounts: any[] }[] = [];
            
            for (let i = 0; i < employees.length; i += FETCH_BATCH_SIZE) {
                if (abortRef.current) throw new Error("Process stopped by user.");
                const batchEmployees = employees.slice(i, i + FETCH_BATCH_SIZE);
                const promises = batchEmployees.map(emp =>
                    fetchLeaveAccounts(emp.id, apiDateFilter, apiStatusParam).then(accounts => ({ emp, accounts }))
                );
                const results = await Promise.all(promises);
                allAccountsWithEmployeeInfo.push(...results);
                
                // Slight delay to allow browser event loop to clear network stack
                await new Promise(resolve => setTimeout(resolve, 125));

                // Progress from 10% to 40%
                const completed = Math.min(i + FETCH_BATCH_SIZE, employees.length);
                const pct = 10 + Math.round((completed / employees.length) * 30);
                setProgress(pct);
            }

            if (abortRef.current) throw new Error("Process stopped by user.");

            // STEP 3: Filter accounts based on mode
            setLoadingText('Processing accounts...');
            setProgress(42);
            const filteredAccountsList: { accountId: number; date: string | null; emp: any; acc: any }[] = [];
            const today = getTodayYYYYMMDD();
            
            allAccountsWithEmployeeInfo.forEach(({ emp, accounts }) => {
                const filtered = accounts.filter(acc => {
                    // Type check
                    if (!selectedTypeIds.has(acc.typeId)) return false;
                    
                    const isFlextime = accountTypes.find(t => t.id === acc.typeId)?.absenceType === 'Flextime';
                    
                    if (!isFlextime) {
                        // CRITICAL REQUIREMENT: Always exclude accounts with NO valid period
                        if (!acc.validityPeriod || !acc.validityPeriod.start) return false;
                        
                        const start = acc.validityPeriod.start.split('T')[0];
                        const end = acc.validityPeriod.end ? acc.validityPeriod.end.split('T')[0] : null;

                        if (validityMode === 'current') {
                            // Current: Today must be >= start AND (end is null or Today <= end)
                            // This implies the account is currently active today.
                            if (start > today) return false; // Starts in future
                            if (end && end < today) return false; // Already ended
                            return true;
                        }
                        
                        if (validityMode === 'current_future') {
                            // Current + Upcoming: End date must be >= Today (or null)
                            // Include if it starts today, started in past (but not ended), or starts in future.
                            if (end && end < today) return false; // Already ended
                            return true;
                        }

                        if (validityMode === 'custom') {
                             if (!includeInactive) {
                                 // If "Include Inactive" is OFF, we also check if it ended before the requested range
                                 // But basic date overlap logic:
                                 // Acc Start must be <= Range End
                                 // Acc End (if exists) must be >= Range Start
                                 if (start > dateRange.end) return false;
                                 if (end && end < dateRange.start) return false;
                             } else {
                                 // "Include Inactive" ON: Just check if it overlaps the requested period at all
                                 // Logic is same as above but conceptually we are allowing things that might be expired relative to today,
                                 // as long as they are valid within the custom range window.
                                 if (start > dateRange.end) return false;
                                 if (end && end < dateRange.start) return false;
                             }
                             return true;
                        }
                    }
                    
                    return true;
                });

                filtered.forEach(acc => {
                    let effectiveBalanceDate: string | null = null;
                    
                    if (includeBalance) {
                        // Standard Mode: Use selected date, capped by account end date
                        effectiveBalanceDate = balanceDate;
                        if (acc.validityPeriod?.end) {
                            const accountEndDate = acc.validityPeriod.end.split('T')[0];
                            if (accountEndDate < effectiveBalanceDate) {
                                effectiveBalanceDate = accountEndDate;
                            }
                        }
                    }

                    filteredAccountsList.push({
                        accountId: acc.id,
                        date: effectiveBalanceDate,
                        emp,
                        acc,
                    });
                });
            });
            
            setProgress(45);

            if (filteredAccountsList.length === 0) {
                setFetchedTemplateData([]);
                throw new Error("No leave accounts found matching the selected criteria.");
            }

            const allTemplateRows: TemplateDataRow[] = [];

            // STEP 4: Fetch balances OR Smart Sample Units
            if (includeBalance) {
                setLoadingText(`Fetching ${filteredAccountsList.length} account balances...`);
                // OPTIMIZATION: Reduced batch size for stability
                for (let i = 0; i < filteredAccountsList.length; i += FETCH_BATCH_SIZE) {
                    if (abortRef.current) throw new Error("Process stopped by user.");
                    const batchJobs = filteredAccountsList.slice(i, i + FETCH_BATCH_SIZE);
                    
                    const balancePromises = batchJobs.map(job => 
                        // job.date is checked by validation to be present if includeBalance is true
                        fetchAccountBalance(job.accountId, job.date!)
                            .catch(err => {
                                console.error(`Failed to fetch balance for account ${job.accountId}`, err);
                                return { balance: 0, unit: 'N/A (Error)' };
                            })
                    );
                    
                    const balanceResults = await Promise.all(balancePromises);

                    const rowsForBatch = balanceResults.map((balance, index) => {
                        const { emp, acc, date } = batchJobs[index];
                        return {
                            employeeId: emp.id,
                            salaryIdentifier: emp.salaryIdentifier || null,
                            employeeName: `${emp.firstName} ${emp.lastName}`,
                            accountId: acc.id,
                            accountName: acc.name,
                            accountTypeCategory: getAccountCategory(acc.typeId, accountTypes),
                            validFrom: formatDateForDisplay(acc.validityPeriod?.start?.split('T')[0], downloadDateFormat),
                            validTo: formatDateForDisplay(acc.validityPeriod?.end?.split('T')[0], downloadDateFormat),
                            balanceDate: formatDateForDisplay(date, downloadDateFormat),
                            availableBalance: balance.balance,
                            balanceUnit: balance.unit,
                        };
                    });
                    allTemplateRows.push(...rowsForBatch);

                    // Slight delay to allow browser event loop to clear network stack
                    await new Promise(resolve => setTimeout(resolve, 125));
                    
                    // Progress from 45% to 90%
                    const completed = Math.min(i + FETCH_BATCH_SIZE, filteredAccountsList.length);
                    const pct = 45 + Math.round((completed / filteredAccountsList.length) * 45);
                    setProgress(pct);
                }
            } else {
                setLoadingText('Detecting unit types (sampling)...');
                
                // Smart Sampling: Group by Account Type ID
                const distinctTypeIds = Array.from(new Set(filteredAccountsList.map(item => item.acc.typeId)));
                const unitLookup = new Map<number, string>();

                // For each type, fetch ONE balance to determine the unit (Days/Hours)
                // This is much faster than fetching all balances but ensures accurate unit reporting
                const samplePromises = distinctTypeIds.map(async (typeId) => {
                    // Check metadata first (if available)
                    const meta = accountTypes.find(t => t.id === typeId);
                    if (meta && meta.unit && meta.unit !== 'N/A') {
                        unitLookup.set(typeId, meta.unit);
                        return;
                    }

                    // Fallback: Sample an account
                    const sampleAccount = filteredAccountsList.find(item => item.acc.typeId === typeId);
                    if (!sampleAccount) {
                        unitLookup.set(typeId, 'N/A');
                        return;
                    }

                    // Determine valid date for sample
                    let sampleDate = getTodayYYYYMMDD();
                    if (sampleAccount.acc.validityPeriod?.end) {
                        const end = sampleAccount.acc.validityPeriod.end.split('T')[0];
                        if (end < sampleDate) sampleDate = end;
                    }

                    try {
                        const result = await fetchAccountBalance(sampleAccount.accountId, sampleDate);
                        unitLookup.set(typeId, result.unit);
                    } catch (e) {
                        console.warn(`Unit sampling failed for type ${typeId}`, e);
                        unitLookup.set(typeId, 'N/A');
                    }
                });

                await Promise.all(samplePromises);

                if (abortRef.current) throw new Error("Process stopped by user.");

                setLoadingText('Generating rows...');

                 // Simply map using lookup
                 filteredAccountsList.forEach(({emp, acc}) => {
                     allTemplateRows.push({
                        employeeId: emp.id,
                        salaryIdentifier: emp.salaryIdentifier || null,
                        employeeName: `${emp.firstName} ${emp.lastName}`,
                        accountId: acc.id,
                        accountName: acc.name,
                        accountTypeCategory: getAccountCategory(acc.typeId, accountTypes),
                        validFrom: formatDateForDisplay(acc.validityPeriod?.start?.split('T')[0], downloadDateFormat),
                        validTo: formatDateForDisplay(acc.validityPeriod?.end?.split('T')[0], downloadDateFormat),
                        balanceDate: 'N/A', // Will be ignored in export logic
                        availableBalance: 0, // Placeholder
                        balanceUnit: unitLookup.get(acc.typeId) || 'N/A'
                     });
                 });
                 setProgress(90);
            }
            
            // STEP 5: Pivot to Editor or Create Excel
            if (updateMethod === 'editor') {
                setLoadingText('Preparing editor table...');
                setProgress(98);
                const editorRows: AdjustmentReview[] = allTemplateRows.map(row => {
                    const adjObj: AdjustmentReview = {
                        id: crypto.randomUUID(),
                        employeeId: row.employeeId,
                        salaryIdentifier: row.salaryIdentifier,
                        employeeName: row.employeeName,
                        accountId: row.accountId,
                        accountName: row.accountName,
                        accountTypeCategory: row.accountTypeCategory,
                        validFrom: row.validFrom !== 'N/A' ? parseDateToIso(row.validFrom, downloadDateFormat) : null,
                        validTo: row.validTo !== 'N/A' ? parseDateToIso(row.validTo, downloadDateFormat) : null,
                        balanceDate: row.balanceDate !== 'N/A' ? parseDateToIso(row.balanceDate, downloadDateFormat) : null,
                        availableBalance: includeBalance ? row.availableBalance : 0,
                        unit: row.balanceUnit !== 'N/A' ? row.balanceUnit : undefined,
                        adjustment: 0,
                        newBalance: includeBalance ? row.availableBalance : 0,
                        comment: "",
                        effectiveDate: "",
                        status: "pending"
                    };
                    return validateRow(adjObj);
                });
                
                clearAdjustmentsHistory(editorRows);
                resetReviewSessionState(editorRows);
                setProgress(100);
                setHasAttemptedSubmit(false);
                setCurrentStep('review');
                setIsLoading(prev => ({ ...prev, template: false }));
                return;
            }

            setLoadingText('Finalizing Excel file...');
            setProgress(95);
            setFetchedTemplateData(allTemplateRows);

            // Dynamically build headers
            const headers = [
                "Planday ID", 
                "Account ID", 
                "Salary Identifier", 
                "Full Name", 
                "Leave Account Name", 
                "Account Type",
                "Valid From", 
                "Valid To"
            ];
            
            if (includeBalance) {
                headers.push("Balance Date");
            }
            
            headers.push("Unit Type (Days or Hours)", "Available Balance");

            if (includeBalance) {
                headers.push("New Balance");
            }
            
            headers.push("Adjustment", "Effective Date", "Comment");
            
            const formatText = downloadDateFormat === 'US' ? 'MM/DD/YYYY' : 'DD/MM/YYYY';

            const headerNotes = [
                "System ID for the employee. DO NOT EDIT.",
                "System ID for the account. DO NOT EDIT.",
                "Payroll identifier. DO NOT EDIT.",
                "Employee Name. DO NOT EDIT.",
                "The specific leave account. DO NOT EDIT.",
                "Account Type (FLEX/TOIL, Fixed, Accrued). DO NOT EDIT.",
                `Account start date. Format: ${formatText}. DO NOT EDIT.`,
                `Account end date. Format: ${formatText}. DO NOT EDIT.`
            ];
            
            if (includeBalance) {
                headerNotes.push(`The date used to check the 'Available Balance'. Format: ${formatText}. DO NOT EDIT.`);
            }
            
            headerNotes.push(
                "Unit of measurement. DO NOT EDIT.",
                "The balance in Planday as of the 'Balance Date'. DO NOT EDIT."
            );

            // New Balance Note text update
            if (includeBalance) {
                headerNotes.push("OPTIONAL HELPER: Enter target final balance here to auto-calculate 'Adjustment'. Leave blank if entering Adjustment directly. Note, only the 'Adjustment' values are sent for leave balance updates.");
            }

            // Comment Note text update
            headerNotes.push(
                "REQUIRED: The value sent to Planday. " + (includeBalance ? "Auto-calculated if New Balance is filled, OR enter value directly." : "Enter adjustment value directly."),
                `REQUIRED: Date of adjustment. Format: ${formatText} or YYYY-MM-DD. Can be left blank and updated in app before upload.`,
                "Optional: Additional reason for this adjustment. Note, the following text is ALWAYS SENT as a comment, whether a comment is entered or not: API BULK UPDATE."
            );

            const dataForSheet = allTemplateRows.map((row, index) => {
                const base: any = {
                    "Planday ID": row.employeeId, 
                    "Account ID": row.accountId,
                    "Salary Identifier": row.salaryIdentifier,
                    "Full Name": row.employeeName, 
                    "Leave Account Name": row.accountName,
                    "Account Type": row.accountTypeCategory,
                    "Valid From": row.validFrom, 
                    "Valid To": row.validTo
                };
                
                if (includeBalance) {
                    base["Balance Date"] = row.balanceDate;
                }
                
                base["Unit Type (Days or Hours)"] = row.balanceUnit;
                base["Available Balance"] = includeBalance ? row.availableBalance : "Not retrieved";
                
                const extra: any = {};
                
                if (includeBalance) {
                    extra["New Balance"] = "";
                    // Col A-H (8 cols)
                    // I: Balance Date
                    // J: Unit Type
                    // K: Available Balance
                    // L: New Balance
                    // M: Adjustment
                    // Adjustment formula: L - K
                    extra["Adjustment"] = { f: `IF(L${index + 2}<>"", L${index + 2}-K${index + 2}, "")` };
                } else {
                    // Manual entry only
                    extra["Adjustment"] = "";
                }

                extra["Effective Date"] = "";
                extra["Comment"] = "";

                return { ...base, ...extra };
            });

            const ws = XLSX.utils.json_to_sheet(dataForSheet, { header: headers });
            
            // --- STYLING START ---
            const range = XLSX.utils.decode_range(ws['!ref']);
            // Colors: Removed Available Balance, Added New Balance (if exists)
            const colsToColor = ["New Balance", "Adjustment", "Effective Date", "Comment"];
            const colIndicesToColor = colsToColor.map(c => headers.indexOf(c)).filter(i => i !== -1);
            
            // Standard border style
            const borderStyle = {
                top: { style: "thin", color: { rgb: "d9d9d9" } },
                bottom: { style: "thin", color: { rgb: "d9d9d9" } },
                left: { style: "thin", color: { rgb: "d9d9d9" } },
                right: { style: "thin", color: { rgb: "d9d9d9" } }
            };

            for (let R = range.s.r; R <= range.e.r; ++R) {
                for (let C = range.s.c; C <= range.e.c; ++C) {
                    const cellAddress = XLSX.utils.encode_cell({ c: C, r: R });
                    if (!ws[cellAddress]) continue;

                    // Ensure style object exists
                    if (!ws[cellAddress].s) ws[cellAddress].s = {};
                    
                    // Apply Border to ALL cells
                    ws[cellAddress].s.border = borderStyle;

                    // Header Row (Row 0): Dark Blue background (#162C34), White text (#FFFFFF), Bold
                    if (R === 0) {
                        ws[cellAddress].s.fill = { fgColor: { rgb: "162C34" } };
                        ws[cellAddress].s.font = { 
                            bold: true, 
                            color: { rgb: "FFFFFF" },
                            name: "Calibri",
                            sz: 11
                        };

                        // ADD CELL COMMENT TO NEW BALANCE HEADER
                        const colName = headers[C];
                        if (colName === "New Balance") {
                            // Using standard xlsx comment structure
                            if (!ws[cellAddress].c) ws[cellAddress].c = [];
                            ws[cellAddress].c.push({
                                t: "New Balance: This is an optional helper field. Enter target final balance here to auto-calculate 'Adjustment'. Leave blank if entering Adjustment directly. Note, only the 'Adjustment' values are sent for leave balance updates.\n\nYou can read more instructions and tips on how to fill out this template by navigating to the second sheet in this file named 'Instructions'.",
                                h: true // hidden by default, visible on hover
                            });
                        }
                    } else {
                        // Data Rows
                        // Background Color for specific columns (Light Blue)
                        if (colIndicesToColor.includes(C)) {
                            ws[cellAddress].s.fill = { fgColor: { rgb: "DCE6F1" } };
                        }
                    }
                }
            }
            // --- STYLING END ---

            // Prepare Instructions Sheet Data
            const focusNote = { "Column Name": "IMPORTANT", "Description": "Columns highlighted in LIGHT BLUE (e.g. New Balance, Adjustment) are intended for your input. Please do NOT edit the other columns as they are required for system matching." };
            
            const proTipRows = [
                { "Column Name": "PRO TIP: IMPORTING FROM OTHER FILES", "Description": "If you have balances in another spreadsheet, you can copy them into a new sheet in this file and use this formula." },
                { "Column Name": "Formula", "Description": "=INDEX('Sheet1'!$B:$B, MATCH(C2, 'Sheet1'!$A:$A, 0))" },
                { "Column Name": "How to use it", "Description": "• 'Sheet1': The sheet with the employee details you copied. Remember the single quotes, and change the name if the sheet name is named differently.\n• C2: The identifier in your Leave Balance sheet (e.g., Salary Identifier). Change to D2, or E2, or something else, to select a different identifier column.\n• $A: The column in the copied employee details sheet (e.g. 'Sheet1') that holds that identifier.\n• $B: The column in the copied employee details sheet (e.g. 'Sheet1') with the value you want returned." },
                { "Column Name": "Where to put it", "Description": "• In your Leave Balance sheet, click the first cell under the target column header (e.g., New Balance or Adjustment).\n• Enter the formula and press Enter.\n• If it is wrong or shows an error like #N/A, check the formula for errors. It might also be that the employee is not found in the copied employee details sheet.\n• If it looks correct, drag the fill handle down to apply to all cells within that column." },
                { "Column Name": "Important", "Description": "When using the formula: to avoid sending #N/A for the leave balance update, delete these values and leave it blank or manually input a specific value." }
            ];

            const columnDescRows = headers.map((header, index) => ({
                "Column Name": header,
                "Description": headerNotes[index]
            }));

            const instructionData = [
                focusNote,
                ...proTipRows,
                { "Column Name": "", "Description": "" }, // Spacer
                { "Column Name": "COLUMN DESCRIPTIONS", "Description": "" },
                ...columnDescRows
            ];
            
            const wsInstructions = XLSX.utils.json_to_sheet(instructionData);
            
            wsInstructions['!cols'] = [
                { wch: 40 },
                { wch: 100 }
            ];

            // --- INSTRUCTIONS STYLING START ---
            const instRange = XLSX.utils.decode_range(wsInstructions['!ref']);
            // Highlight Rows: Focus Note (Row 1) and Pro Tips (Row 2 to 2+length-1)
            // Header is Row 0.
            const highlightStart = 1;
            const highlightEnd = 1 + proTipRows.length; // Focus Note + Pro Tip rows

            for (let R = instRange.s.r; R <= instRange.e.r; ++R) {
                for (let C = instRange.s.c; C <= instRange.e.c; ++C) {
                    const cellAddress = XLSX.utils.encode_cell({ c: C, r: R });
                    if (!wsInstructions[cellAddress]) continue;
                    if (!wsInstructions[cellAddress].s) wsInstructions[cellAddress].s = {};
                    
                    // Borders
                    wsInstructions[cellAddress].s.border = borderStyle;

                    // Header
                    if (R === 0) {
                        wsInstructions[cellAddress].s.fill = { fgColor: { rgb: "162C34" } };
                        wsInstructions[cellAddress].s.font = { 
                            bold: true, 
                            color: { rgb: "FFFFFF" },
                            name: "Calibri",
                            sz: 11
                        };
                    } else if (R >= highlightStart && R <= highlightEnd) {
                        // Highlight Focus Note and Pro Tips with Yellow
                        wsInstructions[cellAddress].s.fill = { fgColor: { rgb: "FFFFA9" } };
                        // Wrap text for description column
                        if (C === 1) {
                             wsInstructions[cellAddress].s.alignment = { wrapText: true, vertical: "top" };
                        }
                    }
                }
            }
            // --- INSTRUCTIONS STYLING END ---

            const wscols = headers.map(h => ({ wch: h.length + 5 }));
            ws['!cols'] = wscols;
            ws['!freeze'] = { xSplit: 4, ySplit: 1 };

            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "Leave Balances");
            XLSX.utils.book_append_sheet(wb, wsInstructions, "Instructions");
            
            XLSX.writeFile(wb, "Planday_Leave_Balance_Template.xlsx");
            
            setProgress(100);

        } catch (err: any) { 
            if (err.message === "Process stopped by user.") {
                setError(null);
            } else {
                handleApiError(err);
            }
        }
        finally {
            setIsLoading(prev => ({ ...prev, template: false }));
            setLoadingText('');
        }
    };
    
    // --- Validation Helper ---
    // Checks if effectiveDate matches the validFrom/validTo constraints in the row
    const validateRow = (row: AdjustmentReview): AdjustmentReview => {
        const item = { ...row };
        
        // Reset previous validation error
        if (item.isValidationError) {
            item.status = 'pending';
            item.error = undefined;
            item.isValidationError = false;
        }

        if (!item.validFrom && !item.validTo) return item; // No constraints found, skip
        
        const eff = item.effectiveDate;
        if (!eff) return item; // Do not check validity yet if date is empty

        const validFrom = item.validFrom;
        const validTo = item.validTo;

        // String comparison works for ISO dates
        if (validFrom && eff < validFrom) {
            item.status = 'error';
            // Use downloadDateFormat for displaying errors
            item.error = `Effective Date (${formatDateForDisplay(eff, downloadDateFormat)}) is before the Account Start Date (${formatDateForDisplay(validFrom, downloadDateFormat)}).`;
            item.isValidationError = true;
            return item;
        }

        if (validTo && eff > validTo) {
            item.status = 'error';
            item.error = `Effective Date (${formatDateForDisplay(eff, downloadDateFormat)}) is after the Account End Date (${formatDateForDisplay(validTo, downloadDateFormat)}).`;
            item.isValidationError = true;
            return item;
        }

        return item;
    };

    const processUploadedFile = (file: File, formatOverride?: 'EU' | 'US') => {
        setError(null);
        setDetectedColumnFormats(null);
        setUploadConflicts(null);
        setUploadValidityErrors(null);
        setDateReportExample(null);
        setLastUploadedFile(file);
        setSortConfig(null);
        setSearchQuery('');
        
        // Use override if provided, otherwise state
        const detectionFormat = formatOverride || downloadDateFormat;
        
        const reader = new FileReader();
        reader.onload = (event) => {
            try {
                const data = new Uint8Array(event.target?.result as ArrayBuffer);
                // Read with cellDates: true to parse Date objects.
                const workbook = XLSX.read(data, { type: 'array', cellDates: true });
                const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                
                // Read JSON first to detect format. Use raw: true for dates.
                const json: any[] = XLSX.utils.sheet_to_json(worksheet, { raw: true });
                
                // PHASE 1: Detect Format Per Column
                // Now uses row-aware detection to catch conflicts
                const detectedEffective = detectColumnFormat(json, 'Effective Date', detectionFormat);
                const detectedValidFrom = detectColumnFormat(json, 'Valid From', detectionFormat);
                const detectedValidTo = detectColumnFormat(json, 'Valid To', detectionFormat);

                // Smart Inheritance Logic: 
                // 1. Cross-reference Valid From <-> Valid To
                let validFromToUse = { ...detectedValidFrom };
                let validToToUse = { ...detectedValidTo };

                if (validFromToUse.source === 'fallback' && validToToUse.source === 'detected') {
                    validFromToUse.format = validToToUse.format;
                    validFromToUse.source = 'inherited' as any;
                } else if (validToToUse.source === 'fallback' && validFromToUse.source === 'detected') {
                    validToToUse.format = validFromToUse.format;
                    validToToUse.source = 'inherited' as any;
                }

                // 2. Effective Date inheritance
                // If Effective Date is empty/ambiguous (fallback), adopt the format from Valid From/To if detected/inherited.
                let effectiveToUse = { ...detectedEffective };
                if (effectiveToUse.source === 'fallback') {
                    if (validFromToUse.source === 'detected' || validFromToUse.source === 'inherited') {
                        effectiveToUse.format = validFromToUse.format;
                        effectiveToUse.source = 'inherited' as any; // Using 'inherited' to suppress warnings
                    } else if (validToToUse.source === 'detected' || validToToUse.source === 'inherited') {
                        effectiveToUse.format = validToToUse.format;
                        effectiveToUse.source = 'inherited' as any;
                    }
                }

                const conflicts: {column: string, details: string[]}[] = [];
                if (effectiveToUse.hasConflict) conflicts.push({ column: 'Effective Date', details: effectiveToUse.conflictDetails });
                if (validFromToUse.hasConflict) conflicts.push({ column: 'Valid From', details: validFromToUse.conflictDetails });
                if (validToToUse.hasConflict) conflicts.push({ column: 'Valid To', details: validToToUse.conflictDetails });
                
                // Strict Consistency Check: Valid From vs Valid To
                // If both are 'detected' (unambiguous) but differ in format, blocking error.
                if (validFromToUse.source === 'detected' && validToToUse.source === 'detected' && validFromToUse.format !== validToToUse.format) {
                     conflicts.push({
                         column: 'Validity Period Mismatch',
                         details: [
                             `Valid From column detected as ${validFromToUse.format} format (e.g. Row ${validFromToUse.exampleRowIndex! + 2}).`,
                             `Valid To column detected as ${validToToUse.format} format (e.g. Row ${validToToUse.exampleRowIndex! + 2}).`,
                             `These columns must use the same date format.`
                         ]
                     });
                }

                if (conflicts.length > 0) {
                    setUploadConflicts(conflicts);
                    // Do not process rows if conflicts exist
                    return;
                }

                const detectedBalanceDate = detectColumnFormat(json, 'Balance Date', detectionFormat);
                
                // Track unique Balance Dates
                const rawBalanceDateValues = new Set<string>();
                json.forEach(row => {
                    const bd = row['Balance Date'];
                    if (bd !== undefined && bd !== null && String(bd).trim() !== '' && String(bd).trim().toUpperCase() !== 'N/A') {
                        const parsedBd = parseDateToIso(bd, detectedBalanceDate.format);
                        if (parsedBd) rawBalanceDateValues.add(parsedBd);
                    }
                });

                if (rawBalanceDateValues.size > 1) {
                    setIsLoading(prev => ({ ...prev, pagedData: false }));
                    setError("Multiple different 'Balance Date' values detected in the uploaded file. Only a single date chosen for all accounts can be uploaded in the same file. Please align them to a single Date before uploading.");
                    return;
                }

                // Check for Missing OR Invalid Dates
                const validationErrors: {row: number, details: string[]}[] = [];
                json.forEach((row, idx) => {
                    // Only check rows that appear to be data rows (have Account ID or Name)
                    if (!row['Account ID'] && !row['Leave Account Name']) return;

                    const issues: string[] = [];
                    
                    // 1. Check Valid From
                    const rawValidFrom = row['Valid From'];
                    if (rawValidFrom === undefined || rawValidFrom === null || String(rawValidFrom).trim() === '') {
                        issues.push("Missing 'Valid From' date (enter 'N/A' if none)");
                    } else {
                        const strVal = String(rawValidFrom).trim();
                        if (strVal.toUpperCase() !== 'N/A') {
                             const parsed = parseDateToIso(rawValidFrom, validFromToUse.format);
                             if (!parsed) issues.push(`Invalid 'Valid From' date format: "${String(rawValidFrom)}"`);
                        }
                    }

                    // 2. Check Valid To
                    const rawValidTo = row['Valid To'];
                    if (rawValidTo === undefined || rawValidTo === null || String(rawValidTo).trim() === '') {
                        issues.push("Missing 'Valid To' date (enter 'N/A' if none)");
                    } else {
                        const strVal = String(rawValidTo).trim();
                        if (strVal.toUpperCase() !== 'N/A') {
                             const parsed = parseDateToIso(rawValidTo, validToToUse.format);
                             if (!parsed) issues.push(`Invalid 'Valid To' date format: "${String(rawValidTo)}"`);
                        }
                    }

                    // 3. Check Effective Date (Optional, but if present must be valid)
                    const rawEffective = row['Effective Date'];
                    if (rawEffective !== undefined && rawEffective !== null && String(rawEffective).trim() !== '') {
                        const parsed = parseDateToIso(rawEffective, effectiveToUse.format);
                        if (!parsed) issues.push(`Invalid 'Effective Date' format: "${String(rawEffective)}"`);
                    }
                    
                    if (issues.length > 0) {
                        validationErrors.push({ row: idx + 2, details: issues });
                    }
                });

                if (validationErrors.length > 0) {
                    setUploadValidityErrors(validationErrors);
                    return;
                }

                setDetectedColumnFormats({
                    effective: effectiveToUse,
                    validFrom: validFromToUse,
                    validTo: validToToUse
                });
                
                // Set Example Row for Report if needed
                // Priority: Fallback reason (ambiguous), then Mismatch reason (format != preference)
                let example: DateReportExample | null = null;
                const culprit = effectiveToUse.source === 'fallback' ? effectiveToUse :
                                validFromToUse.source === 'fallback' ? validFromToUse :
                                validToToUse.source === 'fallback' ? validToToUse :
                                effectiveToUse.format !== detectionFormat ? effectiveToUse :
                                validFromToUse.format !== detectionFormat ? validFromToUse :
                                validToToUse.format !== detectionFormat ? validToToUse : null;
                
                if (culprit && culprit.exampleRowIndex !== undefined) {
                    const rowData = json[culprit.exampleRowIndex];
                    if (rowData) {
                        // For the example, we calculate the ISO interpretation to show the user
                        const rawKey = culprit === effectiveToUse ? 'Effective Date' : culprit === validFromToUse ? 'Valid From' : 'Valid To';
                        const rawVal = rowData[rawKey] || '';
                        
                        // FIX: Parse using the DETECTED (Culprit) format to get a valid date object.
                        // This avoids "N/A" when detectionFormat (Preference) mismatches the file content.
                        const parsedIso = parseDateToIso(rawVal, culprit.format);
                        
                        // Display using the PREFERENCE format (to show what user wants vs what file has)
                        const displayed = formatDateForDisplay(parsedIso, detectionFormat);

                        example = {
                            rowNumber: culprit.exampleRowIndex + 2,
                            employee: rowData['Full Name'] || 'Unknown',
                            account: rowData['Leave Account Name'] || 'Unknown',
                            columnName: rawKey,
                            rawValue: String(rawVal),
                            convertedValue: displayed,
                            detectedFormat: detectionFormat
                        };
                        setDateReportExample(example);
                    }
                }

                // PHASE 2: Parse Rows
                const reviews = json.map((row, idx): AdjustmentReview | null => {
                    const adjustment = parseFloat(row['Adjustment']);
                    if (isNaN(adjustment)) return null;

                    let accountId: number | null = null;
                    let accountName = row['Leave Account Name'];
                    let employeeName = row['Full Name'];
                    let availableBalance = row['Available Balance'];
                    let unit = row['Unit Type (Days or Hours)'];
                    
                    let balanceDate = undefined;
                    const rawBalanceDate = row['Balance Date'];
                    if (rawBalanceDate !== undefined && rawBalanceDate !== null && String(rawBalanceDate).trim() !== '' && String(rawBalanceDate).trim().toUpperCase() !== 'N/A') {
                        balanceDate = parseDateToIso(rawBalanceDate, detectedBalanceDate.format);
                    }

                    let newBalance = row['New Balance'];
                    
                    if (availableBalance === 'Not retrieved') {
                        availableBalance = 'N/A';
                    }

                    if (row['Account ID']) {
                        accountId = parseInt(row['Account ID'], 10);
                        // If it came with account ID but no specific balanceDate parsed from row, check if originalData matches
                        if (!balanceDate) {
                            const originalData = fetchedTemplateData.find(d => d.accountId === accountId);
                            if (originalData) balanceDate = originalData.balanceDate;
                        }
                    } else {
                        const originalData = fetchedTemplateData.find(d => d.employeeId === row['Planday ID'] && d.accountName === row['Leave Account Name']);
                        if (originalData) {
                            accountId = originalData.accountId;
                            if (!availableBalance && availableBalance !== 0) availableBalance = originalData.availableBalance;
                            if (!unit) unit = originalData.balanceUnit;
                            if (!balanceDate) balanceDate = originalData.balanceDate;
                        }
                    }

                    if (accountId && (newBalance === undefined || newBalance === '') && typeof availableBalance === 'number' && !isNaN(adjustment)) {
                        newBalance = availableBalance + adjustment;
                    }

                    if (!accountId) return null;
                    
                    // Parse Dates using Specific Formats detected for each column
                    const rawDate = row['Effective Date'];
                    const rawValidFrom = row['Valid From'];
                    const rawValidTo = row['Valid To'];

                    let effectiveDate: string | null = null;
                    let validFrom: string | null = null;
                    let validTo: string | null = null;

                    // 1. Effective Date
                    if (rawDate === undefined || rawDate === null || String(rawDate).trim() === '') {
                        effectiveDate = getTodayYYYYMMDD();
                    } else {
                        // Apply detected format strictly for entire column (using inherited format if applicable)
                        effectiveDate = parseDateToIso(rawDate, effectiveToUse.format);
                    }
                    
                    // 2. Validity Constraints (from Excel)
                    if (rawValidFrom && String(rawValidFrom).trim().toUpperCase() !== 'N/A') {
                        validFrom = parseDateToIso(rawValidFrom, validFromToUse.format);
                    }
                    if (rawValidTo && String(rawValidTo).trim().toUpperCase() !== 'N/A') {
                         validTo = parseDateToIso(rawValidTo, validToToUse.format);
                    }
                    
                    if (!effectiveDate) return null; // Date was provided but Invalid

                    let item: AdjustmentReview = {
                        // Generate unique ID for UI tracking
                        id: `adj-${idx}-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
                        accountId: accountId,
                        employeeId: row['Planday ID'] ? parseInt(row['Planday ID'], 10) : undefined,
                        employeeName: employeeName || 'Unknown', 
                        salaryIdentifier: row['Salary Identifier'] || null,
                        accountName: accountName || 'Unknown',
                        accountTypeCategory: (row['Account Type'] as 'FLEX/TOIL' | 'Fixed' | 'Accrued' | 'Unknown') || 'Unknown',
                        balanceDate: balanceDate,
                        availableBalance: availableBalance, 
                        newBalance: newBalance, 
                        adjustment: adjustment,
                        unit: unit || 'N/A', // Store unit
                        effectiveDate: effectiveDate, // This is now always YYYY-MM-DD
                        validFrom: validFrom,
                        validTo: validTo,
                        comment: row['Comment'] || '', 
                        status: 'pending',
                    };

                    // Run Validation Logic Immediately
                    item = validateRow(item);

                    return item;
                }).filter((item): item is AdjustmentReview => item !== null);
                
                if (reviews.length === 0) {
                    setError("No valid adjustments found. Ensure 'Adjustment' column is filled. If you refreshed the page, please download the template again to ensure it contains Account IDs.");
                } else {
                    clearAdjustmentsHistory(reviews);
                    resetReviewSessionState(reviews);
                    setHasAttemptedSubmit(false);
                    setCurrentStep('review');
                }
            } catch (err:any) { 
                console.error(err);
                handleApiError(new Error("Failed to parse the uploaded file. Please check the date formats and try again.")); 
            }
        };
        reader.readAsArrayBuffer(file);
    };

    const handleSwitchDateFormat = () => {
        if (!lastUploadedFile) return;
        const newFormat = downloadDateFormat === 'EU' ? 'US' : 'EU';
        setDownloadDateFormat(newFormat);
        // Re-process with new format override immediately
        processUploadedFile(lastUploadedFile, newFormat);
    };

    const handleFileInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
        const fileInput = e.target;
        const file = fileInput.files?.[0];
        if (file) {
            processUploadedFile(file);
        }
        fileInput.value = '';
    };

    const handleDragOver = (e: React.DragEvent<HTMLDivElement>) => {
        e.preventDefault();
        setIsDragging(true);
    };

    const handleDragLeave = (e: React.DragEvent<HTMLDivElement>) => {
        e.preventDefault();
        setIsDragging(false);
    };

    const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
        e.preventDefault();
        setIsDragging(false);
        const file = e.dataTransfer.files?.[0];
        if (file) {
            processUploadedFile(file);
        }
    };

    const handleUpdateEffectiveDate = (id: string, newDate: string) => {
        setAdjustmentsToReview(prev => prev.map(item => {
            if (item.id === id) {
                const updated = { ...item, effectiveDate: newDate };
                return validateRow(updated);
            }
            return item;
        }));
    };

    const handleUpdateAdjustment = (id: string, newAdjustment: number) => {
        setAdjustmentsToReview(prev => prev.map(item => {
            if (item.id === id) {
                let updatedNewBalance = item.newBalance;
                if (typeof item.availableBalance === 'number' && !isNaN(item.availableBalance)) {
                    updatedNewBalance = Number((item.availableBalance + newAdjustment).toFixed(2));
                }
                return validateRow({ ...item, adjustment: newAdjustment, newBalance: updatedNewBalance });
            }
            return item;
        }));
    };

    const handleUpdateNewBalance = (id: string, finalNewBalance: number) => {
        setAdjustmentsToReview(prev => prev.map(item => {
            if (item.id === id) {
                let derivedAdjustment = item.adjustment;
                if (typeof item.availableBalance === 'number' && !isNaN(item.availableBalance)) {
                    // newBalance = availableBalance + adjustment  => adjustment = newBalance - availableBalance
                    derivedAdjustment = Number((finalNewBalance - item.availableBalance).toFixed(2));
                }
                return validateRow({ ...item, adjustment: derivedAdjustment, newBalance: finalNewBalance });
            }
            return item;
        }));
    };

    const handleUpdateComment = (id: string, newComment: string) => {
        setAdjustmentsToReview(prev => prev.map(item => {
            if (item.id === id) {
                return { ...item, comment: newComment };
            }
            return item;
        }));
    };

    const handleUpdateBalances = async () => {
        setIsUpdatingBalances(true);
        setShowUpdateBalanceModal(false);
        setLoadingText('Updating available balances...');
        setProgress(0);
        abortRef.current = false;
        
        try {
            const updatedAdjustments = [...adjustmentsToReview];
            const FETCH_BATCH_SIZE = 5; 

            for (let i = 0; i < updatedAdjustments.length; i += FETCH_BATCH_SIZE) {
                if (abortRef.current) break;
                
                const batch = updatedAdjustments.slice(i, i + FETCH_BATCH_SIZE);
                const promises = batch.map(async (item) => {
                    if (item.balanceDate && item.accountId) {
                        try {
                            let queryDate = item.balanceDate || '';
                            if (queryDate.includes('/')) {
                                const parts = queryDate.split('/');
                                if (downloadDateFormat === 'EU') { 
                                    queryDate = `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
                                } else { 
                                    queryDate = `${parts[2]}-${parts[0].padStart(2, '0')}-${parts[1].padStart(2, '0')}`;
                                }
                            }
                            
                            if (!/^\d{4}-\d{2}-\d{2}$/.test(queryDate)) {
                                if (singleBalanceDate) queryDate = singleBalanceDate;
                                else return item;
                            }

                            const balanceData = await fetchAccountBalance(item.accountId, queryDate);
                            
                            let newAvail = balanceData.balance;
                            let newAdj = item.adjustment;
                            
                            if (typeof item.newBalance === 'number' && !isNaN(item.newBalance)) {
                                newAdj = Number((item.newBalance - newAvail).toFixed(2));
                            } else if (typeof item.newBalance === 'string' && item.newBalance !== '') {
                                const nbParsed = parseFloat(item.newBalance);
                                if (!isNaN(nbParsed)) {
                                    newAdj = Number((nbParsed - newAvail).toFixed(2));
                                }
                            }

                            return {
                                ...item,
                                availableBalance: newAvail,
                                adjustment: newAdj,
                            };
                        } catch (err) {
                            console.error('Failed to fetch balance for account', item.accountId, err);
                            return item;
                        }
                    }
                    return item;
                });

                const batchResults = await Promise.all(promises);
                
                batchResults.forEach((result, idx) => {
                    updatedAdjustments[i + idx] = validateRow(result);
                });

                const pct = Math.round(((i + batch.length) / updatedAdjustments.length) * 100);
                setProgress(pct);
                await new Promise(r => setTimeout(r, 125));
            }

            if (!abortRef.current) {
                setAdjustmentsToReview(updatedAdjustments);
            }
        } catch (err) {
            console.error('Error updating balances', err);
            handleApiError(new Error("Failed to update available balances."));
        } finally {
            setIsUpdatingBalances(false);
            setProgress(0);
        }
    };

    // New Handlers for Bulk Date Updates
    const handleBulkSetToStart = () => {
        setAdjustmentsToReview(prev => prev.map(item => {
            if (item.validFrom) {
                return validateRow({ ...item, effectiveDate: item.validFrom });
            }
            return item;
        }));
    };

    const handleBulkSetToToday = () => {
        const today = getTodayYYYYMMDD();
        setAdjustmentsToReview(prev => prev.map(item => {
            return validateRow({ ...item, effectiveDate: today });
        }));
    };

    const handleBulkSetToCustomDate = () => {
        if (!bulkCustomDate) return;
        setAdjustmentsToReview(prev => prev.map(item => {
            return validateRow({ ...item, effectiveDate: bulkCustomDate });
        }));
    };

    // --- SORT & DELETE HANDLERS ---
    
    const handleSort = (key: keyof AdjustmentReview | 'status' | 'validity') => {
        setSortConfig(current => {
            if (current && current.key === key) {
                return { key, direction: current.direction === 'asc' ? 'desc' : 'asc' };
            }
            return { key, direction: 'asc' };
        });
    };

    const handleDeleteRow = (id: string) => {
        setAdjustmentsToReview(prev => prev.filter(item => item.id !== id));
    };

    const handleDeleteAllErrors = () => {
        setAdjustmentsToReview(prev => prev.filter(item => item.status !== 'error' && !item.isValidationError && item.effectiveDate));
    };

    const handleRemoveSelected = () => {
        if (selectedReviewIds.size === 0) return;
        setShowRemoveConfirmModal(true);
    };

    const confirmRemoveSelected = () => {
        setAdjustmentsToReview(prev => prev.filter(item => !selectedReviewIds.has(item.id)));
        setSelectedReviewIds(new Set());
        setShowRemoveConfirmModal(false);
    };

    const handleClearSelection = () => {
        setSelectedReviewIds(new Set());
    };

    const uniqueAccountNames = useMemo(() => {
        const names = new Set<string>();
        adjustmentsToReview.forEach(r => names.add(r.accountName));
        return Array.from(names).filter(Boolean).map(name => ({ id: name, name }));
    }, [adjustmentsToReview]);

    const sortedReviews = useMemo(() => {
        let filteredReviews = adjustmentsToReview;
        
        // Filter by Review Dropdowns
        if (reviewDepartmentIds.size > 0 || reviewEmployeeGroupIds.size > 0 || reviewEmployeeTypeIds.size > 0 || reviewLeaveAccountNames.size > 0) {
            filteredReviews = filteredReviews.filter(item => {
                // Determine whether this item matches dropdowns using its employeeId mapping
                const emp = item.employeeId ? allEmployees.find(e => e.id === item.employeeId) : undefined;
                
                let matches = true;

                if (reviewDepartmentIds.size < departments.length) {
                    let empMatchesDept = false;
                    if (usePrimaryDepartmentOnly) {
                        if (emp?.primaryDepartmentId !== undefined && emp?.primaryDepartmentId !== null) {
                            empMatchesDept = reviewDepartmentIds.has(String(emp.primaryDepartmentId));
                        }
                    } else {
                        if (emp?.departments && emp.departments.length > 0) {
                            empMatchesDept = emp.departments.some((d: any) => reviewDepartmentIds.has(String(typeof d === 'object' ? d.id : d)));
                        } else if (emp?.departmentId) {
                            empMatchesDept = reviewDepartmentIds.has(String(emp.departmentId));
                        }
                    }
                    if (!empMatchesDept) matches = false;
                }

                if (matches && reviewEmployeeGroupIds.size < employeeGroups.length) {
                    let empMatchesGroup = false;
                    if (emp?.employeeGroups && emp.employeeGroups.length > 0) {
                         empMatchesGroup = emp.employeeGroups.some((g: any) => reviewEmployeeGroupIds.has(String(typeof g === 'object' ? g.id : g)));
                    } else if (emp?.employeeGroupIds && emp.employeeGroupIds.length > 0) {
                         empMatchesGroup = emp.employeeGroupIds.some(g => reviewEmployeeGroupIds.has(String(g)));
                    } else if (emp?.employeeGroupId) {
                         empMatchesGroup = reviewEmployeeGroupIds.has(String(emp.employeeGroupId));
                    }
                    if (!empMatchesGroup) matches = false;
                }

                if (matches && reviewEmployeeTypeIds.size < employeeTypes.length) {
                    let empMatchesType = false;
                    if (emp?.employeeType !== undefined && emp?.employeeType !== null) {
                         empMatchesType = reviewEmployeeTypeIds.has(String(typeof emp.employeeType === 'object' ? emp.employeeType.id : emp.employeeType));
                    } else if (emp?.employeeTypeId !== undefined && emp?.employeeTypeId !== null) {
                         empMatchesType = reviewEmployeeTypeIds.has(String(emp.employeeTypeId));
                    }
                    if (!empMatchesType) matches = false;
                }

                if (matches && reviewLeaveAccountNames.size < uniqueAccountNames.length) {
                    if (!reviewLeaveAccountNames.has(item.accountName)) {
                        matches = false;
                    }
                }

                return matches;
            });
        }
        
        if (searchQuery) {
            const q = searchQuery.toLowerCase();
            filteredReviews = filteredReviews.filter(item => item.employeeName.toLowerCase().includes(q));
        }
        if (!sortConfig) return filteredReviews;
        return [...filteredReviews].sort((a, b) => {
            let valA = a[sortConfig.key as keyof AdjustmentReview];
            let valB = b[sortConfig.key as keyof AdjustmentReview];
            
            // Special handling for Status (Errors) sort to ensure errors come first
            if (sortConfig.key === 'status') {
                const getErrorScore = (item: AdjustmentReview) => item.isValidationError || item.adjustment === 0 || !item.effectiveDate ? 0 : 1;
                valA = getErrorScore(a);
                valB = getErrorScore(b);
            }

            if (sortConfig.key === 'validity') {
                valA = a.validFrom as any;
                valB = b.validFrom as any;
            }

            if (valA! < valB!) return sortConfig.direction === 'asc' ? -1 : 1;
            if (valA! > valB!) return sortConfig.direction === 'asc' ? 1 : -1;
            return 0;
        });
    }, [adjustmentsToReview, sortConfig, searchQuery, reviewDepartmentIds, reviewEmployeeGroupIds, reviewEmployeeTypeIds, reviewLeaveAccountNames, allEmployees, departments.length, employeeGroups.length, employeeTypes.length, usePrimaryDepartmentOnly, uniqueAccountNames.length]);

    useEffect(() => {
        setReviewPage(1);
    }, [searchQuery, reviewDepartmentIds, reviewEmployeeGroupIds, reviewEmployeeTypeIds, reviewLeaveAccountNames, sortConfig, adjustmentsToReview.length]);

    const paginatedReviews = useMemo(() => {
        if (reviewRowsPerPage === 'All') return sortedReviews;
        const start = (reviewPage - 1) * reviewRowsPerPage;
        return sortedReviews.slice(start, start + reviewRowsPerPage);
    }, [sortedReviews, reviewPage, reviewRowsPerPage]);

    const [bulkEditDateWarning, setBulkEditDateWarning] = useState<{skippedflexCount: number} | null>(null);

    const handleApplyBulkEdit = () => {
        if (!bulkEditField || selectedReviewIds.size === 0) return;
        
        let skippedFlexAccounts = 0;

        if (bulkEditField === 'effectiveDate' && bulkEditEffectiveDateType === 'start_date') {
            adjustmentsToReview.forEach(item => {
                if (selectedReviewIds.has(item.id)) {
                    if (item.accountTypeCategory === 'FLEX/TOIL' && (!item.validFrom || item.validFrom === 'N/A')) {
                        skippedFlexAccounts++;
                    }
                }
            });

            if (skippedFlexAccounts > 0) {
                setBulkEditDateWarning({ skippedflexCount: skippedFlexAccounts });
                return; // Stop and wait for user confirmation
            }
        }

        executeApplyBulkEdit();
    };

    const executeApplyBulkEdit = () => {
        setAdjustmentsToReview(prev => prev.map(item => {
            if (selectedReviewIds.has(item.id)) {
                if (bulkEditField === 'adjustment') {
                    const num = parseFloat(bulkEditValueAdjustment);
                    if (!isNaN(num)) {
                        let updatedNewBalance = item.newBalance;
                        if (typeof item.availableBalance === 'number' && !isNaN(item.availableBalance)) {
                            updatedNewBalance = Number((item.availableBalance + num).toFixed(2));
                        }
                        return validateRow({ ...item, adjustment: num, newBalance: updatedNewBalance });
                    }
                } else if (bulkEditField === 'newBalance') {
                    const num = parseFloat(bulkEditValueNewBalance);
                    if (!isNaN(num)) {
                        let derivedAdjustment = item.adjustment;
                        if (typeof item.availableBalance === 'number' && !isNaN(item.availableBalance)) {
                            derivedAdjustment = Number((num - item.availableBalance).toFixed(2));
                        }
                        return validateRow({ ...item, adjustment: derivedAdjustment, newBalance: num });
                    }
                } else if (bulkEditField === 'effectiveDate') {
                    let newDate = bulkEditValueEffectiveDate;
                    if (bulkEditEffectiveDateType === 'today') newDate = getTodayYYYYMMDD();
                    if (bulkEditEffectiveDateType === 'start_date') {
                        if (item.accountTypeCategory === 'FLEX/TOIL' && (!item.validFrom || item.validFrom === 'N/A')) {
                            return item; // Skip modifying this item
                        }
                        newDate = (item.validFrom && item.validFrom !== 'N/A') ? item.validFrom : getTodayYYYYMMDD();
                    }
                    if (newDate) {
                        return validateRow({ ...item, effectiveDate: newDate });
                    }
                } else if (bulkEditField === 'comment') {
                    return { ...item, comment: bulkEditValueComment };
                }
            }
            return item;
        }));
    };

    const handleSelectFiltered = () => {
        setSelectedReviewIds(new Set(sortedReviews.map(r => r.id)));
    };

    const handleSelectAll = () => {
        setSelectedReviewIds(new Set(adjustmentsToReview.map(r => r.id)));
    };

    const toggleReviewSelection = (id: string) => {
        const newSet = new Set(selectedReviewIds);
        if (newSet.has(id)) newSet.delete(id);
        else newSet.add(id);
        setSelectedReviewIds(newSet);
    };

    const executeBatchUpdate = async () => {
        abortRef.current = false;
        setShowConfirmModal(false);
        setCurrentStep('processing');
        setIsLoading(prev => ({...prev, submitting: true}));
        setProgress(0);
        
        // Prevent Computer Sleep
        let wakeLock: any = null;
        try {
            if ('wakeLock' in navigator) {
                // @ts-ignore - navigator.wakeLock is part of modern standard but TS might need type defs
                wakeLock = await navigator.wakeLock.request('screen');
                console.log('Wake Lock active');
            }
        } catch (err) {
            console.warn('Wake Lock request failed:', err);
        }

        const summary: AdjustmentReview[] = [];
        // OPTIMIZATION: Reduced batch size for stability
        const UPDATE_BATCH_SIZE = 5;
        
        try {
            const pendingAdjustments = [...adjustmentsToReview];
            const total = pendingAdjustments.length;
            
            for (let i = 0; i < pendingAdjustments.length; i += UPDATE_BATCH_SIZE) {
                if (abortRef.current) {
                    const remaining = pendingAdjustments.slice(i).map(adj => ({
                        ...adj,
                        status: 'error' as const,
                        error: 'Aborted by user',
                        timestamp: new Date().toLocaleString('en-US')
                    }));
                    summary.push(...remaining);
                    break;
                }

                const batch = pendingAdjustments.slice(i, i + UPDATE_BATCH_SIZE);
                
                // Map batch items to promises
                const promises = batch.map(async (adj) => {
                    // Skip if pre-check failed (shouldn't happen if button disabled, but safe guard)
                    if (adj.status === 'error' && adj.isValidationError) {
                         return adj;
                    }

                    const timestamp = new Date().toLocaleString('en-US');

                    if (adj.adjustment === 0) {
                        return { ...adj, status: 'skipped' as const, error: 'Skipped - adjustment value is 0', timestamp };
                    }

                    try {
                        let payloadComment = "API BULK UPDATE.";
                        if (adj.comment && adj.comment.trim()) {
                            payloadComment = `API BULK UPDATE: ${adj.comment.trim()}`;
                        }

                        if (adj.accountTypeCategory === 'FLEX/TOIL') {
                            await postFlexBalanceAdjustment(adj.accountId, { value: adj.adjustment, effectiveDate: adj.effectiveDate, comment: payloadComment });
                        } else {
                            await postBalanceAdjustment(adj.accountId, { value: adj.adjustment, effectiveDate: adj.effectiveDate, comment: payloadComment });
                        }
                        
                        let postAdjustmentBalance: number | undefined = undefined;
                        try {
                            const newBalanceRes = await fetchAccountBalance(adj.accountId, adj.effectiveDate);
                            postAdjustmentBalance = newBalanceRes.balance;
                        } catch (err) {
                            console.warn(`Failed to fetch new balance for account ${adj.accountId}`, err);
                        }
                        
                        // Success Result
                        return { ...adj, status: 'success' as const, timestamp, postAdjustmentBalance };
                    } catch (err: any) {
                        // Error Result
                        return { ...adj, status: 'error' as const, error: err.message, timestamp };
                    }
                });

                // Wait for batch to complete
                const results = await Promise.all(promises);
                
                // Process results
                summary.push(...results);
                
                // Update UI state for this batch
                setAdjustmentsToReview(prev => {
                    const next = [...prev];
                    results.forEach(res => {
                        const index = next.findIndex(item => item.accountId === res.accountId);
                        if (index !== -1) next[index] = res;
                    });
                    return next;
                });

                // Slight delay to allow browser event loop to clear network stack
                await new Promise(resolve => setTimeout(resolve, 250));
                
                // Update Progress
                const completed = Math.min(i + UPDATE_BATCH_SIZE, total);
                const percent = Math.round((completed / total) * 100);
                setProgress(percent);
            }

        } finally {
             // Release Wake Lock
             if (wakeLock) {
                try {
                    await wakeLock.release();
                    console.log('Wake Lock released');
                } catch(e) { console.error(e); }
             }
             setUpdateSummary(summary);
             setIsLoading(prev => ({...prev, submitting: false}));
             setCurrentStep('summary');
        }
    };

    const handleSelectAllTypes = (e: React.ChangeEvent<HTMLInputElement>) => {
        const filteredTypes = accountTypes.filter(type => type.name.toLowerCase().includes(accountTypesSearch.toLowerCase()));
        if (e.target.checked) {
            setSelectedTypeIds(prev => {
                const newSet = new Set(prev);
                filteredTypes.forEach(t => newSet.add(t.id));
                return newSet;
            });
        } else {
            setSelectedTypeIds(prev => {
                const newSet = new Set(prev);
                filteredTypes.forEach(t => newSet.delete(t.id));
                return newSet;
            });
        }
    };
    
    const handleStartDateChange = (e: React.ChangeEvent<HTMLInputElement>) => {
        const newStart = e.target.value;
        setDateRange(prev => {
            let newEnd = prev.end;
            if (newStart) {
                const [y, m, d] = newStart.split('-').map(Number);
                const date = new Date(Date.UTC(y, m - 1, d));
                date.setUTCFullYear(date.getUTCFullYear() + 1);
                date.setUTCDate(date.getUTCDate() - 1);
                newEnd = date.toISOString().split('T')[0];
            }
            return { start: newStart, end: newEnd };
        });
    };
    
    const handleExportResults = () => {
        const headers = ["Time of Change", "Salary Identifier", "Employee", "Account", "Validity Period", "Unit Type", "Adjustment", "New Balance", "Effective Date", "Result Message", "Comment"];
        const data = updateSummary.map(item => {
            const validFrom = item.validFrom ? formatDateForDisplay(item.validFrom, downloadDateFormat) : 'N/A';
            const validTo = item.validTo ? formatDateForDisplay(item.validTo, downloadDateFormat) : '∞';
            
            return {
                "Time of Change": item.timestamp || "",
                "Salary Identifier": item.salaryIdentifier || "",
                "Employee": item.employeeName,
                "Account": item.accountName,
                "Validity Period": `${validFrom} - ${validTo}`,
                "Unit Type": item.unit || "N/A",
                "Adjustment": item.adjustment,
                "New Balance": item.postAdjustmentBalance !== undefined ? Math.round(item.postAdjustmentBalance * 100) / 100 : "",
                "Effective Date": formatDateForDisplay(item.effectiveDate, downloadDateFormat), // Reuse preference
                "Result Message": item.status === 'error' || item.status === 'skipped' ? item.error : "Updated Successfully",
                "Comment": item.comment || ""
            };
        });
        
        const ws = XLSX.utils.json_to_sheet(data, { header: headers });
        
        // --- EXPORT STYLING START ---
        const range = XLSX.utils.decode_range(ws['!ref']);
        const borderStyle = {
            top: { style: "thin", color: { rgb: "d9d9d9" } },
            bottom: { style: "thin", color: { rgb: "d9d9d9" } },
            left: { style: "thin", color: { rgb: "d9d9d9" } },
            right: { style: "thin", color: { rgb: "d9d9d9" } }
        };

        const resultMsgIndex = headers.indexOf("Result Message");

        for (let R = range.s.r; R <= range.e.r; ++R) {
            for (let C = range.s.c; C <= range.e.c; ++C) {
                const cellAddress = XLSX.utils.encode_cell({ c: C, r: R });
                if (!ws[cellAddress]) continue;

                if (!ws[cellAddress].s) ws[cellAddress].s = {};
                ws[cellAddress].s.border = borderStyle;

                // Header Row (Row 0): Background #112540, White text
                if (R === 0) {
                    ws[cellAddress].s.fill = { fgColor: { rgb: "112540" } };
                    ws[cellAddress].s.font = { 
                        bold: true, 
                        color: { rgb: "FFFFFF" },
                        name: "Calibri",
                        sz: 11
                    };
                } else {
                    // Data Rows
                    if (C === resultMsgIndex) {
                        const cellVal = ws[cellAddress].v;
                        if (cellVal === "Updated Successfully") {
                            ws[cellAddress].s.font = { color: { rgb: "008000" }, bold: true }; // Green
                        } else if (cellVal && cellVal.toString().startsWith("Skipped")) {
                            ws[cellAddress].s.font = { color: { rgb: "FFA500" }, bold: true }; // Orange/Yellow
                        } else {
                            ws[cellAddress].s.font = { color: { rgb: "CC0000" } }; // Red
                        }
                    }
                }
            }
        }
        
        // Auto-width
        const wscols = headers.map(h => ({ wch: h.length + 10 }));
        ws['!cols'] = wscols;
        // --- EXPORT STYLING END ---

        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Update Results");
        XLSX.writeFile(wb, "Planday_LeaveAdjustment_Results.xlsx");
    };

    // Tooltip Logic
    const handleTooltipEnter = (e: React.MouseEvent, content: React.ReactNode) => {
        const rect = e.currentTarget.getBoundingClientRect();
        // Position: Centered horizontally below the element, shifted down by 8px
        setActiveTooltip({
            x: rect.left + rect.width / 2,
            y: rect.bottom + 8,
            content
        });
    };

    const handleTooltipLeave = () => {
        setActiveTooltip(null);
    };


    const stepConfig = getStepConfig(updateMethod);
    const stepIndex = stepConfig.order.indexOf(currentStep as any);

    // Sort summary: Errors first, then skipped, then success
    const sortedSummary = [...updateSummary].sort((a, b) => {
        const order = { 'error': 1, 'skipped': 2, 'success': 3, 'pending': 4 };
        const weightA = order[a.status || 'pending'];
        const weightB = order[b.status || 'pending'];
        return weightA - weightB;
    });

    const hasValidityErrors = adjustmentsToReview.some(a => a.status === 'error' && a.isValidationError);
    const hasMissingDates = adjustmentsToReview.some(a => !a.effectiveDate);
    // hasValidationErrors is true if any error exists, but the user only sees them if attempted submit or validity error exists.
    const hasValidationErrors = hasValidityErrors || hasMissingDates;
    
    // Condition to show Date Format Report:
    // Show if a detection was made that differs from user preference, OR if we had to fallback due to ambiguity.
    // If we detected exactly what the user set, hide it.
    // EXCEPTION: If source is 'inherited' (meaning it was detected by context), we do NOT show report for it.
    const shouldShowDateReport = detectedColumnFormats && (
        ((detectedColumnFormats.effective.source === 'fallback' || detectedColumnFormats.effective.format !== downloadDateFormat) && detectedColumnFormats.effective.source !== 'inherited' && detectedColumnFormats.effective.source !== 'empty') ||
        ((detectedColumnFormats.validFrom.source === 'fallback' || detectedColumnFormats.validFrom.format !== downloadDateFormat) && detectedColumnFormats.validFrom.source !== 'inherited' && detectedColumnFormats.validFrom.source !== 'empty') ||
        ((detectedColumnFormats.validTo.source === 'fallback' || detectedColumnFormats.validTo.format !== downloadDateFormat) && detectedColumnFormats.validTo.source !== 'inherited' && detectedColumnFormats.validTo.source !== 'empty')
    );

    const uniqueBalanceDates = new Set<string>();
    adjustmentsToReview.forEach(r => {
        if (r.balanceDate && r.balanceDate !== 'N/A') uniqueBalanceDates.add(r.balanceDate);
    });
    const singleBalanceDate = uniqueBalanceDates.size === 1 ? Array.from(uniqueBalanceDates)[0] : null;

    // Helper to render label for the detected format
    const renderFormatLabel = (result: ColumnDetectionResult, columnContext: string) => {
        const example = result.format === 'EU' ? '30/01' : '01/30';
        
        if (result.source === 'detected') {
            return (
                <div className="flex items-center gap-1 group relative">
                    <span className="font-mono font-bold text-gray-800">Detected: {result.format}</span>
                    <InfoIcon className="h-3 w-3 text-blue-400 cursor-help" />
                    <div className="hidden group-hover:block absolute bottom-full left-1/2 -translate-x-1/2 mb-1 w-64 p-2 bg-gray-800 text-white text-xs rounded shadow-lg z-50">
                        We determined the {columnContext} in the file to be {result.format} date format (e.g. {example}).
                    </div>
                </div>
            );
        }
        if (result.source === 'empty') {
            return (
                <div className="flex items-center gap-1 group relative">
                    <span className="font-mono font-bold text-gray-500">Not Applicable (Empty)</span>
                    <InfoIcon className="h-3 w-3 text-gray-400 cursor-help" />
                    <div className="hidden group-hover:block absolute bottom-full left-1/2 -translate-x-1/2 mb-1 w-64 p-2 bg-gray-800 text-white text-xs rounded shadow-lg z-50">
                        No valid dates were found in the {columnContext} (e.g. all empty or N/A).
                    </div>
                </div>
            );
        }
        if (result.source === 'inherited') {
            return (
                <div className="flex items-center gap-1 group relative">
                    <span className="font-mono font-bold text-gray-800">Detected: {result.format} (Inherited)</span>
                    <InfoIcon className="h-3 w-3 text-blue-400 cursor-help" />
                    <div className="hidden group-hover:block absolute bottom-full left-1/2 -translate-x-1/2 mb-1 w-64 p-2 bg-gray-800 text-white text-xs rounded shadow-lg z-50">
                        We inherited the {result.format} date format from another context/column.
                    </div>
                </div>
            );
        }
        return (
            <div className="flex items-center gap-1 group relative">
                <span className="font-mono font-bold text-amber-700">Ambiguous (Using {result.format})</span>
                <InfoIcon className="h-3 w-3 text-amber-500 cursor-help" />
                <div className="hidden group-hover:block absolute bottom-full left-1/2 -translate-x-1/2 mb-1 w-64 p-2 bg-gray-800 text-white text-xs rounded shadow-lg z-50">
                    Dates in this column were ambiguous (e.g. 01/02/2025). We are interpreting them as {result.format} based on your App Setting.
                </div>
            </div>
        );
    };

    const shouldShowBalanceColumns = updateMethod === 'editor' ? !!includeBalance : adjustmentsToReview.some(r => r.availableBalance !== undefined && r.availableBalance !== 'N/A' && r.availableBalance !== 'Not retrieved' && r.availableBalance !== '');

    return (
        <div className="min-h-screen font-sans flex flex-col">
            <div className="container mx-auto px-8 py-8 flex-grow">
                <PageHeader />
                <div className="my-12 max-w-[1280px] mx-auto">
                    <Stepper current={stepIndex} steps={stepConfig.labels} />
                </div>

                {currentStep !== 'auth' && (
                    <div className="max-w-[896px] mx-auto flex items-center justify-center gap-3 mb-8 text-sm text-gray-700">
                        <button onClick={loadConfigData} disabled={isLoading.types} className="text-gray-500 hover:text-gray-800 transition-colors disabled:opacity-50" title="Refresh portal data">
                             <RefreshIcon className={`h-4 w-4 ${isLoading.types ? 'animate-spin' : ''}`} />
                        </button>
                        <span>Logged in: <strong>{portalName || 'Loading...'}</strong></span>
                        <span className="text-gray-300">|</span>
                        <button onClick={handleLogout} className="text-gray-600 hover:text-gray-900 hover:underline underline-offset-2">
                            Change credentials (log out)
                        </button>
                    </div>
                )}

                <main>
                    {error && <div className="bg-red-100 border border-red-400 text-red-700 px-4 py-3 rounded-lg relative mb-6 max-w-[896px] mx-auto" role="alert">{error}</div>}

                    {currentStep === 'auth' && <AuthStep onAuthSuccess={handleAuthSuccess} />}

                    {currentStep === 'configure' && <div className="bg-white p-8 rounded-lg shadow-md max-w-[896px] mx-auto">
                        <div className="mb-10">
                            <h2 className="text-2xl font-bold mb-6 text-gray-800">Select Update Method</h2>
                            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                                <button
                                    onClick={() => setUpdateMethod('excel')}
                                    className={`flex flex-col items-center justify-center p-8 rounded-xl border-2 transition-all ${
                                        updateMethod === 'excel'
                                            ? 'border-blue-500 bg-blue-50/50'
                                            : 'border-gray-200 hover:border-blue-200 hover:bg-gray-50'
                                    }`}
                                >
                                    <div className="h-12 w-12 bg-blue-100 text-blue-600 rounded-full flex items-center justify-center mb-4">
                                        <svg className="h-6 w-6" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 17v-2m3 2v-4m3 4v-6m2 10H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                                        </svg>
                                    </div>
                                    <h3 className="text-lg font-semibold text-gray-900 mb-2">Use Excel</h3>
                                    <p className="text-sm text-gray-500 text-center">Download a template and edit in Excel</p>
                                </button>
                                <button
                                    onClick={() => setUpdateMethod('editor')}
                                    className={`flex flex-col items-center justify-center p-8 rounded-xl border-2 transition-all ${
                                        updateMethod === 'editor'
                                            ? 'border-blue-500 bg-blue-50/50'
                                            : 'border-gray-200 hover:border-blue-200 hover:bg-gray-50'
                                    }`}
                                >
                                    <div className="h-12 w-12 bg-blue-100 text-blue-600 rounded-full flex items-center justify-center mb-4">
                                        <svg className="h-6 w-6" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z" />
                                        </svg>
                                    </div>
                                    <h3 className="text-lg font-semibold text-gray-900 mb-2">Use Editor</h3>
                                    <p className="text-sm text-gray-500 text-center">Edit directly in the browser table</p>
                                </button>
                            </div>
                        </div>
                        
                        <div className="pt-8 border-t border-gray-100">
                            {/* ... Configuration Step UI ... */}
                            <h2 className="text-2xl font-bold mb-6 text-gray-800">{updateMethod === 'excel' ? 'Configure & Download Template' : 'Select Accounts to Update'}</h2>
                            <div className="space-y-6">
                                 <div>
                                <div className="flex justify-between items-center mb-2">
                                    <label className="block text-sm font-medium text-gray-700">1. Select Account Types (Policies) to Include</label>
                                    {accountTypes.length > 0 && (
                                        <div className="flex items-center">
                                            <input
                                                id="select-all-types"
                                                type="checkbox"
                                                className="h-4 w-4 rounded border-gray-300 text-blue-600 focus:ring-blue-500 cursor-pointer"
                                                checked={accountTypes.length > 0 && accountTypes.filter(type => type.name.toLowerCase().includes(accountTypesSearch.toLowerCase())).length > 0 && accountTypes.filter(type => type.name.toLowerCase().includes(accountTypesSearch.toLowerCase())).every(type => selectedTypeIds.has(type.id))}
                                                ref={input => { 
                                                    if (input) { 
                                                        const filtered = accountTypes.filter(type => type.name.toLowerCase().includes(accountTypesSearch.toLowerCase()));
                                                        const someSelected = filtered.some(type => selectedTypeIds.has(type.id));
                                                        const allSelected = filtered.length > 0 && filtered.every(type => selectedTypeIds.has(type.id));
                                                        input.indeterminate = someSelected && !allSelected;
                                                    }
                                                }}
                                                onChange={handleSelectAllTypes}
                                            />
                                            <label htmlFor="select-all-types" className="ml-2 block text-sm text-gray-900 cursor-pointer">
                                                Select All
                                            </label>
                                        </div>
                                    )}
                                </div>
                                {accountTypes.length > 0 && (
                                    <div className="mb-3">
                                        <input
                                            type="text"
                                            placeholder="Search account types..."
                                            value={accountTypesSearch}
                                            onChange={(e) => setAccountTypesSearch(e.target.value)}
                                            className="w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm text-sm focus:outline-none focus:ring-blue-500 focus:border-blue-500"
                                        />
                                    </div>
                                )}
                                {isLoading.types ? <Loader text="Loading account types..." /> : <div className="grid grid-cols-2 md:grid-cols-3 gap-3 max-h-72 overflow-y-auto p-3 bg-gray-50 rounded-md border">{accountTypes.filter(type => type.name.toLowerCase().includes(accountTypesSearch.toLowerCase())).map(type => (
                                    <label key={type.id} className="flex items-start space-x-3 p-2 bg-white border border-gray-200 rounded-md cursor-pointer hover:bg-gray-100 transition-colors"><input type="checkbox" checked={selectedTypeIds.has(type.id)} onChange={() => setSelectedTypeIds(p => {const n=new Set(p); n.has(type.id)?n.delete(type.id):n.add(type.id); return n;})} className="mt-0.5 h-4 w-4 rounded border-gray-300 text-blue-600 focus:ring-blue-500 flex-shrink-0"/><div className="flex-1 flex items-start justify-between gap-2"><span className="text-sm break-words leading-tight" title={type.name}>{type.name}</span>{type.absenceType === 'Flextime' ? <span className="flex-shrink-0 inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-semibold bg-purple-100 text-purple-800">FLEX/TOIL</span> : (type.accruingRate?.value === 0 && type.accruingRate?.unit?.type === 'Percent' ? <span className="flex-shrink-0 inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-semibold bg-sky-100 text-sky-800">Fixed</span> : <span className="flex-shrink-0 inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-semibold bg-teal-100 text-teal-800">Accrued</span>)}</div></label>
                                ))}</div>}
                            </div>
                             
                            <div className="grid md:grid-cols-2 gap-6">
                                <div>
                                    <div className="flex items-center mb-2">
                                        <label className="block text-sm font-medium text-gray-700 mr-2">2. Select Validity Period (Required)</label>
                                        <div className="relative group">
                                            <InfoIcon className="h-5 w-5 text-gray-400 hover:text-blue-500 cursor-help" />
                                            <div className="absolute bottom-full left-1/2 transform -translate-x-1/2 mb-2 w-72 p-3 bg-gray-900 text-white text-xs rounded-md shadow-xl opacity-0 invisible group-hover:opacity-100 group-hover:visible transition-all duration-200 z-50 text-left">
                                                <p><strong>Current:</strong> Only active accounts (Today is within the valid period).</p>
                                                <p className="mt-1"><strong>Current + Upcoming:</strong> Active accounts plus any that start in the future. (Excludes past/expired).</p>
                                                <p className="mt-1"><strong>Select Dates:</strong> Manually choose a range. You can check 'Include Inactive' to find expired accounts in that range.</p>
                                                <div className="absolute top-full left-1/2 transform -translate-x-1/2 border-4 border-transparent border-t-gray-900"></div>
                                            </div>
                                        </div>
                                    </div>
                                    
                                    <div className="space-y-2 mb-4">
                                        <label className="flex items-center space-x-3 cursor-pointer">
                                            <input type="radio" name="validityMode" value="current" checked={validityMode === 'current'} onChange={() => setValidityMode('current')} className="form-radio text-blue-600 h-4 w-4" />
                                            <span className="text-sm text-gray-700">Current</span>
                                        </label>
                                        <label className="flex items-center space-x-3 cursor-pointer">
                                            <input type="radio" name="validityMode" value="current_future" checked={validityMode === 'current_future'} onChange={() => setValidityMode('current_future')} className="form-radio text-blue-600 h-4 w-4" />
                                            <span className="text-sm text-gray-700">Current + Upcoming</span>
                                        </label>
                                        <label className="flex items-center space-x-3 cursor-pointer">
                                            <input type="radio" name="validityMode" value="custom" checked={validityMode === 'custom'} onChange={() => setValidityMode('custom')} className="form-radio text-blue-600 h-4 w-4" />
                                            <span className="text-sm text-gray-700">Select Dates</span>
                                        </label>
                                    </div>

                                    {validityMode === 'custom' && (
                                        <div className="pl-7 space-y-3 transition-all duration-300 ease-in-out">
                                            <div className="flex items-center space-x-2">
                                               <input type="date" value={dateRange.start} onChange={handleStartDateChange} className="w-full px-3 py-2 bg-white border border-gray-300 rounded-md shadow-sm text-gray-900 focus:outline-none focus:ring-2 focus:ring-blue-500" placeholder="Start Date" />
                                               <span className="text-gray-500 text-sm">to</span>
                                               <input type="date" value={dateRange.end} onChange={e => setDateRange(p => ({...p, end: e.target.value}))} className="w-full px-3 py-2 bg-white border border-gray-300 rounded-md shadow-sm text-gray-900 focus:outline-none focus:ring-2 focus:ring-blue-500" placeholder="End Date" />
                                            </div>
                                            <label className="inline-flex items-center cursor-pointer select-none">
                                                <input 
                                                    type="checkbox" 
                                                    checked={includeInactive} 
                                                    onChange={e => setIncludeInactive(e.target.checked)} 
                                                    className="h-4 w-4 rounded border-gray-300 text-blue-600 focus:ring-blue-500" 
                                                />
                                                <span className="ml-2 text-sm text-gray-700">Include inactive (expired) accounts in this range</span>
                                            </label>
                                        </div>
                                    )}
                                </div>

                                <div className={`p-4 rounded-md transition-colors ${balanceOptionError ? 'bg-red-50 border border-red-300' : ''}`}>
                                    <div className="flex items-center mb-2">
                                        <label className={`block text-sm font-medium mr-2 ${balanceOptionError ? 'text-red-700' : 'text-gray-700'}`}>3. Include Available Balance?</label>
                                        <div className="relative group">
                                            <InfoIcon className={`h-5 w-5 cursor-help ${balanceOptionError ? 'text-red-400 hover:text-red-500' : 'text-gray-400 hover:text-blue-500'}`} />
                                            <div className="absolute bottom-full left-1/2 transform -translate-x-1/2 mb-2 w-96 p-4 bg-gray-900 text-white text-xs rounded-md shadow-xl opacity-0 invisible group-hover:opacity-100 group-hover:visible transition-all duration-200 z-50 text-left leading-relaxed">
                                                <ul className="mb-2 space-y-2 list-disc pl-4">
                                                    <li><strong>No:</strong> The template will not contain the available balance from the employees' leave accounts. Only use this option if you already know the current available balance and want to make quick adjustments.</li>
                                                    <li><strong>Yes:</strong> The app will fetch the available balance on the date selected. Inside the file, you will be able to make adjustments according to the available balance and set a "New balance". The file will calculate the difference between the "Available balance" and "New balance" to post the needed adjustments. Consider matching the date selected (Balance Date) with the “Effective date” in the template.</li>
                                                </ul>
                                                
                                                <div className="pt-2 border-t border-gray-700 mt-2">
                                                    <p className="text-yellow-300 font-bold mb-1">Important info on the Available Balance:</p>
                                                    <ul className="space-y-1 list-disc pl-4 text-yellow-300">
                                                        <li><strong>Accounts with Accruals:</strong> If you select a date that is earlier than the account end date, the file will only show the balance accrued up to that specific date. Future accruals will not be included in the available balance.</li>
                                                        <li><strong>Past Leave:</strong> All approved leave requests BEFORE this date are already deducted from the displayed available balance.</li>
                                                        <li><strong>Future Leave:</strong> Any approved leave requests AFTER this date are NOT deducted from the displayed available balance.</li>
                                                    </ul>
                                                </div>
                                                <div className="absolute top-full left-1/2 transform -translate-x-1/2 border-4 border-transparent border-t-gray-900"></div>
                                            </div>
                                        </div>
                                    </div>

                                    {balanceOptionError && (
                                        <p className="text-xs text-red-600 font-medium mb-2 animate-pulse">{balanceOptionError}</p>
                                    )}

                                    <div className="flex items-center space-x-4 mb-3">
                                        <label className="inline-flex items-center">
                                            <input type="radio" className="form-radio text-blue-600" name="includeBalance" checked={includeBalance === true} onChange={() => { setIncludeBalance(true); setBalanceOptionError(null); }} />
                                            <span className="ml-2 text-sm text-gray-700">Yes</span>
                                        </label>
                                        <label className="inline-flex items-center">
                                            <input type="radio" className="form-radio text-blue-600" name="includeBalance" checked={includeBalance === false} onChange={() => { setIncludeBalance(false); setBalanceOptionError(null); }} />
                                            <span className="ml-2 text-sm text-gray-700">No</span>
                                        </label>
                                    </div>

                                    <div className={`flex items-center space-x-2 transition-opacity ${!includeBalance ? 'opacity-50 pointer-events-none' : ''}`}>
                                        <input 
                                            type="date" 
                                            value={balanceDate} 
                                            onChange={e => { setBalanceDate(e.target.value); setBalanceOptionError(null); }}
                                            disabled={!includeBalance}
                                            className="flex-1 px-3 py-2 bg-white border border-gray-300 rounded-md shadow-sm text-gray-900 focus:outline-none focus:ring-2 focus:ring-blue-500 disabled:bg-gray-100" 
                                        />
                                        <button 
                                            onClick={() => { setBalanceDate(getTodayYYYYMMDD()); setBalanceOptionError(null); }}
                                            disabled={!includeBalance}
                                            className="px-4 py-2 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded-md border border-gray-300 text-sm font-medium transition-colors disabled:opacity-50 whitespace-nowrap"
                                        >
                                            Today
                                        </button>
                                    </div>
                                </div>
                            </div>
                        </div>

                        {/* Filters Section */}
                        <div className="mt-8 bg-gray-50 p-4 border border-gray-200 rounded-lg">
                            <h3 className="text-sm font-bold text-gray-700 mb-3">Filter Employees (Optional)</h3>
                            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                                <div>
                                    <div className="flex justify-between items-center mb-2">
                                        <label className="block text-xs font-semibold text-gray-500 uppercase tracking-wider">
                                            {usePrimaryDepartmentOnly ? 'PRIMARY DEPARTMENT' : 'DEPARTMENT'}
                                        </label>
                                        <div className="flex items-center space-x-1 group relative">
                                            <button 
                                                onClick={() => setUsePrimaryDepartmentOnly(!usePrimaryDepartmentOnly)}
                                                className={`relative inline-flex h-4 w-7 items-center rounded-full transition-colors ${usePrimaryDepartmentOnly ? 'bg-orange-500' : 'bg-gray-300'}`}
                                            >
                                                <span className={`inline-block h-3 w-3 transform rounded-full bg-white transition-transform ${usePrimaryDepartmentOnly ? 'translate-x-3.5' : 'translate-x-0.5'}`} />
                                            </button>
                                            <InfoIcon className="h-4 w-4 text-orange-400 hover:text-orange-500 cursor-help" />
                                            <div className="absolute bottom-full left-1/2 transform -translate-x-1/2 mb-2 w-64 p-3 bg-gray-900 text-white text-xs rounded-md shadow-xl opacity-0 invisible group-hover:opacity-100 group-hover:visible transition-all duration-200 z-50 text-left leading-relaxed">
                                                <p className="mb-1">You can filter by Primary Department.</p>
                                                <p>Note: Employees without a designated primary department will not appear in the template file. To use this filter, ensure the feature is enabled and assigned to everyone.</p>
                                            </div>
                                        </div>
                                    </div>
                                    <MultiSelectDropdown 
                                        label={usePrimaryDepartmentOnly ? "Primary Department" : "Department"} 
                                        pluralLabel={usePrimaryDepartmentOnly ? "Primary Departments" : "Departments"}
                                        options={departments} 
                                        selectedIds={selectedDepartmentIds} 
                                        onChange={setSelectedDepartmentIds} 
                                    />
                                </div>
                                <div>
                                    <label className="block text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">EMPLOYEE GROUP</label>
                                    <MultiSelectDropdown 
                                        label="Employee Group" 
                                        pluralLabel="Employee groups"
                                        options={employeeGroups} 
                                        selectedIds={selectedEmployeeGroupIds} 
                                        onChange={setSelectedEmployeeGroupIds} 
                                    />
                                </div>
                                <div>
                                    <label className="block text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">EMPLOYEE TYPE</label>
                                    <MultiSelectDropdown 
                                        label="Employee Type" 
                                        pluralLabel="Employee types"
                                        options={employeeTypes} 
                                        selectedIds={selectedEmployeeTypeIds} 
                                        onChange={setSelectedEmployeeTypeIds} 
                                    />
                                </div>
                            </div>
                        </div>
                        
                        <div className="mt-8 pt-6 border-t border-gray-200">
                           {isLoading.template ? (
                               <div className="flex flex-col items-center">
                                   <ProgressBar progress={progress} text={loadingText} />
                                   <button 
                                       onClick={() => setShowStopProcessModal(true)}
                                       className="mt-6 bg-red-100 hover:bg-red-200 text-red-700 font-semibold py-2 px-6 rounded-md shadow-sm transition-colors duration-200"
                                    >
                                       Stop Process
                                   </button>
                               </div>
                           ) : (
                               <div className="flex justify-end space-x-4 items-center">
                                    {/* Date Format Selection for Download */}
                                    <div className="flex items-center mr-4 bg-gray-50 px-3 py-2 rounded-md border border-gray-200">
                                            <span className="text-sm font-medium text-gray-700 mr-2">Date Format:</span>
                                            <select 
                                                value={downloadDateFormat} 
                                                onChange={(e) => setDownloadDateFormat(e.target.value as 'EU' | 'US')}
                                                className="text-sm border-gray-300 rounded focus:ring-blue-500 focus:border-blue-500"
                                            >
                                                <option value="EU">EU (DD/MM/YYYY)</option>
                                                <option value="US">US (MM/DD/YYYY)</option>
                                            </select>
                                            <div className="relative group ml-2">
                                                <InfoIcon className="h-4 w-4 text-gray-400 hover:text-blue-500 cursor-help" />
                                                <div className="absolute bottom-full left-1/2 transform -translate-x-1/2 mb-2 w-64 p-3 bg-gray-900 text-white text-xs rounded-md shadow-xl opacity-0 invisible group-hover:opacity-100 group-hover:visible transition-all duration-200 z-50 text-left">
                                                    <p>Choose the date format for the downloaded Excel file.</p>
                                                    <ul className="list-disc list-inside mt-1 space-y-1 text-gray-300">
                                                        <li><strong>EU:</strong> 31/01/2025 (Day First)</li>
                                                        <li><strong>US:</strong> 01/31/2025 (Month First)</li>
                                                    </ul>
                                                </div>
                                            </div>
                                    </div>

                                    {updateMethod === 'excel' && (
                                        <button onClick={() => setCurrentStep('upload')} className="bg-white hover:bg-gray-100 text-gray-700 font-semibold py-2 px-4 rounded-md border border-gray-300 shadow-sm">Next &rarr; Upload file</button>
                                    )}
                                    <button onClick={handleDownloadTemplate} disabled={selectedTypeIds.size === 0} className="bg-blue-600 hover:bg-blue-700 text-white font-bold py-2 px-4 rounded-md disabled:bg-gray-400 flex items-center justify-center min-w-[200px]">{updateMethod === 'excel' ? 'Download Template' : 'Next \u2192 Make Adjustments'}</button>
                               </div>
                           )}
                        </div>
                        </div>
                    </div>}

                    {currentStep === 'upload' && <div className="bg-white p-8 rounded-lg shadow-md max-w-[672px] mx-auto">
                         {/* ... Upload Step UI ... */}
                        <h2 className="text-2xl font-bold mb-2 text-gray-800">Upload Template</h2>
                        <p className="text-gray-500 mb-6">Select the completed Excel file with your leave balance adjustments. Make sure to use the provided adjustment template generated from this app. Adjustments will be read from the file and prepared for your review before the update process starts.</p>
                        
                        {uploadConflicts && (
                            <div className="mb-6 bg-red-50 border-l-4 border-red-500 p-4 rounded-md">
                                <div className="flex">
                                    <div className="flex-shrink-0">
                                        <ExclamationIcon className="h-5 w-5 text-red-400" />
                                    </div>
                                    <div className="ml-3 w-full">
                                        <h3 className="text-sm font-bold text-red-800">CRITICAL ERROR: Ambiguous or Conflicting Date Formats</h3>
                                        <p className="text-sm text-red-700 mt-1">
                                            The app found conflicting date formats (US vs EU) within the same column or between Validity columns. We cannot proceed because we don't know which date is correct.
                                        </p>
                                        <div className="mt-3 text-sm">
                                            {uploadConflicts.map((conflict, i) => (
                                                <div key={i} className="mb-2">
                                                    <p className="font-semibold text-red-900">Column: {conflict.column}</p>
                                                    <ul className="list-disc list-inside ml-2 text-red-800">
                                                        {conflict.details.map((msg, j) => <li key={j}>{msg}</li>)}
                                                    </ul>
                                                </div>
                                            ))}
                                        </div>
                                        <p className="text-sm text-red-800 mt-2 font-bold mb-3">Please fix these rows in your Excel file and re-upload.</p>
                                        
                                        <button 
                                            onClick={() => {
                                                setUploadConflicts(null);
                                                setCurrentStep('configure');
                                            }}
                                            className="bg-red-600 hover:bg-red-700 text-white font-bold py-2 px-4 rounded-md text-sm transition-colors shadow-sm"
                                        >
                                            Try again
                                        </button>
                                    </div>
                                </div>
                            </div>
                        )}

                        {uploadValidityErrors && (
                            <div className="mb-6 bg-red-50 border-l-4 border-red-500 p-4 rounded-md">
                                <div className="flex">
                                    <div className="flex-shrink-0">
                                        <ExclamationIcon className="h-5 w-5 text-red-400" />
                                    </div>
                                    <div className="ml-3 w-full">
                                        <h3 className="text-sm font-bold text-red-800">Data Validation Errors</h3>
                                        <p className="text-sm text-red-700 mt-1">
                                            The uploaded file contains rows with missing or invalid dates.
                                        </p>
                                        <div className="mt-3 text-sm max-h-48 overflow-y-auto">
                                            {uploadValidityErrors.map((err, i) => (
                                                <div key={i} className="mb-2">
                                                    <p className="font-semibold text-red-900">Row {err.row}</p>
                                                    <ul className="list-disc list-inside ml-2 text-red-800">
                                                        {err.details.map((msg, j) => <li key={j}>{msg}</li>)}
                                                    </ul>
                                                </div>
                                            ))}
                                        </div>
                                        <p className="text-sm text-red-800 mt-2 font-bold mb-3">Please correct these cells in your Excel file and re-upload.</p>
                                        
                                        <button 
                                            onClick={() => {
                                                setUploadValidityErrors(null);
                                                setCurrentStep('configure');
                                            }}
                                            className="bg-red-600 hover:bg-red-700 text-white font-bold py-2 px-4 rounded-md text-sm transition-colors shadow-sm"
                                        >
                                            Try again
                                        </button>
                                    </div>
                                </div>
                            </div>
                        )}
                        
                        {/* File input is always present but hidden */}
                        <input 
                            id="file-upload-input"
                            type="file" 
                            onChange={handleFileInputChange} 
                            accept=".xlsx, .xls" 
                            className="hidden" 
                        />

                        {!uploadConflicts && !uploadValidityErrors && <div 
                            className={`border-2 border-dashed rounded-lg p-12 text-center cursor-pointer transition-colors ${isDragging ? 'border-blue-500 bg-blue-50' : 'border-gray-300 hover:border-gray-400'}`}
                            onDragOver={handleDragOver}
                            onDragLeave={handleDragLeave}
                            onDrop={handleDrop}
                            onClick={() => document.getElementById('file-upload-input')?.click()}
                        >
                            <div className="flex flex-col items-center justify-center">
                                 <UploadCloudIcon className={`w-12 h-12 mb-4 ${isDragging ? 'text-blue-500' : 'text-gray-400'}`} />
                                 <p className="text-gray-600 font-medium">Drop file or click to browse</p>
                                 <p className="text-gray-400 text-sm mt-2">Supports .xlsx, .xls</p>
                            </div>
                        </div>}

                        <div className="mt-8 pt-6 border-t border-gray-200 flex justify-end"><button onClick={() => setCurrentStep('configure')} className="bg-white hover:bg-gray-100 text-gray-700 font-semibold py-2 px-4 rounded-md border border-gray-300 shadow-sm">&larr; Back</button></div>
                    </div>}
                    
                    {currentStep === 'review' && <div className="bg-white px-4 py-8 md:px-8 border-x border-b shadow-md w-full max-w-[1400px] mx-auto rounded-lg">
                         <div className="flex items-center gap-2 mb-2">
                             <h2 className="text-2xl font-bold text-gray-800">Review & Update Balances</h2>
                             <div className="relative group flex items-center">
                                 <InfoIcon className="h-6 w-6 text-orange-500 cursor-help" />
                                 <div className="absolute left-full top-1/2 transform -translate-y-1/2 ml-2 w-64 p-2.5 bg-gray-900 text-white text-sm font-normal rounded shadow-lg opacity-0 invisible group-hover:opacity-100 group-hover:visible transition-opacity z-20 pointer-events-none">
                                     Note: Balance updates are processed as adjustments. Therefore, the adjustment amount will not equal the final available balance; instead, it functions as an addition or deduction to the existing amount.
                                 </div>
                             </div>
                         </div>
                         
                         {/* Validation Errors Alert */}
                         {(hasAttemptedSubmit && (hasValidityErrors || hasMissingDates)) && (
                             <div className="mb-6 bg-red-50 border-l-4 border-red-500 p-4">
                                <div className="flex">
                                    <div className="flex-shrink-0">
                                        <ExclamationIcon className="h-5 w-5 text-red-400" />
                                    </div>
                                    <div className="ml-3">
                                        <p className="text-sm text-red-700 font-bold">
                                            Validation Errors Detected
                                        </p>
                                        <ul className="list-disc ml-5 mt-1 text-sm text-red-600">
                                            {hasValidityErrors && <li>Some Effective Dates are outside the Account's validity period.</li>}
                                            {(hasAttemptedSubmit && hasMissingDates) && <li>Some rows are missing Effective Dates.</li>}
                                        </ul>
                                        <p className="text-sm text-red-600 mt-2">
                                            Please correct the highlighted dates below or remove these rows before continuing.
                                        </p>
                                    </div>
                                </div>
                             </div>
                         )}

                         {/* Date Format Detection Report - Conditionally Shown */}
                         {shouldShowDateReport && detectedColumnFormats && (
                             <div className="mb-6 bg-blue-50 border-l-4 border-blue-500 p-4">
                                <div className="flex flex-col">
                                    <div className="flex items-center mb-2">
                                        <InfoIcon className="h-5 w-5 text-blue-400 mr-2" />
                                        <p className="text-sm text-blue-700 font-bold">
                                            Date Format Report
                                        </p>
                                    </div>
                                    <p className="text-sm text-blue-600 mb-3">
                                        We found ambiguous dates or formats differing from your preference. 
                                        We <strong>automatically converted</strong> these dates to match your app setting.
                                    </p>
                                    
                                    {dateReportExample && (
                                        <>
                                            <p className="text-sm text-blue-700 font-bold mb-2">Detected Ambiguous Date and Conversion Example</p>

                                            <div className="mb-4 p-3 bg-white rounded border border-blue-200 text-xs text-gray-700 shadow-sm space-y-2">
                                                <div>
                                                    <span className="font-semibold text-blue-800">Location:</span> Row {dateReportExample.rowNumber} <br/>
                                                    <span className="text-gray-500">Employee: {dateReportExample.employee} | Account: {dateReportExample.account}</span>
                                                </div>
                                                <div>
                                                    <span className="font-semibold text-blue-800">Detection:</span> Found value <code className="bg-gray-100 px-1 py-0.5 rounded font-bold">{dateReportExample.rawValue}</code> in column <em>{dateReportExample.columnName}</em>.
                                                </div>
                                                <div>
                                                    <span className="font-semibold text-blue-800">Conversion:</span> Because your app setting is <strong>{downloadDateFormat}</strong>, we converted the full column to match this format:
                                                    <div className="mt-1 flex items-center gap-2">
                                                        <span className="line-through text-gray-400">{dateReportExample.rawValue}</span>
                                                        <span>&rarr;</span>
                                                        <span className="bg-green-100 text-green-800 px-2 py-0.5 rounded font-bold border border-green-200">{dateReportExample.convertedValue}</span>
                                                    </div>
                                                </div>
                                            </div>
                                        </>
                                    )}

                                    <div className="flex flex-col md:flex-row gap-4 items-start md:items-center justify-between">
                                        <div className="grid grid-cols-1 md:grid-cols-2 gap-2 text-sm text-blue-800 bg-blue-100 p-2 rounded w-full md:w-auto">
                                            <div className="flex justify-between border-b md:border-b-0 md:border-r border-blue-200 px-2 gap-3 items-center">
                                                <span className="font-semibold">Validity Period:</span>
                                                {renderFormatLabel(detectedColumnFormats.validFrom, "Validity Period Columns")}
                                            </div>
                                            <div className="flex justify-between px-2 gap-3 items-center">
                                                <span className="font-semibold">Effective Date:</span>
                                                {renderFormatLabel(detectedColumnFormats.effective, "Effective Date Column")}
                                            </div>
                                        </div>
                                        
                                        <div className="flex items-center gap-2">
                                            <span className="text-xs font-semibold text-gray-500">Current Format Setting: <span className="text-gray-800">{downloadDateFormat}</span></span>
                                            <button 
                                                onClick={handleSwitchDateFormat}
                                                className="whitespace-nowrap bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold py-2 px-3 rounded shadow transition-colors flex items-center"
                                            >
                                                <RefreshIcon className="h-3 w-3 mr-1" />
                                                Switch to {downloadDateFormat === 'EU' ? 'US' : 'EU'}
                                            </button>
                                        </div>
                                    </div>
                                </div>
                             </div>
                         )}

                         {/* Filters and Bulk Actions */}
                         <div className="mb-6 space-y-6">
                            {/* Table Filters */}
                            <div className="bg-gray-50 p-4 rounded border border-gray-200">
                                <h3 className="text-sm font-semibold text-gray-700 mb-3 flex items-center">
                                    <FilterIcon className="h-4 w-4 mr-2" /> Filter Accounts in Table
                                </h3>
                                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
                                    <div>
                                        <div className="flex justify-between items-center mb-2">
                                            <label className="block text-xs font-semibold text-gray-500 uppercase tracking-wider">
                                                {usePrimaryDepartmentOnly ? 'PRIMARY DEPARTMENT' : 'DEPARTMENT'}
                                            </label>
                                            <div className="flex items-center space-x-1 group relative">
                                                <button 
                                                    onClick={() => setUsePrimaryDepartmentOnly(!usePrimaryDepartmentOnly)}
                                                    className={`relative inline-flex h-4 w-7 items-center rounded-full transition-colors ${usePrimaryDepartmentOnly ? 'bg-orange-500' : 'bg-gray-300'}`}
                                                >
                                                    <span className={`inline-block h-3 w-3 transform rounded-full bg-white transition-transform ${usePrimaryDepartmentOnly ? 'translate-x-3.5' : 'translate-x-0.5'}`} />
                                                </button>
                                                <InfoIcon className="h-4 w-4 text-orange-400 hover:text-orange-500 cursor-help" />
                                                <div className="absolute bottom-full left-1/2 transform -translate-x-1/2 mb-2 w-64 p-3 bg-gray-900 text-white text-xs rounded-md shadow-xl opacity-0 invisible group-hover:opacity-100 group-hover:visible transition-all duration-200 z-50 text-left leading-relaxed">
                                                    <p className="mb-1">You can filter by Primary Department.</p>
                                                    <p>Note: Employees without a designated primary department will not appear when this is active. To use this filter, ensure the feature is enabled and assigned to everyone.</p>
                                                </div>
                                            </div>
                                        </div>
                                        <MultiSelectDropdown 
                                            label={usePrimaryDepartmentOnly ? "Primary Department" : "Department"} 
                                            pluralLabel={usePrimaryDepartmentOnly ? "Primary Departments" : "Departments"}
                                            options={departments} 
                                            selectedIds={reviewDepartmentIds} 
                                            onChange={setReviewDepartmentIds} 
                                        />
                                    </div>
                                    <div>
                                        <label className="block text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">EMPLOYEE GROUP</label>
                                        <MultiSelectDropdown 
                                            label="Employee Group" 
                                            pluralLabel="Employee groups"
                                            options={employeeGroups} 
                                            selectedIds={reviewEmployeeGroupIds} 
                                            onChange={setReviewEmployeeGroupIds} 
                                        />
                                    </div>
                                    <div>
                                        <label className="block text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">EMPLOYEE TYPE</label>
                                        <MultiSelectDropdown 
                                            label="Employee Type" 
                                            pluralLabel="Employee types"
                                            options={employeeTypes} 
                                            selectedIds={reviewEmployeeTypeIds} 
                                            onChange={setReviewEmployeeTypeIds} 
                                        />
                                    </div>
                                    <div>
                                        <label className="block text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">LEAVE ACCOUNT</label>
                                        <MultiSelectDropdown 
                                            label="Leave Account Name" 
                                            pluralLabel="Leave account names"
                                            options={uniqueAccountNames} 
                                            selectedIds={reviewLeaveAccountNames} 
                                            onChange={setReviewLeaveAccountNames} 
                                        />
                                    </div>
                                </div>
                            </div>

                            {/* Bulk Edit Panel */}
                            <div className="bg-white p-4 rounded border border-gray-200 shadow-sm relative pt-6">
                                <h3 className="text-sm font-semibold text-gray-800 absolute top-4 left-4 flex items-center mb-0">
                                    BULK EDIT
                                    <div 
                                        className="flex items-center ml-1 cursor-help"
                                        onClick={(e) => e.stopPropagation()}
                                        onMouseEnter={(e) => handleTooltipEnter(e, (
                                            <div className="text-left w-72 normal-case">
                                                Use the checkboxes in the table below to choose which employee accounts to include for the bulk edit. Use the filters to easily find specific employees within Departments, Employee Groups, and Types.
                                            </div>
                                        ))}
                                        onMouseLeave={handleTooltipLeave}
                                    >
                                        <InfoIcon className="h-4 w-4 text-gray-400 hover:text-blue-500" />
                                    </div>
                                </h3>
                                
                                <div className="mt-8 flex flex-col lg:flex-row gap-6 items-start lg:items-end w-full">
                                    <div className="w-full lg:w-1/4">
                                        <label className="block text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">CHANGE VALUE FOR</label>
                                        <select 
                                            value={bulkEditField} 
                                            onChange={(e) => setBulkEditField(e.target.value as any)}
                                            className="w-full rounded border-gray-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm"
                                        >
                                            <option value="">-- Select Field --</option>
                                            {shouldShowBalanceColumns && <option value="newBalance">New Balance</option>}
                                            <option value="adjustment">Adjustment</option>
                                            <option value="effectiveDate">Effective Date</option>
                                            <option value="comment">Comment</option>
                                        </select>
                                    </div>
                                    
                                    <div className="w-full lg:w-1/4">
                                        <label className="block text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">NEW VALUE</label>
                                        {bulkEditField === 'adjustment' && (
                                            <input 
                                                type="number"
                                                step="any"
                                                value={bulkEditValueAdjustment}
                                                onChange={(e) => setBulkEditValueAdjustment(e.target.value)}
                                                placeholder="Value to apply..."
                                                className="w-full rounded border-gray-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm"
                                            />
                                        )}
                                        {bulkEditField === 'newBalance' && (
                                            <input 
                                                type="number"
                                                step="any"
                                                value={bulkEditValueNewBalance}
                                                onChange={(e) => setBulkEditValueNewBalance(e.target.value)}
                                                placeholder="Value to apply..."
                                                className="w-full rounded border-gray-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm"
                                            />
                                        )}
                                        {bulkEditField === 'effectiveDate' && (
                                            <div className="flex flex-col gap-2">
                                                <select
                                                    value={bulkEditEffectiveDateType}
                                                    onChange={(e) => setBulkEditEffectiveDateType(e.target.value as any)}
                                                    className="w-full rounded border-gray-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm"
                                                >
                                                    <option value="today">Today</option>
                                                    <option value="start_date">Account Start Date</option>
                                                    <option value="custom">Custom Date</option>
                                                </select>
                                                {bulkEditEffectiveDateType === 'custom' && (
                                                    <input 
                                                        type="date"
                                                        value={bulkEditValueEffectiveDate}
                                                        onChange={(e) => setBulkEditValueEffectiveDate(e.target.value)}
                                                        className="w-full rounded border-gray-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm"
                                                    />
                                                )}
                                            </div>
                                        )}
                                        {bulkEditField === 'comment' && (
                                            <input 
                                                type="text"
                                                value={bulkEditValueComment}
                                                onChange={(e) => setBulkEditValueComment(e.target.value)}
                                                placeholder="Comment..."
                                                className="w-full rounded border-gray-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm"
                                            />
                                        )}
                                        {!bulkEditField && (
                                            <input 
                                                type="text"
                                                placeholder="Value to apply..."
                                                disabled
                                                className="w-full rounded border-gray-300 shadow-sm bg-gray-50 sm:text-sm opacity-60"
                                            />
                                        )}
                                    </div>

                                    <div className="w-full lg:w-1/2 flex flex-col gap-2">
                                        <button 
                                            onClick={handleApplyBulkEdit}
                                            disabled={!bulkEditField || selectedReviewIds.size === 0}
                                            className="w-full bg-blue-400 hover:bg-blue-500 text-white font-semibold py-2 px-4 rounded shadow disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
                                        >
                                            Apply to {selectedReviewIds.size} Selected
                                        </button>
                                        <div className="flex gap-2">
                                            <button onClick={handleSelectFiltered} className="flex-1 px-2 py-1 text-xs border border-gray-300 bg-white hover:bg-gray-50 rounded">Select only the filtered accounts in the table</button>
                                            <button onClick={handleSelectAll} className="flex-1 px-2 py-1 text-xs border border-gray-300 bg-white hover:bg-gray-50 rounded">Select all accounts in the table</button>
                                        </div>
                                    </div>
                                </div>
                            </div>
                         </div>

                         {/* Table Toolbar */}
                         <div className="mb-4 flex flex-col xl:flex-row justify-between items-start xl:items-center gap-4">
                            <div className="flex items-center gap-3">
                                <h3 className="text-xl font-bold text-gray-800">Pending Adjustments</h3>
                                <div className="bg-gray-100 rounded-full px-3 py-1 flex items-center justify-center gap-1">
                                    <span className="text-orange-500 text-sm">
                                        <span className="font-bold">{adjustmentsToReview.filter(a => !a.status || a.status === 'pending').length}</span> accounts with updates pending
                                    </span>
                                    <div 
                                        className="flex items-center cursor-help"
                                        onClick={(e) => e.stopPropagation()}
                                        onMouseEnter={(e) => handleTooltipEnter(e, (
                                            <div className="text-left w-64 normal-case text-xs text-white">
                                                All accounts with updates pending will be updated upon confirmation, not just the filtered/selected ones. Values can be changed via manual individual input or via the bulk edit option.
                                            </div>
                                        ))}
                                        onMouseLeave={handleTooltipLeave}
                                    >
                                        <InfoIcon className="h-4 w-4 text-orange-400 hover:text-orange-600" />
                                    </div>
                                </div>
                                {(hasValidationErrors && hasAttemptedSubmit) && (
                                    <button
                                        onClick={handleDeleteAllErrors}
                                        className="text-sm bg-red-100 text-red-800 hover:bg-red-200 px-3 py-1.5 rounded font-semibold transition-colors flex items-center gap-1 ml-2"
                                    >
                                        <TrashIcon className="h-4 w-4" />
                                        Delete all rows with errors
                                    </button>
                                )}
                            </div>

                            <div className="flex items-center gap-2">
                                <div className="relative mr-2 border border-gray-300 rounded shadow-sm overflow-hidden flex items-center bg-white">
                                    <svg xmlns="http://www.w3.org/2000/svg" className="h-4 w-4 ml-2 text-gray-400" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" /></svg>
                                    <input 
                                        type="text"
                                        placeholder="Search Employee..."
                                        value={searchQuery}
                                        onChange={(e) => setSearchQuery(e.target.value)}
                                        className="pl-2 pr-3 py-1.5 text-xs text-gray-700 border-none focus:ring-0 focus:outline-none w-48"
                                    />
                                </div>

                                <span className="text-xs font-semibold text-gray-400 uppercase">Sort:</span>
                                <button onClick={() => handleSort('employeeName')} className={`px-3 py-1.5 text-xs font-medium border rounded transition-colors ${sortConfig?.key === 'employeeName' ? 'bg-blue-50 border-blue-300 text-blue-700' : 'bg-white border-gray-300 text-gray-700 hover:bg-gray-50'}`}>Employee</button>
                                <button onClick={() => handleSort('accountName')} className={`px-3 py-1.5 text-xs font-medium border rounded transition-colors ${sortConfig?.key === 'accountName' ? 'bg-blue-50 border-blue-300 text-blue-700' : 'bg-white border-gray-300 text-gray-700 hover:bg-gray-50'}`}>Account</button>
                                <button onClick={() => handleSort('validity')} className={`px-3 py-1.5 text-xs font-medium border rounded transition-colors ${sortConfig?.key === 'validity' ? 'bg-blue-50 border-blue-300 text-blue-700' : 'bg-white border-gray-300 text-gray-700 hover:bg-gray-50'}`}>Validity</button>
                                <button onClick={() => handleSort('status')} className={`px-3 py-1.5 text-xs font-medium border rounded transition-colors ${sortConfig?.key === 'status' ? 'bg-blue-50 border-blue-300 text-blue-700' : 'bg-white border-gray-300 text-gray-700 hover:bg-gray-50'}`}>Errors</button>
                                
                                <div className="flex items-center gap-2 pl-2">
                                    <button
                                        onClick={undoAdjustments}
                                        disabled={pastLength === 0}
                                        className={`px-3 py-1.5 text-sm font-medium border rounded-md flex items-center justify-center gap-1.5 transition-colors ${pastLength > 0 ? 'bg-white border-slate-300 text-slate-500 hover:bg-slate-50 hover:text-slate-600' : 'bg-white border-slate-200 text-slate-300 cursor-not-allowed'}`}
                                        title="Undo"
                                    >
                                        <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                                            <path d="M9 14 4 9l5-5" />
                                            <path d="M4 9h10.5a5.5 5.5 0 0 1 5.5 5.5v1.5" />
                                        </svg>
                                        Undo
                                    </button>
                                    <button
                                        onClick={redoAdjustments}
                                        disabled={futureLength === 0}
                                        className={`px-3 py-1.5 text-sm font-medium border rounded-md flex items-center justify-center gap-1.5 transition-colors ${futureLength > 0 ? 'bg-white border-slate-300 text-slate-500 hover:bg-slate-50 hover:text-slate-600' : 'bg-white border-slate-200 text-slate-300 cursor-not-allowed'}`}
                                        title="Redo"
                                    >
                                        <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                                            <path d="M15 14l5-5-5-5" />
                                            <path d="M20 9H9.5A5.5 5.5 0 0 0 4 14.5v1.5" />
                                        </svg>
                                        Redo
                                    </button>
                                </div>
                            </div>
                         </div>

                         <div className="max-h-[800px] overflow-y-auto border rounded-lg shadow-sm">
                            <table className="w-full text-sm text-left relative">
                                <thead className="text-xs text-gray-700 uppercase bg-gray-50 sticky top-0 z-10 shadow-sm">
                                    <tr>
                                        <th scope="col" className="px-4 py-3 w-10">
                                            <input 
                                                type="checkbox" 
                                                className="rounded border-gray-300 text-blue-600 focus:ring-blue-500"
                                                checked={sortedReviews.length > 0 && Array.from(selectedReviewIds).some(id => sortedReviews.some(r => r.id === id))}
                                                onChange={(e) => {
                                                    if (e.target.checked) handleSelectFiltered();
                                                    else setSelectedReviewIds(new Set());
                                                }}
                                            />
                                        </th>
                                        <th scope="col" className="px-4 py-3 align-middle text-left">
                                            <div className="flex flex-col items-start justify-center h-full gap-2">
                                                <div className="flex items-center gap-1 cursor-pointer hover:bg-gray-100 transition-colors group px-2 py-1 rounded" onClick={() => handleSort('employeeName')}>
                                                    Employee ({sortedReviews.length}/{adjustmentsToReview.length})
                                                    <div 
                                                        className="flex items-center cursor-help"
                                                        onClick={(e) => e.stopPropagation()}
                                                        onMouseEnter={(e) => handleTooltipEnter(e, (
                                                            <div className="text-left w-64 normal-case">
                                                                The first number in the bracket is the current filtered number, and the second number is the total number of employee accounts included for updates.
                                                            </div>
                                                        ))}
                                                        onMouseLeave={handleTooltipLeave}
                                                    >
                                                        <InfoIcon className="h-4 w-4 text-gray-400 hover:text-blue-500 ml-1" />
                                                    </div>
                                                    {sortConfig?.key === 'employeeName' && (sortConfig.direction === 'asc' ? <SortAscIcon className="h-3 w-3"/> : <SortDescIcon className="h-3 w-3"/>)}
                                                </div>
                                                {selectedReviewIds.size > 0 && (
                                                    <div className="flex gap-2 w-full justify-start">
                                                        <button 
                                                            onClick={(e) => { e.stopPropagation(); handleClearSelection(); }}
                                                            className="text-xs px-2 py-1 rounded border border-gray-300 bg-white hover:bg-gray-100 text-gray-700"
                                                        >
                                                            Clear
                                                        </button>
                                                        <button 
                                                            onClick={(e) => { e.stopPropagation(); handleRemoveSelected(); }}
                                                            className="text-xs px-2 py-1 rounded border border-red-300 bg-red-50 hover:bg-red-100 text-red-700"
                                                        >
                                                            Remove
                                                        </button>
                                                    </div>
                                                )}
                                            </div>
                                        </th>
                                        <th scope="col" className="px-4 py-3 align-middle text-left cursor-pointer hover:bg-gray-100 transition-colors group" onClick={() => handleSort('accountName')}>
                                            <div className="flex items-center justify-start gap-1">
                                                Account
                                                {sortConfig?.key === 'accountName' && (sortConfig.direction === 'asc' ? <SortAscIcon className="h-3 w-3"/> : <SortDescIcon className="h-3 w-3"/>)}
                                            </div>
                                        </th>
                                        <th scope="col" className="px-4 py-3 align-middle text-left">Validity</th>
                                        <th scope="col" className="px-2 py-3 align-middle text-left w-16">Type</th>
                                        {shouldShowBalanceColumns && (
                                            <>
                                                <th scope="col" className="px-4 py-3 align-middle text-left">
                                                    <div className="flex flex-col items-start justify-center">
                                                        <div className="flex items-center justify-start gap-1">
                                                            Avail. Balance
                                                            <div 
                                                                className="flex items-center cursor-help"
                                                                onClick={(e) => e.stopPropagation()}
                                                                onMouseEnter={(e) => handleTooltipEnter(e, (
                                                                    <div className="text-left w-72 normal-case text-sm space-y-3">
                                                                        {updateMethod === 'excel' ? (
                                                                            <>
                                                                                <p>The available balance as of the provided Balance Date.</p>
                                                                                <p className="text-orange-400">Note: This displayed balance may not reflect recent updates (e.g., approved leave, accruals, manual adjustments, etc.) processed after the template file was generated. Click "Update available balance" to ensure that they match.</p>
                                                                            </>
                                                                        ) : (
                                                                            <p>The available balance as of the provided date.</p>
                                                                        )}
                                                                        <p className="italic font-semibold pt-1 text-xs">More info:</p>
                                                                        <ul className="list-disc pl-4 space-y-2 text-xs">
                                                                            <li><strong>Accounts with Accruals:</strong> If you select a date that is earlier than the account end date, the file will only show the balance accrued up to that specific date. Future accruals will not be included in the available balance.</li>
                                                                            <li><strong>Past Leave:</strong> All approved leave requests BEFORE this date are already deducted from the displayed available balance.</li>
                                                                            <li><strong>Future Leave:</strong> Any approved leave requests AFTER this date are NOT deducted from the displayed available balance.</li>
                                                                        </ul>
                                                                    </div>
                                                                ))}
                                                                onMouseLeave={handleTooltipLeave}
                                                            >
                                                                <InfoIcon className="h-4 w-4 text-gray-400 hover:text-blue-500 ml-1" />
                                                            </div>
                                                        </div>
                                                        {singleBalanceDate && (
                                                            <div className="text-[10px] text-gray-500 font-normal mt-0.5">
                                                                As of: {formatDateForDisplay(singleBalanceDate, downloadDateFormat)}
                                                            </div>
                                                        )}
                                                        {updateMethod === 'excel' && (
                                                            <button 
                                                                onClick={(e) => { e.stopPropagation(); setShowUpdateBalanceModal(true); }}
                                                                className="mt-1 px-2 py-0.5 text-[10px] font-bold bg-orange-100 text-orange-700 hover:bg-orange-200 rounded border border-orange-200 transition-colors"
                                                            >
                                                                UPDATE
                                                            </button>
                                                        )}
                                                    </div>
                                                </th>
                                                <th scope="col" className="px-4 py-3 align-middle text-left">
                                                    <div className="flex items-center justify-start gap-1">
                                                        New Balance
                                                        <div 
                                                            className="flex items-center cursor-help"
                                                            onClick={(e) => e.stopPropagation()}
                                                            onMouseEnter={(e) => handleTooltipEnter(e, (
                                                                <div className="text-left w-64 normal-case">
                                                                    The final expected balance. Changing this value will automatically auto-calculate and populate the Adj. column value.
                                                                </div>
                                                            ))}
                                                            onMouseLeave={handleTooltipLeave}
                                                        >
                                                            <InfoIcon className="h-4 w-4 text-gray-400 hover:text-blue-500 ml-1" />
                                                        </div>
                                                    </div>
                                                </th>
                                            </>
                                        )}
                                        <th scope="col" className="px-2 py-3 align-middle text-left cursor-pointer hover:bg-gray-100 transition-colors group w-20" onClick={() => handleSort('adjustment')}>
                                            <div className="flex items-center justify-start gap-1">
                                                Adj.
                                                <div 
                                                    className="flex items-center cursor-help"
                                                    onClick={(e) => e.stopPropagation()}
                                                    onMouseEnter={(e) => handleTooltipEnter(e, (
                                                        <div className="text-left w-64 normal-case">
                                                            Adjustments. Note: Balance updates are processed as adjustments. Therefore, the adjustment amount will not equal the final available balance; instead, it functions as an addition or deduction to the existing amount.
                                                        </div>
                                                    ))}
                                                    onMouseLeave={handleTooltipLeave}
                                                >
                                                    <InfoIcon className="h-4 w-4 text-gray-400 hover:text-blue-500 ml-1" />
                                                </div>
                                                {sortConfig?.key === 'adjustment' && (sortConfig.direction === 'asc' ? <SortAscIcon className="h-3 w-3"/> : <SortDescIcon className="h-3 w-3"/>)}
                                            </div>
                                        </th>
                                        <th scope="col" className="px-4 py-3 align-middle text-left">
                                            <div className="flex items-center justify-start gap-1">
                                                Effective Date
                                                <div 
                                                    className="flex items-center cursor-help"
                                                    onClick={(e) => e.stopPropagation()}
                                                    onMouseEnter={(e) => handleTooltipEnter(e, (
                                                        <div className="text-left w-80 normal-case">
                                                            <p className="mb-2">The date you choose will be the effective date for the assigned balance adjustment. This means the balance will only be available for the employee to use on or after this date, so please ensure you select it carefully.</p>
                                                            <p className="mb-2">If you want the balance adjustment to be available from the very beginning of the period (essentially acting as a starting balance - if no other balance has been assigned/accrued), then select the Account Start Date.</p>
                                                            <p>Note: The date picker below displays dates according to your browser settings, not the app settings. However, the system will automatically convert the date to the correct format before performing the leave balance update.</p>
                                                        </div>
                                                    ))}
                                                    onMouseLeave={handleTooltipLeave}
                                                >
                                                    <InfoIcon className="h-4 w-4 text-gray-400 hover:text-blue-500 ml-1" />
                                                </div>
                                            </div>
                                        </th>
                                        <th scope="col" className="px-4 py-3 align-middle text-left">
                                            <div className="flex items-center justify-start gap-1">
                                                Comment
                                                <div 
                                                    className="flex items-center cursor-help"
                                                    onClick={(e) => e.stopPropagation()}
                                                    onMouseEnter={(e) => handleTooltipEnter(e, (
                                                        <div className="text-left w-72 normal-case">
                                                            <p>Comments are optional and can be added from the excel file. Note, the following text is ALWAYS SENT as a comment, whether a comment is entered or not: API BULK UPDATE.</p>
                                                        </div>
                                                    ))}
                                                    onMouseLeave={handleTooltipLeave}
                                                >
                                                    <InfoIcon className="h-4 w-4 text-gray-400 hover:text-blue-500" />
                                                </div>
                                            </div>
                                        </th>
                                    </tr>
                                </thead>
                                <tbody>{paginatedReviews.map((adj, idx) => {
                                    const isMissingEffectiveDate = !adj.effectiveDate;
                                    const isRowError = hasAttemptedSubmit && (adj.isValidationError || isMissingEffectiveDate);
                                    return (
                                    <tr key={adj.id} className={`border-b hover:bg-gray-50 ${isRowError ? 'bg-red-50' : 'bg-white'}`}>
                                        <td className="px-4 py-3 text-center">
                                            <input 
                                                type="checkbox" 
                                                className="rounded border-gray-300 text-blue-600 focus:ring-blue-500"
                                                checked={selectedReviewIds.has(adj.id)}
                                                onChange={() => toggleReviewSelection(adj.id)}
                                            />
                                        </td>
                                        <td className="px-4 py-3 font-medium text-gray-900 leading-tight align-middle">
                                            <div>{adj.employeeName}</div>
                                            {adj.salaryIdentifier && <div className="text-xs font-mono text-gray-500 mt-0.5">SID: {adj.salaryIdentifier}</div>}
                                        </td>
                                        <td className="px-4 py-3 align-middle">{adj.accountName}</td>
                                        <td className="px-4 py-3 text-xs text-gray-500 whitespace-nowrap align-middle">
                                            {/* Use user's downloadDateFormat preference */}
                                            {adj.validFrom ? formatDateForDisplay(adj.validFrom, downloadDateFormat) : 'N/A'} - <br/>
                                            {adj.validTo ? formatDateForDisplay(adj.validTo, downloadDateFormat) : '∞'}
                                        </td>
                                        <td className="px-2 py-3 text-xs text-gray-700 whitespace-nowrap align-middle">
                                            {adj.unit || 'N/A'}
                                        </td>
                                        {shouldShowBalanceColumns && (
                                            <>
                                                <td className="px-4 py-3 align-middle text-left bg-gray-50/50">
                                                    <div className="font-mono font-medium text-gray-800">
                                                        {typeof adj.availableBalance === 'number' ? adj.availableBalance.toFixed(2) : adj.availableBalance}
                                                    </div>
                                                </td>
                                                <td className="px-4 py-3 align-middle">
                                                    <div className="flex justify-start items-center w-full h-full">
                                                        <EditableAdjustmentCell 
                                                            value={typeof adj.newBalance === 'number' ? adj.newBalance : 0} 
                                                            onChange={(val) => handleUpdateNewBalance(adj.id, val)}
                                                            disabled={isLoading.submitting}
                                                            isBalance={true}
                                                        />
                                                    </div>
                                                </td>
                                            </>
                                        )}
                                        <td 
                                            className={`px-2 py-3 align-middle ${adj.adjustment === 0 ? 'bg-orange-50' : ''}`}
                                            title={adj.adjustment === 0 && !hasAttemptedSubmit ? "An adjustment cannot be 0 and therefore, this adjustment will get skipped if you don't change it to a positive/negative value." : undefined}
                                        >
                                            <div className="flex justify-start items-center w-full h-full">
                                                <EditableAdjustmentCell 
                                                    value={adj.adjustment} 
                                                    onChange={(val) => handleUpdateAdjustment(adj.id, val)}
                                                    disabled={isLoading.submitting}
                                                />
                                            </div>
                                        </td>
                                        {/* Edit Date Input */}
                                        <td className="px-4 py-3 font-mono text-gray-600 align-middle">
                                            <div className="flex flex-col justify-center items-center w-full h-full">
                                                <input 
                                                    type="date" 
                                                    value={adj.effectiveDate} 
                                                    onChange={(e) => handleUpdateEffectiveDate(adj.id, e.target.value)}
                                                    className={`text-sm border rounded px-2 py-1 w-full max-w-[140px] ${(hasAttemptedSubmit && (!adj.effectiveDate || adj.isValidationError)) ? 'border-red-500 ring-1 ring-red-500 text-red-700' : 'border-gray-300'}`}
                                                    disabled={isLoading.submitting}
                                                />
                                                {(!adj.effectiveDate && hasAttemptedSubmit) && <p className="text-xs font-semibold text-red-600 mt-1">Date Required</p>}
                                                {(adj.isValidationError && adj.effectiveDate && hasAttemptedSubmit) && <p className="text-xs text-red-600 mt-1">{adj.error}</p>}
                                            </div>
                                        </td>
                                        <td className="px-4 py-3 max-w-[320px] align-middle">
                                            <div className="flex justify-center items-center w-full h-full">
                                                <input 
                                                    type="text" 
                                                    className="w-full bg-transparent border border-gray-300 rounded focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 text-sm px-2 py-1"
                                                    value={adj.comment || ''} 
                                                    onChange={(e) => handleUpdateComment(adj.id, e.target.value)}
                                                    placeholder="Add optional comment..."
                                                    disabled={isLoading.submitting}
                                                />
                                            </div>
                                        </td>
                                    </tr>);
                                })}
                                </tbody>
                            </table>
                        </div>
                        
                        {/* Pagination Footer */}
                        <div className="flex items-center justify-start gap-6 mt-4 text-sm text-slate-700">
                            <div className="flex items-center border border-gray-300 rounded-md shadow-sm bg-white h-9">
                                <button 
                                    onClick={() => setReviewPage(p => Math.max(1, p - 1))}
                                    disabled={reviewPage === 1}
                                    className="px-3 h-full hover:bg-gray-50 border-r border-gray-300 disabled:opacity-30 disabled:cursor-not-allowed flex items-center justify-center text-slate-500"
                                >
                                    <svg xmlns="http://www.w3.org/2000/svg" className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15 19l-7-7 7-7" /></svg>
                                </button>
                                <span className="px-4 h-full flex items-center font-bold whitespace-nowrap min-w-[7rem] justify-center text-slate-800">
                                    Page {reviewPage} of {reviewRowsPerPage === 'All' ? 1 : Math.ceil(sortedReviews.length / reviewRowsPerPage) || 1}
                                </span>
                                <button 
                                    onClick={() => setReviewPage(p => Math.min(Math.ceil(sortedReviews.length / (reviewRowsPerPage as number)), p + 1))}
                                    disabled={reviewRowsPerPage === 'All' || reviewPage === Math.ceil(sortedReviews.length / reviewRowsPerPage) || sortedReviews.length === 0}
                                    className="px-3 h-full hover:bg-gray-50 border-l border-gray-300 disabled:opacity-30 disabled:cursor-not-allowed flex items-center justify-center text-slate-500"
                                >
                                    <svg xmlns="http://www.w3.org/2000/svg" className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 5l7 7-7 7" /></svg>
                                </button>
                            </div>
                            
                            <div className="flex items-center gap-2 font-medium text-slate-600">
                                <span>Rows per page:</span>
                                <div className="relative flex items-center">
                                    <select 
                                        value={reviewRowsPerPage} 
                                        onChange={(e) => {
                                            const val = e.target.value === 'All' ? 'All' : Number(e.target.value);
                                            setReviewRowsPerPage(val);
                                            setReviewPage(1);
                                        }}
                                        className="appearance-none bg-transparent border-none focus:ring-0 focus:outline-none pr-5 cursor-pointer text-slate-800 font-medium"
                                    >
                                        <option value={50}>50</option>
                                        <option value={100}>100</option>
                                        <option value={200}>200</option>
                                        <option value="All">All</option>
                                    </select>
                                    <svg xmlns="http://www.w3.org/2000/svg" className="h-4 w-4 absolute right-0 pointer-events-none text-slate-600" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M19 9l-7 7-7-7" /></svg>
                                </div>
                            </div>

                            <div className="font-medium text-slate-600">
                                Showing {sortedReviews.length === 0 ? 0 : (reviewRowsPerPage === 'All' ? 1 : (reviewPage - 1) * reviewRowsPerPage + 1)} to {reviewRowsPerPage === 'All' || sortedReviews.length === 0 ? sortedReviews.length : Math.min(reviewPage * reviewRowsPerPage, sortedReviews.length)} of {sortedReviews.length} results
                            </div>
                        </div>
                        
                        <div className="mt-8 pt-6 border-t border-gray-200">
                            <div className="flex justify-end space-x-4 items-center">
                                <button onClick={() => setShowBackConfirmModal(true)} disabled={isLoading.submitting} className="bg-white hover:bg-gray-100 text-gray-700 font-semibold py-2 px-4 rounded-md border border-gray-300 shadow-sm disabled:opacity-50">&larr; Back</button>
                                <button 
                                    onClick={() => {
                                        if (hasValidationErrors) {
                                            setHasAttemptedSubmit(true);
                                        } else {
                                            setShowConfirmModal(true);
                                        }
                                    }} 
                                    disabled={isLoading.submitting || adjustmentsToReview.length === 0} 
                                    className="bg-blue-600 hover:bg-blue-700 text-white font-bold py-2 px-4 rounded-md disabled:bg-gray-400">
                                    Send Adjustments
                                </button>
                            </div>
                        </div>
                    </div>}

                    {currentStep === 'processing' && <div className="bg-white p-8 rounded-lg shadow-md max-w-[672px] mx-auto text-center">
                        <h2 className="text-2xl font-bold mb-6 text-gray-800">Processing Updates</h2>
                        <div className="py-8">
                            <ProgressBar progress={progress} text={`Processing updates (${Math.round(progress)}%)...`} />
                            
                            <div className="flex justify-center gap-8 mt-6">
                                <div className="text-green-600 font-bold text-lg">Successful: {adjustmentsToReview.filter(a => a.status === 'success').length}</div>
                                <div className="text-red-600 font-bold text-lg">Failed: {adjustmentsToReview.filter(a => a.status === 'error').length}</div>
                            </div>

                            <div className="mt-6 text-amber-600 text-sm font-semibold flex items-center justify-center animate-pulse">
                                <ExclamationIcon className="w-5 h-5 mr-1" />
                                ⚠️ Please keep this tab active and do not let your computer sleep.
                            </div>
                            <div className="mt-8 flex justify-center">
                                <button onClick={() => setShowAbortModal(true)} className="bg-red-600 hover:bg-red-700 text-white font-bold py-2 px-6 rounded-md shadow transition-colors">
                                    Abort Process
                                </button>
                            </div>
                        </div>
                    </div>}

                    {currentStep === 'summary' && <div className="bg-white p-8 rounded-lg shadow-md max-w-[1024px] mx-auto text-center">
                        <h2 className="text-2xl font-bold mb-6 text-gray-800">Update Summary</h2>
                        <div className="flex justify-around text-center my-8">
                            <div><p className="text-5xl font-bold text-green-500">{updateSummary.filter(s => s.status === 'success').length}</p><p className="text-gray-500 mt-1">Successful Updates</p></div>
                            <div><p className="text-5xl font-bold text-amber-500">{updateSummary.filter(s => s.status === 'skipped').length}</p><p className="text-gray-500 mt-1">Skipped Updates</p></div>
                            <div><p className="text-5xl font-bold text-red-500">{updateSummary.filter(s => s.status === 'error').length}</p><p className="text-gray-500 mt-1">Failed Updates</p></div>
                        </div>
                        
                        <div className="flex justify-between items-center mt-8 mb-4">
                            <h3 className="text-lg font-semibold text-gray-800">Detailed Results</h3>
                        </div>
                        
                        <div className="max-h-96 overflow-y-auto border rounded-lg shadow-inner">
                            <table className="w-full text-sm text-left">
                                <thead className="text-xs text-gray-700 uppercase bg-gray-50 sticky top-0">
                                    <tr>
                                        {/* Removed Time */}
                                        <th className="px-4 py-3">Employee</th>
                                        <th className="px-4 py-3">Account</th>
                                        <th className="px-4 py-3">Validity Period</th>
                                        <th className="px-4 py-3">Unit Type</th>
                                        <th className="px-4 py-3 text-right">Adj.</th>
                                        <th className="px-4 py-3 text-right group cursor-help relative">
                                            New Balance
                                            <div className="absolute top-full left-1/2 transform -translate-x-1/2 mt-2 w-48 p-2 bg-gray-900 text-white text-xs rounded transition-opacity normal-case shadow-xl z-50 text-left invisible opacity-0 group-hover:visible group-hover:opacity-100 pointer-events-none text-center font-normal">
                                                Calculated on the "Effective date" (it could differ from current/today's balance).
                                            </div>
                                        </th>
                                        <th className="px-4 py-3">Result Message</th>
                                        <th className="px-4 py-3">Comment</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {sortedSummary.map(s => (
                                        <tr key={s.id} className={`border-b hover:bg-gray-50 ${s.status === 'error' ? 'bg-red-50' : s.status === 'skipped' ? 'bg-amber-50' : 'bg-white'}`}>
                                            {/* Removed Time data cell */}
                                            <td className="px-4 py-3 font-medium text-gray-900 leading-tight">
                                                <div>{s.employeeName}</div>
                                                {s.salaryIdentifier && <div className="text-xs font-mono text-gray-500 mt-0.5">SID: {s.salaryIdentifier}</div>}
                                            </td>
                                            <td className="px-4 py-3">{s.accountName}</td>
                                            {/* Added Validity Period data cell */}
                                            <td className="px-4 py-3 text-xs text-gray-500 whitespace-nowrap">
                                                {s.validFrom ? formatDateForDisplay(s.validFrom, downloadDateFormat) : 'N/A'} - {s.validTo ? formatDateForDisplay(s.validTo, downloadDateFormat) : '∞'}
                                            </td>
                                            <td className="px-4 py-3 text-xs text-gray-700 whitespace-nowrap">{s.unit || 'N/A'}</td>
                                            <td className={`px-4 py-3 text-right font-mono ${s.adjustment >= 0 ? 'text-green-600' : 'text-red-600'}`}>
                                                {s.adjustment > 0 ? '+' : ''}{s.adjustment.toFixed(2)}
                                            </td>
                                            <td className="px-4 py-3 text-right font-mono font-medium">
                                                {s.postAdjustmentBalance !== undefined && !isNaN(s.postAdjustmentBalance) ? s.postAdjustmentBalance.toFixed(2) : '-'}
                                            </td>
                                            <td className={`px-4 py-3 text-xs ${s.status === 'error' ? 'text-red-600 font-semibold' : s.status === 'skipped' ? 'text-amber-600 font-semibold' : 'text-green-600 font-semibold'}`}>
                                                {s.status === 'error' || s.status === 'skipped' ? s.error : "Updated Successfully"}
                                            </td>
                                            <td className="px-4 py-3 text-xs text-gray-500 truncate max-w-[320px]" title={s.comment}>
                                                {s.comment}
                                            </td>
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>
                        
                         <div className="mt-8 pt-6 border-t border-gray-200 flex justify-end gap-3">
                            <button onClick={handleExportResults} className="flex items-center bg-green-600 hover:bg-green-700 text-white font-bold py-2 px-4 rounded-md">
                                <DownloadIcon className="h-4 w-4 mr-2"/> Export Results
                            </button>
                            <button onClick={handleStartOver} className="bg-blue-600 hover:bg-blue-700 text-white font-bold py-2 px-4 rounded-md">Start New Adjustment</button>
                         </div>
                    </div>}

                </main>
            </div>
            
            <footer className="py-6 text-center">
                 <div className="relative group inline-flex items-center gap-1 cursor-help">
                    <span className="text-gray-500 text-sm font-medium">App Info</span>
                    <InfoIcon className="h-4 w-4 text-gray-400 group-hover:text-blue-500" />
                    <div className="absolute bottom-full left-1/2 transform -translate-x-1/2 mb-3 w-80 p-4 bg-gray-900 text-white text-xs rounded-lg shadow-xl opacity-0 invisible group-hover:opacity-100 group-hover:visible transition-all duration-200 z-50 text-center leading-relaxed">
                         <p className="mb-2">
                            <strong>Client-Side Processing & Secure:</strong> This app is a React-based JavaScript web application that runs entirely in your browser.
                        </p>
                        <p className="mb-3">
                            Your Excel files and employee data are processed locally in your browser and are never sent to any third‑party servers; they are only transmitted directly to the official <a href="https://openapi.planday.com/" target="_blank" rel="noopener noreferrer" className="text-blue-300 hover:text-blue-200 underline">Planday Open API</a> over a secure encrypted connection (HTTPS).
                        </p>
                        <p className="text-gray-400 border-t border-gray-700 pt-2 mt-2">Version 2.1</p>
                        <p className="text-gray-400">Made with ❤️ by the Planday Community</p>
                        <div className="absolute top-full left-1/2 transform -translate-x-1/2 border-4 border-transparent border-t-gray-900"></div>
                    </div>
                </div>
            </footer>
            
            {/* Fixed Floating Tooltip Container */}
            {activeTooltip && (
                <div 
                    className="fixed z-[100] p-3 bg-gray-900 text-white text-xs font-normal normal-case rounded-md shadow-xl pointer-events-none transition-opacity duration-200"
                    style={{
                        top: activeTooltip.y,
                        left: activeTooltip.x,
                        transform: 'translate(-50%, 0)',
                    }}
                >
                    {activeTooltip.content}
                    <div className="absolute bottom-full left-1/2 transform -translate-x-1/2 border-4 border-transparent border-b-gray-900"></div>
                </div>
            )}

            {/* Back Confirmation Modal */}
            {showBackConfirmModal && (
                <div className="fixed inset-0 z-50 flex items-center justify-center bg-black bg-opacity-50">
                    <div className="bg-white rounded-lg p-6 max-w-[448px] w-full shadow-xl">
                        <h3 className="text-lg font-bold text-gray-900 mb-4">Go Back?</h3>
                        <p className="text-gray-600 mb-6">
                            Are you sure you want to go back? Any edits you have made to the table will not be saved.
                        </p>
                        <div className="flex justify-end gap-3 flex-col sm:flex-row">
                            <button 
                                onClick={() => setShowBackConfirmModal(false)}
                                className="px-4 py-2 text-gray-700 bg-gray-100 hover:bg-gray-200 rounded-md font-medium"
                            >
                                Cancel
                            </button>
                            <button 
                                onClick={() => {
                                    setShowBackConfirmModal(false);
                                    setCurrentStep(updateMethod === 'excel' ? 'upload' : 'configure');
                                }}
                                className="px-4 py-2 bg-red-600 hover:bg-red-700 text-white rounded-md font-bold"
                            >
                                Yes, Go Back
                            </button>
                        </div>
                    </div>
                </div>
            )}

            {/* Update Balance Modal */}
            {showUpdateBalanceModal && (
                <div className="fixed inset-0 z-50 flex items-center justify-center bg-black bg-opacity-50">
                    <div className="bg-white rounded-lg p-6 max-w-[600px] w-full shadow-xl">
                        <h3 className="text-lg font-bold text-orange-600 mb-4">Update available balance?</h3>
                        <p className="text-gray-600 mb-4 whitespace-pre-wrap">
                            Since the table was generated based on your file upload, the current displayed balance may not reflect recent updates (e.g., approved leave, accruals, manual adjustments, etc.) processed after the template file was generated. Click "Yes, update now" to ensure that they match.
                            {'\n\n'}
                            Note, that the app will have to fetch the account balances again which can take some time. Table edits and inputs will not change, however the adjustments will be recalculated based on the new potential difference between Available balance and New balance target.
                        </p>
                        <div className="flex justify-end gap-3 flex-col sm:flex-row">
                            <button 
                                onClick={() => setShowUpdateBalanceModal(false)}
                                className="px-4 py-2 text-gray-700 bg-gray-100 hover:bg-gray-200 rounded-md font-medium"
                            >
                                No, keep file dates
                            </button>
                            <button 
                                onClick={handleUpdateBalances}
                                className="px-4 py-2 text-white bg-orange-500 hover:bg-orange-600 rounded-md font-bold"
                            >
                                Yes, update now
                            </button>
                        </div>
                    </div>
                </div>
            )}

            {/* Updating Balances Loader UI */}
            {isUpdatingBalances && (
                <div className="fixed inset-0 z-50 flex items-center justify-center bg-black bg-opacity-50">
                    <div className="bg-white rounded-lg p-8 max-w-[600px] w-full shadow-xl text-center relative">
                        <h2 className="text-2xl font-bold mb-6 text-gray-800">Updating Balances</h2>
                        <ProgressBar progress={progress} text={`${loadingText} (${Math.round(progress)}%)`} />
                        <div className="mt-8">
                            <button 
                                onClick={handleAbort}
                                className="bg-red-100 hover:bg-red-200 text-red-700 font-semibold py-2 px-6 rounded-md shadow-sm transition-colors duration-200"
                            >
                                Cancel
                            </button>
                        </div>
                    </div>
                </div>
            )}

            {/* Confirmation Modal */}
            {showConfirmModal && (
                <div className="fixed inset-0 z-50 flex items-center justify-center bg-black bg-opacity-50">
                    <div className="bg-white rounded-lg p-6 max-w-[448px] w-full shadow-xl">
                        <h3 className="text-lg font-bold text-gray-900 mb-4">Confirm Update</h3>
                        <p className="text-gray-600 mb-6">
                            The update process is about to start for <span className="text-blue-600 font-semibold">{adjustmentsToReview.length} accounts</span>. Have you reviewed the balance adjustments and are you ready to proceed?
                        </p>
                        <div className="flex justify-end gap-3 flex-col sm:flex-row">
                            <button 
                                onClick={() => setShowConfirmModal(false)}
                                className="px-4 py-2 text-gray-700 bg-gray-100 hover:bg-gray-200 rounded-md font-medium"
                            >
                                Cancel
                            </button>
                            <button 
                                onClick={executeBatchUpdate}
                                className="px-4 py-2 text-white bg-blue-600 hover:bg-blue-700 rounded-md font-bold"
                            >
                                Yes, start now
                            </button>
                        </div>
                    </div>
                </div>
            )}

            {/* Remove Confirmation Modal */}
            {showRemoveConfirmModal && (
                <div className="fixed inset-0 z-50 flex items-center justify-center bg-black bg-opacity-50">
                    <div className="bg-white rounded-lg p-6 max-w-[448px] w-full shadow-xl">
                        <h3 className="text-lg font-bold text-gray-900 mb-4 text-center">Remove Selected</h3>
                        <p className="text-gray-600 mb-6 text-center">
                            Are you sure you want to remove the <span className="font-bold text-gray-900">{selectedReviewIds.size}</span> selected employee account(s) from the table? This will omit them from being updated.
                        </p>
                        <div className="flex justify-center gap-3">
                            <button 
                                onClick={() => setShowRemoveConfirmModal(false)}
                                className="px-4 py-2 border border-gray-300 text-gray-700 bg-white hover:bg-gray-50 rounded-md font-medium"
                            >
                                Cancel
                            </button>
                            <button 
                                onClick={confirmRemoveSelected}
                                className="px-4 py-2 bg-red-600 hover:bg-red-700 text-white rounded-md font-bold"
                            >
                                Remove Selection
                            </button>
                        </div>
                    </div>
                </div>
            )}

            {/* Stop Process Modal */}
            {showStopProcessModal && (
                <div className="fixed inset-0 z-50 flex items-center justify-center bg-black bg-opacity-50">
                    <div className="bg-white rounded-lg p-6 max-w-[448px] w-full shadow-xl">
                        <h3 className="text-lg font-bold text-red-600 mb-4 flex items-center">
                            <ExclamationIcon className="w-6 h-6 mr-2" />
                            Stop Process?
                        </h3>
                        <p className="text-gray-600 mb-6">
                            Are you sure you want to stop the current process? This will cancel the fetching of accounts.
                        </p>
                        <div className="flex justify-end gap-3 flex-col sm:flex-row">
                            <button 
                                onClick={() => setShowStopProcessModal(false)}
                                className="px-4 py-2 text-gray-700 bg-gray-100 hover:bg-gray-200 rounded-md font-medium"
                            >
                                No, keep going
                            </button>
                            <button 
                                onClick={() => {
                                    abortRef.current = true;
                                    setShowStopProcessModal(false);
                                }}
                                className="px-4 py-2 text-white bg-red-600 hover:bg-red-700 rounded-md font-bold"
                            >
                                Yes, stop process
                            </button>
                        </div>
                    </div>
                </div>
            )}

            {/* Abort Modal */}
            {showAbortModal && (
                <div className="fixed inset-0 z-50 flex items-center justify-center bg-black bg-opacity-50">
                    <div className="bg-white rounded-lg p-6 max-w-[448px] w-full shadow-xl">
                        <h3 className="text-lg font-bold text-red-600 mb-4 flex items-center">
                            <ExclamationIcon className="w-6 h-6 mr-2" />
                            Abort Process?
                        </h3>
                        <p className="text-gray-600 mb-6">
                            Are you sure you want to abort the update process? This will cancel all remaining balance adjustments. The currently processing batch will finish, and remaining items will be marked as aborted in the results summary.
                        </p>
                        <div className="flex justify-end gap-3 flex-col sm:flex-row">
                            <button 
                                onClick={() => setShowAbortModal(false)}
                                className="px-4 py-2 text-gray-700 bg-gray-100 hover:bg-gray-200 rounded-md font-medium"
                            >
                                Continue Update
                            </button>
                            <button 
                                onClick={() => {
                                    setShowAbortModal(false);
                                    handleAbort();
                                }}
                                className="px-4 py-2 text-white bg-red-600 hover:bg-red-700 rounded-md font-bold"
                            >
                                Yes, Abort
                            </button>
                        </div>
                    </div>
                </div>
            )}

            {/* Bulk Edit Notification Modal */}
            {bulkEditDateWarning && (
                <div className="fixed inset-0 z-50 flex items-center justify-center bg-black bg-opacity-50 transition-opacity">
                    <div className="bg-white rounded-lg p-6 max-w-[500px] w-full shadow-2xl relative">
                        <h3 className="text-lg font-bold text-gray-900 mb-3 flex items-center">
                            <ExclamationIcon className="w-6 h-6 mr-2 text-amber-500" />
                            Notice: Accounts without a Start Date
                        </h3>
                        <p className="text-gray-600 mb-6 leading-relaxed">
                            The Account Start Date will not be applied to <span className="font-bold text-gray-900">{bulkEditDateWarning.skippedflexCount}</span> FLEX/TOIL account(s) you selected because they do not have a defined start date.
                            <br/><br/>
                            It will be successfully applied to the remaining selected accounts that have a valid start date.
                        </p>
                        <div className="flex justify-end gap-3 flex-col sm:flex-row">
                            <button 
                                onClick={() => setBulkEditDateWarning(null)}
                                className="px-4 py-2 text-gray-700 bg-gray-100 hover:bg-gray-200 rounded-md font-medium transition-colors"
                            >
                                Apply different date
                            </button>
                            <button 
                                onClick={() => {
                                    setBulkEditDateWarning(null);
                                    executeApplyBulkEdit();
                                }}
                                className="px-5 py-2 text-white bg-green-600 hover:bg-green-700 rounded-md font-bold shadow-sm transition-colors"
                            >
                                OK
                            </button>
                        </div>
                    </div>
                </div>
            )}
        </div>
    );
};

const AuthStep: React.FC<{ onAuthSuccess: (credentials: PlandayApiCredentials) => void; }> = ({ onAuthSuccess }) => {
    return (
        <div className="grid md:grid-cols-2 gap-8 items-start max-w-[1152px] mx-auto">
            <CredentialsForm onSave={onAuthSuccess} />
            <HelpPanel />
        </div>
    );
}

const CredentialsForm: React.FC<{ onSave: (credentials: PlandayApiCredentials) => void;}> = ({ onSave }) => {
  const [refreshToken, setRefreshToken] = useState('');
  const [error, setError] = useState('');
  const APP_ID = "4cb66728-94bf-416b-8d6c-892e4d36b38e";

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!refreshToken) { setError('Refresh token is required.'); return; }
    setError('');
    onSave({ clientId: APP_ID, refreshToken });
  };

  return (
    <div className="bg-white p-8 rounded-lg shadow-md border border-gray-200">
        <h2 className="text-2xl font-bold mb-2 text-gray-800">Connect to Planday</h2>
        <p className="text-gray-500 mb-6">Enter your Planday refresh token to connect with the App.</p>
        <form onSubmit={handleSubmit} className="space-y-4">
            <div>
                <label className="block text-sm font-medium text-gray-700 mb-1" htmlFor="refreshToken">Refresh Token</label>
                <input id="refreshToken" type="password" value={refreshToken} onChange={(e) => setRefreshToken(e.target.value)} className="w-full px-3 py-2 bg-white border border-gray-300 rounded-md shadow-sm text-gray-900 focus:outline-none focus:ring-2 focus:ring-blue-500" placeholder="Enter Refresh Token"/>
            </div>
            {error && <p className="text-red-500 text-sm">{error}</p>}
            <button type="submit" className="w-full bg-blue-600 hover:bg-blue-700 text-white font-bold py-2.5 px-4 rounded-md focus:outline-none focus:shadow-outline transition duration-300">Connect to Planday</button>
        </form>
    </div>
  );
};

const HelpPanel: React.FC = () => {
    const APP_ID = "4cb66728-94bf-416b-8d6c-892e4d36b38e";
    const [copied, setCopied] = useState(false);

    const handleCopy = () => {
        navigator.clipboard.writeText(APP_ID);
        setCopied(true);
        setTimeout(() => setCopied(false), 2000);
    };

    return (
        <div className="bg-white p-8 rounded-lg shadow-md border border-gray-200">
             <h2 className="text-2xl font-bold mb-2 text-gray-800">How to get your refresh token</h2>
             <p className="text-gray-500 mb-6">Follow these steps to generate the necessary credentials from your Planday portal.</p>
             <ol className="list-decimal list-inside space-y-3 text-gray-600">
                <li>Log in to your Planday portal</li>
                <li>Go to Settings &rarr; API Access</li>
                <li>
                    Click "Connect APP" and connect to app:
                    <div className="flex items-center gap-2 mt-1 p-2 bg-gray-100 rounded-md">
                        <code className="text-sm text-gray-800 flex-grow">{APP_ID}</code>
                        <button onClick={handleCopy} className="p-1.5 rounded-md hover:bg-gray-200 text-gray-500 hover:text-gray-800">
                            {copied ? <CheckIcon className="h-5 w-5 text-green-600"/> : <CopyIcon className="h-5 w-5"/>}
                        </button>
                    </div>
                </li>
                <li>Authorize the app when prompted</li>
                <li>Copy the "Token" value (this is your Refresh Token)</li>
             </ol>
        </div>
    );
};

export default App;
