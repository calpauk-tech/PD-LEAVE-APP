import { PlandayApiCredentials, Employee, LeaveAccount, LeaveAccountBalance, BalanceAdjustmentPayload, AccountType } from '../types';

const AUTH_URL = 'https://id.planday.com/connect/token';
const API_BASE_URL = 'https://openapi.planday.com';

let credentials_internal: PlandayApiCredentials | null = null;
let accessToken: string | null = null;
let tokenExpiry: number | null = null;

export function initializeService(credentials: PlandayApiCredentials) {
    if (credentials_internal?.clientId !== credentials.clientId || credentials_internal?.refreshToken !== credentials.refreshToken) {
        accessToken = null;
        tokenExpiry = null;
    }
    credentials_internal = { ...credentials };
}

async function wait(ms: number) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

// Adds a small random delay (0-500ms) to prevent 'thundering herd' when retrying batches
function getJitter() {
    return Math.floor(Math.random() * 500);
}

async function getAccessToken(): Promise<string> {
    if (accessToken && tokenExpiry && Date.now() < tokenExpiry) {
        return accessToken;
    }

    if (!credentials_internal) {
        throw new Error("Planday service not initialized with credentials.");
    }

    const payload = new URLSearchParams({
        'client_id': credentials_internal.clientId,
        'grant_type': 'refresh_token',
        'refresh_token': credentials_internal.refreshToken,
    });

    const response = await fetch(AUTH_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: payload.toString(),
    });

    if (!response.ok) {
        const errorText = await response.text();
        sessionStorage.removeItem('plandayCredentials');
        credentials_internal = null;
        accessToken = null;
        tokenExpiry = null;
        throw new Error(`Failed to refresh access token: ${response.status} ${errorText}. Your credentials may be invalid or expired. Please re-enter them.`);
    }

    const data = await response.json();
    accessToken = data.access_token;
    tokenExpiry = Date.now() + (data.expires_in - 60) * 1000;

    if (data.refresh_token && data.refresh_token !== credentials_internal.refreshToken) {
        credentials_internal.refreshToken = data.refresh_token;
        sessionStorage.setItem('plandayCredentials', JSON.stringify(credentials_internal));
    }
    
    return accessToken;
}

async function fetchWithAuth(url: string, options: RequestInit = { method: 'GET' }, retries = 5): Promise<Response> {
    if (!credentials_internal) throw new Error("Service not initialized");

    // We get the token inside the retry loop in case it expires during retries
    try {
        const token = await getAccessToken();
        const headers = {
            ...options.headers,
            'Authorization': `Bearer ${token}`,
            'X-ClientId': credentials_internal.clientId,
        };

        const response = await fetch(url, { ...options, headers });

        // Handle Rate Limiting (429)
        if (response.status === 429) {
            if (retries > 0) {
                const retryAfterHeader = response.headers.get('Retry-After');
                const xRateLimitReset = response.headers.get('x-ratelimit-reset');
                
                let waitTime = 2000; // Default 2s
                
                if (retryAfterHeader) {
                    waitTime = parseInt(retryAfterHeader, 10) * 1000;
                } else if (xRateLimitReset) {
                     // Add buffer to be safe
                     waitTime = (parseInt(xRateLimitReset, 10) + 1) * 1000;
                }
                
                // Add Jitter to prevent all throttled requests from hitting at the exact same millisecond
                waitTime += getJitter();

                console.warn(`Rate limited (429). Retrying in ${waitTime}ms...`);
                await wait(waitTime);
                return fetchWithAuth(url, options, retries - 1);
            }
        }

        // Handle Server Errors (5xx)
        if (response.status >= 500 && retries > 0) {
             console.warn(`Server error ${response.status}. Retrying...`);
             await wait(2000 + getJitter());
             return fetchWithAuth(url, options, retries - 1);
        }

        return response;
    } catch (error: any) {
        // Handle Network Errors (Fetch failed to leave browser)
        // Checks for "TypeError: Failed to fetch" (Chrome) or "NetworkError" (Firefox)
        const isNetworkError = error instanceof TypeError || error.name === 'TypeError' || error.name === 'NetworkError' || error.message?.includes('NetworkError');
        
        if (isNetworkError && retries > 0) {
            console.warn(`Network error detected: ${error.message}. Retries remaining: ${retries - 1}`);
            
            // Progressive backoff: 2s, 4s, 6s...
            // e.g. If retries=5 (first attempt), we wait 2000 * 1 = 2000ms
            const attempt = 6 - retries; 
            const backoff = 2000 * attempt; 
            
            await wait(backoff + getJitter());
            return fetchWithAuth(url, options, retries - 1);
        }
        throw error;
    }
}

async function fetchPaginatedData(endpoint: string): Promise<any[]> {
    let allData: any[] = [];
    let offset = 0;
    const limit = 50; // Keep limit reasonable to avoid massive response bodies

    while (true) {
        const url = `${API_BASE_URL}${endpoint}${endpoint.includes('?') ? '&' : '?'}limit=${limit}&offset=${offset}`;
        const response = await fetchWithAuth(url);
        if (!response.ok) throw new Error(`Failed to fetch ${endpoint}: ${await response.text()}`);
        const result = await response.json();

        if (result && result.data && Array.isArray(result.data)) {
            allData = allData.concat(result.data);
            if (result.data.length < limit || (result.paging && result.paging.total <= allData.length)) {
                break;
            }
            offset += result.data.length;
        } else {
            break;
        }
    }
    return allData;
}

export async function fetchEmployees(): Promise<Employee[]> {
    return fetchPaginatedData('/hr/v1.0/employees');
}

export async function fetchAccountTypes(): Promise<AccountType[]> {
    const response = await fetchWithAuth(`${API_BASE_URL}/absence/v1.0/accounttypes`);
    if (!response.ok) throw new Error(`Failed to fetch account types: ${await response.text()}`);
    const result = await response.json();
    return result.data;
}

export async function fetchLeaveAccounts(employeeId: number, dateRange?: { start: string; end: string }, status?: string): Promise<LeaveAccount[]> {
    let endpoint = `/absence/v1.0/accounts?employeeId=${employeeId}`;
    if (dateRange?.start) {
        endpoint += `&startDate=${dateRange.start}`;
    }
    if (dateRange?.end) {
        endpoint += `&endDate=${dateRange.end}`;
    }
    if (status) {
        endpoint += `&status=${status}`;
    }
    return fetchPaginatedData(endpoint);
}

export async function fetchAccountBalance(accountId: number, date: string): Promise<LeaveAccountBalance> {
    const result = await fetchWithAuth(`${API_BASE_URL}/absence/v1.0/accounts/${accountId}/balance?balanceDate=${date}`);
    
    if (result.status === 404) {
        return { balance: 0, unit: 'N/A' };
    }

    if (!result.ok) {
        const errorText = await result.text();
        throw new Error(`Failed to fetch account balance: ${errorText}`);
    }

    const data = await result.json();
    const accountData = data.data;

    if (!accountData || !accountData.balance || !Array.isArray(accountData.balance) || accountData.balance.length === 0) {
       return { balance: 0, unit: 'N/A' };
    }

    const balanceEntry = accountData.balance[0];
    
    return {
        balance: balanceEntry.value,
        unit: balanceEntry.unit?.type || 'N/A'
    };
}

export async function postBalanceAdjustment(accountId: number, payload: BalanceAdjustmentPayload): Promise<any> {
    const result = await fetchWithAuth(`${API_BASE_URL}/absence/v1.0/accounts/${accountId}/adjustments`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
    });
    if (!result.ok) {
        const errorText = await result.text();
        throw new Error(`Failed to post adjustment: ${errorText}`);
    }
    if (result.status === 201 || result.status === 200) {
        return { success: true };
    }
    return result.json();
}