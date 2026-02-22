// ============================================
// GAS API Service Layer
// Replaces google.script.run with fetch-based calls
// ============================================

const API_URL = "https://script.google.com/macros/s/AKfycbwDuDK2BYwykf0Z-u2FNFwxqyu0NZbE4emYceSMAIa3oD5JRUB9zIzRHfbVxtHdEzfnlg/exec";

export interface ApiResponse<T = unknown> {
  status: 'success' | 'error';
  data?: T;
  message?: string;
}

export async function callGasApi<T = unknown>(action: string, ...args: unknown[]): Promise<T> {
  console.log(`[GAS API] ${action}`, args);

  const response = await fetch(API_URL, {
    method: 'POST',
    body: JSON.stringify({ action, args }),
  });

  const result: ApiResponse<T> = await response.json();
  console.log(`[GAS API] Response: ${action}`, result);

  if (result.status === 'success') {
    return result.data as T;
  } else {
    throw new Error(result.message || 'API 呼叫失敗');
  }
}

// ============================================
// Typed API functions
// ============================================

export interface UserSession {
  isLoggedIn: boolean;
  name: string;
  role: string;
  email?: string;
}

export interface LoginResult {
  success: boolean;
  message: string;
  user?: UserSession;
}

export interface ProjectData {
  id: string;
  name: string;
  code: string;
  location?: string;
  status?: string;
  inspectors?: string[];
}

export interface InspectorData {
  id: string;
  name: string;
  department: string;
  code: string;
  status: string;
}

export interface DashboardData {
  year: number;
  month: number;
  workDays: { name: string; days: number }[];
  hazards: { type: string; count: number }[];
}

export interface LogEntry {
  date: string;
  projectName: string;
  inspector: string;
  weather: string;
  workers: number;
  workContent: string;
  hazards: string[];
  notes: string;
}

export interface SummaryData {
  date: string;
  projectName: string;
  inspector: string;
  status: string;
  workContent?: string;
}

export interface HolidayInfo {
  [date: string]: string;
}

// Auth
export const authenticateUser = (identifier: string, password: string) =>
  callGasApi<LoginResult>('authenticateUser', identifier, password);

export const getCurrentSession = () =>
  callGasApi<UserSession>('getCurrentSession');

export const logoutUser = () =>
  callGasApi<{ success: boolean }>('logoutUser');

// Data loading
export const loadInitialData = () =>
  callGasApi<{
    projects: ProjectData[];
    inspectors: InspectorData[];
    disasterTypes: string[];
  }>('loadInitialData');

// Dashboard
export const getMonthlyDashboardData = (year: number, month: number) =>
  callGasApi<DashboardData>('getMonthlyDashboardData', year, month);

// Summary
export const getSummaryData = (year: number, month: number, day?: number) =>
  callGasApi<SummaryData[]>('getSummaryData', year, month, day);

export const getGuestSummaryData = (mode: string) =>
  callGasApi<SummaryData[]>('getGuestSummaryData', mode);

// Log Entry
export const saveLogEntry = (entry: LogEntry) =>
  callGasApi<{ success: boolean; message: string }>('saveLogEntry', entry);

export const getLogEntry = (date: string, projectId: string) =>
  callGasApi<LogEntry>('getLogEntry', date, projectId);

// Holidays
export const getHolidays = (year: number, month: number) =>
  callGasApi<HolidayInfo>('getHolidays', year, month);

// Admin - Projects
export const getProjects = () =>
  callGasApi<ProjectData[]>('getProjects');

export const saveProject = (project: ProjectData) =>
  callGasApi<{ success: boolean }>('saveProject', project);

// Admin - Inspectors
export const getInspectors = () =>
  callGasApi<InspectorData[]>('getInspectors');

export const saveInspector = (inspector: InspectorData) =>
  callGasApi<{ success: boolean }>('saveInspector', inspector);

// Admin - Users
export interface UserData {
  id: string;
  name: string;
  email: string;
  role: string;
  status: string;
}

export const getUsers = () =>
  callGasApi<UserData[]>('getUsers');

export const saveUser = (user: UserData) =>
  callGasApi<{ success: boolean }>('saveUser', user);

export const changePassword = (oldPassword: string, newPassword: string) =>
  callGasApi<{ success: boolean; message: string }>('changePassword', oldPassword, newPassword);

// Filled dates for calendar
export const getFilledDates = (year: number, month: number) =>
  callGasApi<string[]>('getFilledDates', year, month);

// Log status
export const getLogStatus = (year: number, month: number) =>
  callGasApi<unknown[]>('getLogStatus', year, month);
