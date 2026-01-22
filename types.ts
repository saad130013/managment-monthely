
export interface Employee {
  id: string;
  nameEng: string;
  nameAr: string;
  nationality: string;
  gender: string;
  idNumber: string; 
  badgeNo: string;
  empId: string;
  company: string;
  jobTitle: string;
  location: string;
  mrn: string;
  shift: string;
  month: number;
  sheetName: string;
  fileName: string;
}

export interface MovementReport {
  newJoiners: Employee[];
  leavers: { nameEng: string; badgeNo: string; lastLocation: string; lastJob: string }[];
  locationSwaps: { employee: Employee; oldLocation: string; newLocation: string }[];
}

export interface LocationMatch {
  location: string;
  expected: number;
  found: number;
  variance: number;
}

export interface AuditResult {
  fileName: string;
  area: string;
  category: string;
  actualOnSite: number; 
  masterTarget: number; 
  calculatedCount: number; 
  difference: number;
  status: 'PASS' | 'FAIL';
  processedSheets: string[];
  locationAnalysis: { location: string; count: number }[];
  locationMatches: LocationMatch[];
  duplicates: { identifier: string; name: string; sheets: string[]; jobTitle: string; mrn: string }[];
  employees: Employee[];
}
