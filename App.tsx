
import React, { useState, useMemo } from 'react';
import { 
  AlertTriangle, 
  CheckCircle, 
  Upload, 
  Search, 
  MapPin, 
  ShieldCheck,
  Layers,
  Loader2,
  TrendingUp,
  ArrowRight,
  Users,
  ArrowLeftRight,
  UserPlus,
  UserMinus,
  RefreshCw,
  FileSpreadsheet,
  FileDown,
  TrendingDown,
  Info,
  BarChart3,
  CalendarDays,
  ChevronRight,
  FileText,
  Activity,
  ArrowRightLeft,
  Printer,
  FileBarChart,
  History,
  UserCheck,
  UserX,
  Repeat,
  Globe,
  Contact,
  VenetianMask,
  PieChart as PieIcon,
  Table as TableIcon
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import 'jspdf-autotable';
import { AreaChart, Area, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, PieChart, Pie, Cell } from 'recharts';
import { auditExcelFile } from './utils/excelParser';
import { AuditResult, Employee, MovementReport } from './types';

const MONTH_NAMES = [
  { en: "January", short: "Jan", ar: "يناير", num: "01", year: 2025 },
  { en: "February", short: "Feb", ar: "فبراير", num: "02", year: 2025 },
  { en: "March", short: "Mar", ar: "مارس", num: "03", year: 2025 },
  { en: "April", short: "Apr", ar: "أبريل", num: "04", year: 2025 },
  { en: "May", short: "May", ar: "مايو", num: "05", year: 2025 },
  { en: "June", short: "Jun", ar: "يونيو", num: "06", year: 2025 },
  { en: "July", short: "Jul", ar: "يوليو", num: "07", year: 2025 },
  { en: "August", short: "Aug", ar: "أغسطس", num: "08", year: 2025 },
  { en: "September", short: "Sep", ar: "سبتمبر", num: "09", year: 2025 },
  { en: "October", short: "Oct", ar: "أكتوبر", num: "10", year: 2025 },
  { en: "November", short: "Nov", ar: "نوفمبر", num: "11", year: 2025 },
  { en: "December", short: "Dec", ar: "ديسمبر", num: "12", year: 2025 },
  { en: "January", short: "Jan '26", ar: "يناير 2026", num: "01", year: 2026 }
];

const COLORS = ['#6366f1', '#10b981', '#f59e0b', '#f43f5e', '#8b5cf6', '#06b6d4', '#ec4899'];

const App: React.FC = () => {
  const [auditData, setAuditData] = useState<Record<number, AuditResult[]>>({});
  const [selectedMonth, setSelectedMonth] = useState<number>(new Date().getMonth() + 1);
  const [isProcessing, setIsProcessing] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [activeTab, setActiveTab] = useState<'roster' | 'movement' | 'yearly' | 'lifecycle' | 'analytics'>('roster');

  const currentResult = useMemo(() => auditData[selectedMonth]?.[0] || null, [auditData, selectedMonth]);
  
  const previousResult = useMemo(() => {
    for (let i = selectedMonth - 1; i >= 1; i--) {
      if (auditData[i]?.[0]) return auditData[i][0];
    }
    return null;
  }, [auditData, selectedMonth]);

  const demographics = useMemo(() => {
    if (!currentResult) return null;
    const nationalityMap = new Map<string, number>();
    const genderMap = new Map<string, number>();
    const jobMap = new Map<string, number>();

    currentResult.employees.forEach(e => {
      const nat = e.nationality || 'Other';
      nationalityMap.set(nat, (nationalityMap.get(nat) || 0) + 1);
      
      const gender = e.gender?.toUpperCase() === 'FEMALE' ? 'Female' : 'Male';
      genderMap.set(gender, (genderMap.get(gender) || 0) + 1);
      
      const job = e.jobTitle || 'Unassigned';
      jobMap.set(job, (jobMap.get(job) || 0) + 1);
    });

    const natData = Array.from(nationalityMap.entries()).map(([name, value]) => ({ name, value })).sort((a,b) => b.value - a.value);
    const genData = Array.from(genderMap.entries()).map(([name, value]) => ({ name, value }));
    const jobData = Array.from(jobMap.entries()).map(([name, value]) => ({ name, value })).sort((a,b) => b.value - a.value);

    return { natData, genData, jobData };
  }, [currentResult]);

  const yearlyStats = useMemo(() => {
    const data: any[] = [];
    let lastFoundCount = 0;
    for (let i = 1; i <= MONTH_NAMES.length; i++) {
      const result = auditData[i]?.[0];
      if (result) {
        const currentCount = result.calculatedCount;
        const variance = lastFoundCount === 0 ? 0 : currentCount - lastFoundCount;
        let predecessorResult = null;
        for (let j = i - 1; j >= 1; j--) {
          if (auditData[j]?.[0]) { predecessorResult = auditData[j][0]; break; }
        }
        let joiners = 0, leavers = 0;
        if (predecessorResult) {
          const currIds = new Set(result.employees.map(e => (e.mrn || e.badgeNo || e.nameEng.toUpperCase()).trim()));
          const prevIds = new Set(predecessorResult.employees.map(e => (e.mrn || e.badgeNo || e.nameEng.toUpperCase()).trim()));
          result.employees.forEach(e => { if (!prevIds.has((e.mrn || e.badgeNo || e.nameEng.toUpperCase()).trim())) joiners++; });
          predecessorResult.employees.forEach(e => { if (!currIds.has((e.mrn || e.badgeNo || e.nameEng.toUpperCase()).trim())) leavers++; });
        }
        data.push({
          month: i,
          monthName: MONTH_NAMES[i - 1].en,
          monthNameAr: MONTH_NAMES[i - 1].ar,
          year: MONTH_NAMES[i - 1].year,
          count: currentCount,
          variance,
          joiners,
          leavers
        });
        lastFoundCount = currentCount;
      }
    }
    return data;
  }, [auditData]);

  const staffLifecycle = useMemo(() => {
    const historyMap = new Map<string, {
      name: string,
      badge: string,
      presence: (string | null)[], 
      status: 'Stable' | 'Returned' | 'Left-No-Return',
      lastSeen: number,
      firstSeen: number,
      gaps: number
    }>();

    const latestMonthUploaded = Math.max(...Object.keys(auditData).map(Number), 0);

    for (let m = 1; m <= MONTH_NAMES.length; m++) {
      const result = auditData[m]?.[0];
      if (!result) continue;

      result.employees.forEach(emp => {
        const id = (emp.mrn || emp.badgeNo || emp.nameEng.toUpperCase()).trim();
        if (!historyMap.has(id)) {
          historyMap.set(id, {
            name: emp.nameEng,
            badge: emp.badgeNo || emp.mrn || "N/A",
            presence: new Array(MONTH_NAMES.length).fill(null),
            status: 'Stable',
            lastSeen: m,
            firstSeen: m,
            gaps: 0
          });
        }
        const record = historyMap.get(id)!;
        record.presence[m - 1] = emp.location;
        record.lastSeen = Math.max(record.lastSeen, m);
      });
    }

    const list = Array.from(historyMap.values());
    list.forEach(item => {
      let gaps = 0;
      let returned = false;
      let latestMonth = Math.max(...Object.keys(auditData).map(Number), 0);
      
      for (let i = item.firstSeen - 1; i < latestMonth; i++) {
        if (item.presence[i] === null) {
          gaps++;
        } else if (gaps > 0) {
          returned = true;
        }
      }

      if (returned) {
        item.status = 'Returned';
      } else if (item.lastSeen < latestMonth) {
        item.status = 'Left-No-Return';
      } else {
        item.status = 'Stable';
      }
      item.gaps = gaps;
    });

    return list;
  }, [auditData]);

  const movementReport = useMemo((): MovementReport | null => {
    if (!currentResult || !previousResult) return null;
    const currMap = new Map<string, Employee>();
    currentResult.employees.forEach(e => currMap.set((e.mrn || e.badgeNo || e.nameEng.toUpperCase()).trim(), e));
    const prevMap = new Map<string, Employee>();
    previousResult.employees.forEach(e => prevMap.set((e.mrn || e.badgeNo || e.nameEng.toUpperCase()).trim(), e));
    const newJoiners: Employee[] = [];
    const leavers: MovementReport['leavers'] = [];
    const locationSwaps: MovementReport['locationSwaps'] = [];
    currMap.forEach((emp, id) => {
      if (!prevMap.has(id)) newJoiners.push(emp);
      else {
        const prevEmp = prevMap.get(id)!;
        if (prevEmp.location.trim().toUpperCase() !== emp.location.trim().toUpperCase()) {
          locationSwaps.push({ employee: emp, oldLocation: prevEmp.location, newLocation: emp.location });
        }
      }
    });
    prevMap.forEach((emp, id) => {
      if (!currMap.has(id)) leavers.push({ nameEng: emp.nameEng, badgeNo: emp.badgeNo, lastLocation: emp.location, lastJob: emp.jobTitle });
    });
    return { newJoiners, leavers, locationSwaps };
  }, [currentResult, previousResult]);

  const renderTextToImage = (text: string, width: number, height: number, color: string, isRtl: boolean = false): string => {
    const canvas = document.createElement('canvas');
    const scale = 4;
    canvas.width = width * scale;
    canvas.height = height * scale;
    const ctx = canvas.getContext('2d');
    if (!ctx) return '';
    ctx.scale(scale, scale);
    ctx.font = 'bold 24px "Noto Sans Arabic", "Inter", sans-serif';
    ctx.fillStyle = color;
    ctx.textBaseline = 'middle';
    if (isRtl) { ctx.textAlign = 'right'; ctx.fillText(text, width, height / 2); }
    else { ctx.textAlign = 'left'; ctx.fillText(text, 0, height / 2); }
    return canvas.toDataURL('image/png');
  };

  const exportMonthlyExcel = () => {
    // تم التعديل للسماح بالتصدير حتى لو لم يتوفر movementReport
    if (!currentResult || !demographics) return;
    const wb = XLSX.utils.book_new();
    const monthObj = MONTH_NAMES[selectedMonth - 1];
    const monthNameEn = monthObj.en;

    // Sheet 1: Nationalities and Gender
    const summaryData = [
      [`Summary Statistics for ${monthNameEn} ${monthObj.year}`, ""],
      [],
      ["Nationality", "Count", "Percentage"],
      ...demographics.natData.map(d => [d.name, d.value, `${((d.value/currentResult.calculatedCount)*100).toFixed(1)}%`]),
      [],
      ["Gender", "Count", "Percentage"],
      ...demographics.genData.map(d => [d.name, d.value, `${((d.value/currentResult.calculatedCount)*100).toFixed(1)}%`])
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summaryData), "Stats Summary");

    // Sheet 2: Job Titles
    const jobSheetData = [["Job Title", "Total Staff"], ...demographics.jobData.map(d => [d.name, d.value])];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(jobSheetData), "Job Roles");

    // الشيتات التفاعلية (فقط في حال وجود حركة)
    const joinersList = movementReport?.newJoiners || [];
    const leaversList = movementReport?.leavers || [];
    const swapsList = movementReport?.locationSwaps || [];

    // Sheet 3: Joiners
    const joinersSheetData = [["Staff Name", "Badge/MRN", "Location", "Job Title"], ...joinersList.map(e => [e.nameEng, e.badgeNo || e.mrn, e.location, e.jobTitle])];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(joinersSheetData), "Joiners Report");

    // Sheet 4: Leavers
    const leaversSheetData = [["Staff Name", "Badge/MRN", "Last Known Location", "Job Title"], ...leaversList.map(e => [e.nameEng, e.badgeNo, e.lastLocation, e.lastJob])];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(leaversSheetData), "Leavers Report");

    // Sheet 5: Transfers
    const transfersSheetData = [["Staff Name", "Badge/MRN", "Previous Location", "New Location"], ...swapsList.map(s => [s.employee.nameEng, s.employee.badgeNo || s.employee.mrn, s.oldLocation, s.newLocation])];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(transfersSheetData), "Site Transfers");

    XLSX.writeFile(wb, `Monthly_Workforce_Audit_${monthNameEn}_${monthObj.year}.xlsx`);
  };

  const exportLifecycleExcel = () => {
    if (staffLifecycle.length === 0) return;
    const wb = XLSX.utils.book_new();
    const departed = staffLifecycle.filter(i => i.status === 'Left-No-Return');
    const returned = staffLifecycle.filter(i => i.status === 'Returned');

    const departedData = [
      ["Departed Staff (No Return History in 2025-2026)"],
      [],
      ["Name", "Badge/MRN", "Month of Departure", "Last Site Location"],
      ...departed.map(i => [i.name, i.badge, `${MONTH_NAMES[i.lastSeen-1].ar} ${MONTH_NAMES[i.lastSeen-1].year}`, i.presence[i.lastSeen-1]])
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(departedData), "Departed Staff");

    const returnedData = [
      ["Returned Staff (Gaps in Presence Identified)"],
      [],
      ["Name", "Badge/MRN", "Last Return Month", "Presence Journey"],
      ...returned.map(i => [i.name, i.badge, `${MONTH_NAMES[i.lastSeen-1].ar} ${MONTH_NAMES[i.lastSeen-1].year}`, i.presence.filter(p => p).join(' -> ')])
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(returnedData), "Returned Staff");

    XLSX.writeFile(wb, "Staff_Lifecycle_Tracking_Audit_2025_2026.xlsx");
  };

  const exportAnnualExcel = () => {
    if (yearlyStats.length === 0) return;
    const wb = XLSX.utils.book_new();
    const annualData = [
      ["Annual Workforce Performance Summary - 2025-2026"],
      [],
      ["Month", "Year", "Total Headcount", "Joiners", "Leavers", "Net Variance"],
      ...yearlyStats.map(s => [s.monthName, s.year, s.count, s.joiners, s.leavers, (s.variance > 0 ? '+' : '') + s.variance])
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(annualData), "Annual Summary");
    XLSX.writeFile(wb, "Annual_Workforce_Report_2025_2026.xlsx");
  };

  const exportLifecyclePDF = () => {
    if (staffLifecycle.length === 0) return;
    const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    
    doc.setFillColor(242, 242, 242); doc.rect(0, 0, pageWidth, 40, 'F');
    doc.setTextColor(0, 0, 0); doc.setFont("helvetica", "bold"); doc.setFontSize(16);
    doc.text("STAFF RETENTION & ATTRITION REPORT - 2025-2026", 15, 18);
    const arabicTitleImg = renderTextToImage("تقرير استمرارية ومغادرة الكوادر البشرية ٢٠٢٥-٢٠٢٦", 600, 40, "#000000", false);
    if (arabicTitleImg) doc.addImage(arabicTitleImg, 'PNG', 15, 23, 100, 7);

    const leaversNoReturn = staffLifecycle.filter(i => i.status === 'Left-No-Return');
    const returnees = staffLifecycle.filter(i => i.status === 'Returned');

    doc.setFontSize(11);
    doc.setFont("helvetica", "bold");
    doc.text("1. DEPARTED STAFF (NO RETURN)", 15, 50);
    
    const departedTotalArImg = renderTextToImage(`إجمالي المغادرين بدون عودة: ${leaversNoReturn.length} موظف`, 600, 40, "#333333", false);
    if (departedTotalArImg) doc.addImage(departedTotalArImg, 'PNG', 15, 53, 60, 5);

    (doc as any).autoTable({
      startY: 62, 
      head: [['Name', 'Badge', 'Last Month', 'Last Site']],
      body: leaversNoReturn.map(i => [i.name, i.badge, `${MONTH_NAMES[i.lastSeen-1].en} ${MONTH_NAMES[i.lastSeen-1].year}`, i.presence[i.lastSeen-1]]),
      theme: 'grid', 
      styles: { fontSize: 8 }, 
      headStyles: { fillColor: [0, 0, 0] }
    });

    doc.addPage();
    doc.setFontSize(11);
    doc.setFont("helvetica", "bold");
    doc.text("2. STAFF RE-JOINERS (RETURNED)", 15, 20);
    
    if (returnees.length === 0) {
      const zeroReturnImg = renderTextToImage("الإجمالي: صفر - لا يوجد موظفين عائدين", 600, 40, "#991b1b", false);
      if (zeroReturnImg) doc.addImage(zeroReturnImg, 'PNG', 15, 25, 65, 6);
    } else {
      const returnTotalArImg = renderTextToImage(`إجمالي الموظفين العائدين: ${returnees.length} موظف`, 600, 40, "#333333", false);
      if (returnTotalArImg) doc.addImage(returnTotalArImg, 'PNG', 15, 25, 60, 5);

      (doc as any).autoTable({
        startY: 32, 
        head: [['Name', 'Badge', 'Return Month', 'Sites Visited']],
        body: returnees.map(i => [i.name, i.badge, `${MONTH_NAMES[i.lastSeen-1].en} ${MONTH_NAMES[i.lastSeen-1].year}`, Array.from(new Set(i.presence.filter(p => p))).join(' -> ')]),
        theme: 'grid', 
        styles: { fontSize: 8 }, 
        headStyles: { fillColor: [0, 0, 0] }
      });
    }

    const pageCount = (doc as any).internal.getNumberOfPages();
    for (let i = 1; i <= pageCount; i++) {
        doc.setPage(i);
        doc.setFontSize(7);
        doc.setTextColor(100, 100, 100);
        doc.text(`Prepared by Inspector Layla Alotaibi | Official Document | Page ${i} of ${pageCount}`, pageWidth / 2, pageHeight - 10, { align: 'center' });
    }

    doc.save(`Staff_Lifecycle_Audit_2025_2026.pdf`);
  };

  const exportMonthlyPDF = () => {
    if (!currentResult || !demographics) return;
    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
    const monthObj = MONTH_NAMES[selectedMonth - 1];
    const monthNameEn = monthObj.en;
    const monthNameAr = monthObj.ar;
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();

    const addHeader = (titleEn: string, titleAr: string, bgColor: [number, number, number]) => {
      doc.setFillColor(...bgColor); doc.rect(0, 0, pageWidth, 35, 'F');
      doc.setTextColor(255, 255, 255); doc.setFontSize(22); doc.setFont("helvetica", "bold");
      doc.text(titleEn, 15, 18);
      const arabicImg = renderTextToImage(titleAr, 450, 60, '#ffffff', true);
      if (arabicImg) doc.addImage(arabicImg, 'PNG', pageWidth - 125, 10, 110, 15);
      doc.setFontSize(10); doc.text(`Workforce Audit | Month: ${monthNameEn} ${monthObj.year} | Headcount: ${currentResult.calculatedCount}`, 15, 28);
    };

    const addFooter = () => {
      const pageCount = (doc as any).internal.getNumberOfPages();
      for (let i = 1; i <= pageCount; i++) {
        doc.setPage(i); doc.setFontSize(9); doc.setTextColor(100, 116, 139); doc.setFont("helvetica", "bold");
        doc.text("This report was prepared by Inspector Layla Alotaibi", pageWidth / 2, pageHeight - 15, { align: 'center' });
        doc.setFontSize(8); doc.setTextColor(148, 163, 184); doc.text(`Page ${i} of ${pageCount} | Workforce Intelligent Auditor 2025-2026`, pageWidth / 2, pageHeight - 10, { align: 'center' });
      }
    };

    // PAGE 1: DEMOGRAPHICS
    addHeader(`Demographic Profile: ${monthNameEn} ${monthObj.year}`, `تقرير تصنيفات الجنسيات والنوع: ${monthNameAr}`, [79, 70, 229]);
    doc.setTextColor(0,0,0); doc.setFontSize(11); doc.text("Nationality Breakdown", 15, 45);
    (doc as any).autoTable({ startY: 48, head: [['Nationality', 'Count', 'Ratio']], body: demographics.natData.map(d => [d.name, d.value, `${((d.value/currentResult.calculatedCount)*100).toFixed(1)}%`]), theme: 'grid', styles: { fontSize: 8 }, headStyles: { fillColor: [79, 70, 229] } });
    
    doc.text("Gender Distribution", 15, (doc as any).lastAutoTable.finalY + 10);
    (doc as any).autoTable({ startY: (doc as any).lastAutoTable.finalY + 13, head: [['Gender', 'Count', 'Ratio']], body: demographics.genData.map(d => [d.name, d.value, `${((d.value/currentResult.calculatedCount)*100).toFixed(1)}%`]), theme: 'grid', styles: { fontSize: 8 }, headStyles: { fillColor: [236, 72, 153] } });

    // PAGE 2: JOB TITLES
    doc.addPage();
    addHeader(`Workforce Job Roles: ${monthNameEn} ${monthObj.year}`, `تقرير المسميات الوظيفية الشامل: ${monthNameAr}`, [15, 118, 110]);
    (doc as any).autoTable({ startY: 45, head: [['Job Title (Category)', 'Staff Count']], body: demographics.jobData.map(d => [d.name, d.value]), theme: 'grid', styles: { fontSize: 8 }, headStyles: { fillColor: [13, 148, 136] } });

    // PAGE 3: JOINERS
    const joiners = movementReport?.newJoiners || [];
    if (joiners.length > 0) {
      doc.addPage();
      addHeader(`Joiners Audit: ${monthNameEn} ${monthObj.year}`, `تقرير المنضمين الجدد خلال الشهر: ${monthNameAr}`, [16, 185, 129]);
      (doc as any).autoTable({ startY: 45, head: [['Staff Name', 'Badge', 'Site Location', 'Job Title']], body: joiners.map(e => [e.nameEng, e.badgeNo || e.mrn, e.location, e.jobTitle]), headStyles: { fillColor: [5, 150, 105], fontSize: 9 }, styles: { fontSize: 8 }, margin: { bottom: 25 } });
    }

    // PAGE 4: LEAVERS
    const leavers = movementReport?.leavers || [];
    if (leavers.length > 0) {
      doc.addPage();
      addHeader(`Leavers Audit: ${monthNameEn} ${monthObj.year}`, `تقرير الموظفين المغادرين: ${monthNameAr}`, [244, 63, 94]);
      (doc as any).autoTable({ startY: 45, head: [['Staff Name', 'Badge', 'Last Known Location', 'Job Title']], body: leavers.map(e => [e.nameEng, e.badgeNo, e.lastLocation, e.lastJob]), headStyles: { fillColor: [225, 29, 72], fontSize: 9 }, styles: { fontSize: 8 }, margin: { bottom: 25 } });
    }
    
    // PAGE 5: TRANSFERS
    const swaps = movementReport?.locationSwaps || [];
    if (swaps.length > 0) {
      doc.addPage();
      addHeader(`Transfer Audit: ${monthNameEn} ${monthObj.year}`, `تقرير تنقلات الموظفين بين المواقع: ${monthNameAr}`, [245, 158, 11]);
      (doc as any).autoTable({ startY: 45, head: [['Staff Name', 'Badge', 'From Location (Previous)', 'To Location (Current)']], body: swaps.map(s => [s.employee.nameEng, s.employee.badgeNo || s.employee.mrn, s.oldLocation, s.newLocation]), headStyles: { fillColor: [217, 119, 6], fontSize: 9 }, styles: { fontSize: 8 }, margin: { bottom: 25 } });
    }

    addFooter();
    doc.save(`Detailed_Workforce_Report_${monthNameEn}_${monthObj.year}.pdf`);
  };

  const exportAnnualPDF = () => {
    if (yearlyStats.length === 0) return;
    const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    
    doc.setFillColor(242, 242, 242); doc.rect(0, 0, pageWidth, 35, 'F');
    doc.setTextColor(0, 0, 0); doc.setFont("helvetica", "bold"); doc.setFontSize(16);
    doc.text("ANNUAL WORKFORCE REPORT - 2025-2026", 15, 18);
    const arabicTitleImg = renderTextToImage("التقرير السنوي الإجمالي للقوى العاملة لعام ٢٠٢٥-٢٠٢٦", 600, 40, "#000000", false);
    if (arabicTitleImg) doc.addImage(arabicTitleImg, 'PNG', 15, 23, 90, 7);
    
    (doc as any).autoTable({
      startY: 48,
      margin: { left: 15, right: 15 },
      head: [['Month', 'Year', 'Headcount', 'Joiners', 'Leavers', 'Variance']],
      body: yearlyStats.map(s => [s.monthName, s.year, s.count, s.joiners, s.leavers, s.variance]),
      theme: 'grid',
      headStyles: { fillColor: [242, 242, 242], textColor: [0, 0, 0], fontStyle: 'bold' },
      styles: { fontSize: 8 }
    });

    doc.save(`Annual_Workforce_Report_2025_2026.pdf`);
  };

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const files = event.target.files; if (!files) return;
    setIsProcessing(true); const newResults = { ...auditData }; let firstDetectedMonth = -1;
    for (let i = 0; i < files.length; i++) {
      const file = files[i]; const lowerName = file.name.toLowerCase(); let monthIdx = -1;
      
      MONTH_NAMES.forEach((m, idx) => {
        const yearStr = m.year.toString();
        const monthMatch = [m.en.toLowerCase(), m.short.toLowerCase().replace("'26", ""), m.ar].some(p => lowerName.includes(p.toLowerCase()));
        const yearMatch = lowerName.includes(yearStr);
        if (monthMatch && yearMatch) monthIdx = idx + 1;
      });

      if (monthIdx === -1) {
        const numMatch = lowerName.match(/(?:^|[\s\-_])(0[1-9]|1[0-2]|[1-9])(?:[\s\-_]|$)/);
        const yearMatch = lowerName.match(/(2025|2026)/);
        if (numMatch && yearMatch) {
            const mNum = parseInt(numMatch[1]);
            const yNum = parseInt(yearMatch[1]);
            if (yNum === 2026 && mNum === 1) monthIdx = 13;
            else if (yNum === 2025) monthIdx = mNum;
        }
      }

      if (monthIdx !== -1) { 
        try { 
          const result = await auditExcelFile(file, monthIdx); 
          newResults[monthIdx] = [result]; 
          if (firstDetectedMonth === -1) firstDetectedMonth = monthIdx; 
        } catch (e) { console.error(e); } 
      }
    }
    setAuditData(newResults); if (firstDetectedMonth !== -1) setSelectedMonth(firstDetectedMonth);
    setIsProcessing(false);
  };

  return (
    <div className="min-h-screen bg-[#f8fafc] text-slate-900 font-inter">
      <header className="bg-slate-900 text-white px-8 py-6 sticky top-0 z-50 shadow-2xl">
        <div className="max-w-[1600px] mx-auto flex flex-col md:flex-row items-center justify-between gap-4">
          <div className="flex items-center gap-4">
            <div className="bg-indigo-500 p-2.5 rounded-2xl shadow-lg shadow-indigo-500/20"><ShieldCheck size={28} /></div>
            <div>
              <h1 className="text-xl font-black tracking-tight uppercase">Manning Auditor <span className="text-indigo-400">2025-2026</span></h1>
              <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Inspector: Layla Alotaibi</p>
            </div>
          </div>
          <div className="flex gap-3">
             <label className="flex items-center gap-3 bg-indigo-600 hover:bg-indigo-500 px-6 py-3 rounded-2xl text-xs font-black cursor-pointer transition-all active:scale-95 shadow-xl shadow-indigo-600/20">
              {isProcessing ? <Loader2 className="animate-spin" size={16} /> : <Upload size={16} />}
              رفع ملفات السنة
              <input type="file" multiple className="hidden" onChange={handleFileUpload} />
            </label>
          </div>
        </div>
      </header>

      <main className="max-w-[1600px] mx-auto px-6 py-8">
        <div className="grid grid-cols-4 md:grid-cols-6 lg:grid-cols-13 gap-2 mb-10 overflow-x-auto pb-4">
          {MONTH_NAMES.map((m, idx) => {
            const num = idx + 1; const hasData = !!auditData[num];
            return (
              <button key={idx} onClick={() => setSelectedMonth(num)} className={`min-w-[80px] p-4 rounded-[1.25rem] border-2 flex flex-col items-center transition-all duration-300 ${selectedMonth === num ? 'bg-indigo-600 border-indigo-600 text-white shadow-xl scale-105' : hasData ? 'bg-emerald-50 border-emerald-200 text-emerald-700' : 'bg-white border-slate-200 text-slate-300 opacity-60'}`}>
                <span className="text-[10px] font-black uppercase tracking-tighter mb-1">{m.short}</span>
                <span className="text-sm font-black whitespace-nowrap">{m.ar}</span>
                {hasData && <div className="w-1.5 h-1.5 bg-current rounded-full mt-2 animate-pulse"></div>}
              </button>
            );
          })}
        </div>

        {Object.keys(auditData).length > 0 ? (
          <div className="grid lg:grid-cols-12 gap-8">
            <div className="lg:col-span-3 space-y-6">
              <div className="bg-white p-8 rounded-[2.5rem] border border-slate-200 shadow-sm">
                <h3 className="text-[11px] font-black uppercase text-slate-400 mb-6 flex items-center gap-2 tracking-widest"><Activity size={14} /> حالة القوى العاملة</h3>
                <div className="space-y-4">
                  <div className="bg-slate-50 p-5 rounded-3xl border border-slate-100">
                    <p className="text-[10px] font-bold text-slate-400 uppercase mb-1">العدد الكلي في {MONTH_NAMES[selectedMonth-1].ar} {MONTH_NAMES[selectedMonth-1].year}</p>
                    <p className="text-4xl font-black text-slate-900">{currentResult?.calculatedCount || 0}</p>
                  </div>
                  {previousResult && (
                    <div className={`p-5 rounded-3xl border ${currentResult!.calculatedCount >= previousResult.calculatedCount ? 'bg-emerald-50 border-emerald-100 text-emerald-700' : 'bg-rose-50 border-rose-100 text-rose-700'}`}>
                      <p className="text-[10px] font-bold uppercase mb-1">صافي التغيير</p>
                      <div className="flex items-center gap-2">
                        {currentResult!.calculatedCount >= previousResult.calculatedCount ? <TrendingUp size={24} /> : <TrendingDown size={24} />}
                        <p className="text-3xl font-black">{currentResult!.calculatedCount - previousResult.calculatedCount}</p>
                      </div>
                    </div>
                  )}
                </div>
              </div>

              <div className="bg-slate-900 p-8 rounded-[2.5rem] shadow-2xl text-white space-y-4">
                 <h3 className="text-[11px] font-black uppercase text-indigo-400 tracking-widest">مركز التقارير (Reports Hub)</h3>
                 
                 <div className="space-y-2 pb-2 border-b border-white/10">
                    <p className="text-[10px] font-bold text-slate-400 uppercase tracking-tighter">تقرير الشهر ({MONTH_NAMES[selectedMonth-1].en} {MONTH_NAMES[selectedMonth-1].year})</p>
                    <div className="grid grid-cols-2 gap-2">
                      <button onClick={exportMonthlyPDF} className="flex items-center justify-center gap-2 bg-white/10 hover:bg-white/20 p-3 rounded-xl transition-all text-[10px] font-black uppercase">
                        PDF <FileDown size={14} />
                      </button>
                      <button onClick={exportMonthlyExcel} className="flex items-center justify-center gap-2 bg-emerald-600/20 hover:bg-emerald-600/30 p-3 rounded-xl transition-all text-[10px] font-black uppercase text-emerald-400">
                        Excel <FileSpreadsheet size={14} />
                      </button>
                    </div>
                 </div>

                 <div className="space-y-2 pb-2 border-b border-white/10">
                    <p className="text-[10px] font-bold text-slate-400 uppercase tracking-tighter">سجل المغادرين والعودة</p>
                    <div className="grid grid-cols-2 gap-2">
                      <button onClick={exportLifecyclePDF} className="flex items-center justify-center gap-2 bg-white/10 hover:bg-white/20 p-3 rounded-xl transition-all text-[10px] font-black uppercase">
                        PDF <History size={14} />
                      </button>
                      <button onClick={exportLifecycleExcel} className="flex items-center justify-center gap-2 bg-emerald-600/20 hover:bg-emerald-600/30 p-3 rounded-xl transition-all text-[10px] font-black uppercase text-emerald-400">
                        Excel <TableIcon size={14} />
                      </button>
                    </div>
                 </div>

                 <div className="space-y-2">
                    <p className="text-[10px] font-bold text-slate-400 uppercase tracking-tighter">التقرير السنوي الإجمالي</p>
                    <div className="grid grid-cols-2 gap-2">
                      <button onClick={exportAnnualPDF} className="flex items-center justify-center gap-2 bg-indigo-600 hover:bg-indigo-500 p-3 rounded-xl transition-all text-[10px] font-black uppercase">
                        PDF <Printer size={14} />
                      </button>
                      <button onClick={exportAnnualExcel} className="flex items-center justify-center gap-2 bg-emerald-600 hover:bg-emerald-500 p-3 rounded-xl transition-all text-[10px] font-black uppercase text-white shadow-lg shadow-emerald-600/20">
                        Excel <FileBarChart size={14} />
                      </button>
                    </div>
                 </div>

                 <p className="text-[9px] text-slate-400 leading-relaxed italic text-center border-t border-white/10 pt-4 mt-4">
                    "Official Document prepared by Inspector Layla Alotaibi"
                 </p>
              </div>
            </div>

            <div className="lg:col-span-9">
              <div className="bg-white rounded-[3rem] border border-slate-200 shadow-sm overflow-hidden min-h-[700px]">
                <div className="flex bg-slate-50/80 p-3 gap-3 border-b border-slate-100 overflow-x-auto custom-scrollbar">
                  <button onClick={() => setActiveTab('roster')} className={`px-8 py-3 rounded-2xl text-xs font-black transition-all shrink-0 flex items-center gap-2 ${activeTab === 'roster' ? 'bg-white shadow-md text-indigo-600' : 'text-slate-400 hover:text-slate-600'}`}><Users size={16}/> القائمة الحالية</button>
                  <button onClick={() => setActiveTab('movement')} className={`px-8 py-3 rounded-2xl text-xs font-black transition-all shrink-0 flex items-center gap-2 ${activeTab === 'movement' ? 'bg-white shadow-md text-indigo-600' : 'text-slate-400 hover:text-slate-600'}`}><ArrowRightLeft size={16}/> تحليل الحركات</button>
                  <button onClick={() => setActiveTab('analytics')} className={`px-8 py-3 rounded-2xl text-xs font-black transition-all shrink-0 flex items-center gap-2 ${activeTab === 'analytics' ? 'bg-white shadow-md text-indigo-600' : 'text-slate-400 hover:text-slate-600'}`}><PieIcon size={16}/> تصنيفات الموظفين</button>
                  <button onClick={() => setActiveTab('lifecycle')} className={`px-8 py-3 rounded-2xl text-xs font-black transition-all shrink-0 flex items-center gap-2 ${activeTab === 'lifecycle' ? 'bg-white shadow-md text-indigo-600' : 'text-slate-400 hover:text-slate-600'}`}><History size={16}/> تتبع الموظفين</button>
                  <button onClick={() => setActiveTab('yearly')} className={`px-8 py-3 rounded-2xl text-xs font-black transition-all shrink-0 flex items-center gap-2 ${activeTab === 'yearly' ? 'bg-white shadow-md text-indigo-600' : 'text-slate-400 hover:text-slate-600'}`}><BarChart3 size={16}/> ملخص 2025-2026</button>
                </div>

                <div className="p-8">
                  {activeTab === 'roster' && (
                    <div className="space-y-6 animate-in fade-in duration-500">
                      <div className="relative group">
                        <Search className="absolute left-5 top-1/2 -translate-y-1/2 text-slate-300 group-focus-within:text-indigo-400 transition-colors" size={20} />
                        <input type="text" placeholder="ابحث عن موظف بالاسم أو الموقع أو الرقم الوظيفي..." className="w-full pl-14 pr-8 py-5 bg-slate-50 rounded-[1.5rem] border-none text-sm font-medium focus:ring-4 ring-indigo-500/10 transition-all" value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} />
                      </div>
                      <div className="overflow-x-auto rounded-[2rem] border border-slate-100">
                        <table className="w-full text-left text-xs">
                          <thead className="bg-slate-50 text-slate-400 uppercase font-black tracking-widest">
                            <tr>
                              <th className="px-6 py-5">الموظف</th>
                              <th className="px-6 py-5">الموقع الحالي</th>
                              <th className="px-6 py-5">المسمى الوظيفي</th>
                              <th className="px-6 py-5">الجنس</th>
                              <th className="px-6 py-5 text-right">المعرّف (Badge)</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-50">
                            {currentResult?.employees.filter(e => e.nameEng.toLowerCase().includes(searchTerm.toLowerCase()) || e.location.toLowerCase().includes(searchTerm.toLowerCase()) || e.badgeNo?.includes(searchTerm)).map((e, i) => (
                              <tr key={i} className="hover:bg-slate-50/50 transition-colors group">
                                <td className="px-6 py-5 font-bold text-slate-800 text-sm">{e.nameEng}</td>
                                <td className="px-6 py-5"><span className="px-3 py-1 bg-slate-100 rounded-full font-bold text-slate-500">{e.location}</span></td>
                                <td className="px-6 py-5 text-slate-500 font-medium">{e.jobTitle}</td>
                                <td className="px-6 py-5">
                                  <span className={`px-2 py-0.5 rounded-md text-[10px] font-black ${e.gender === 'FEMALE' ? 'bg-pink-100 text-pink-700' : 'bg-blue-100 text-blue-700'}`}>
                                    {e.gender}
                                  </span>
                                </td>
                                <td className="px-6 py-5 text-right font-mono font-black text-indigo-600">{e.badgeNo || e.mrn || "N/A"}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  )}

                  {activeTab === 'analytics' && demographics && (
                    <div className="space-y-10 animate-in fade-in duration-500">
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                        <div className="bg-slate-50 p-8 rounded-[2.5rem] border border-slate-100">
                          <h4 className="text-sm font-black text-slate-800 uppercase mb-6 flex items-center gap-2"><Globe size={18} className="text-indigo-500"/> الجنسيات</h4>
                          <div className="space-y-3 max-h-[300px] overflow-y-auto pr-2 custom-scrollbar">
                            {demographics.natData.map((d, i) => (
                              <div key={i} className="flex justify-between items-center bg-white p-3 rounded-xl border border-slate-100">
                                <span className="text-xs font-bold text-slate-600">{d.name}</span>
                                <span className="text-xs font-black text-indigo-600 bg-indigo-50 px-3 py-1 rounded-full">{d.value}</span>
                              </div>
                            ))}
                          </div>
                        </div>
                        <div className="bg-slate-50 p-8 rounded-[2.5rem] border border-slate-100">
                          <h4 className="text-sm font-black text-slate-800 uppercase mb-6 flex items-center gap-2"><VenetianMask size={18} className="text-rose-500"/> النوع (Gender)</h4>
                          <div className="h-[200px] w-full">
                            <ResponsiveContainer width="100%" height="100%">
                              <PieChart>
                                <Pie data={demographics.genData} innerRadius={60} outerRadius={80} paddingAngle={5} dataKey="value">
                                  {demographics.genData.map((entry, index) => (
                                    <Cell key={`cell-${index}`} fill={entry.name === 'Female' ? '#ec4899' : '#6366f1'} />
                                  ))}
                                </Pie>
                                <Tooltip />
                              </PieChart>
                            </ResponsiveContainer>
                          </div>
                          <div className="flex justify-center gap-8 mt-4">
                            {demographics.genData.map((d, i) => (
                              <div key={i} className="flex items-center gap-2">
                                <div className={`w-3 h-3 rounded-full ${d.name === 'Female' ? 'bg-pink-500' : 'bg-indigo-500'}`}></div>
                                <span className="text-xs font-bold text-slate-600">{d.name}: {d.value}</span>
                              </div>
                            ))}
                          </div>
                        </div>
                      </div>
                    </div>
                  )}

                  {activeTab === 'movement' && (
                    <div className="space-y-10 animate-in slide-in-from-bottom-4 duration-500">
                      {!movementReport ? (
                         <div className="flex flex-col items-center justify-center p-20 bg-amber-50 rounded-[3rem] border-2 border-dashed border-amber-200">
                            <Info className="text-amber-500 mb-4" size={48} />
                            <p className="text-amber-700 font-black text-center">يرجى رفع ملف الشهر السابق لتفعيل تحليل الحركات (المنضمون والمغادرون)</p>
                         </div>
                      ) : (
                        <div className="grid md:grid-cols-3 gap-8">
                          <div className="space-y-5">
                            <h4 className="text-sm font-black text-emerald-600 flex items-center justify-between bg-emerald-50 p-4 rounded-2xl">
                               <div className="flex items-center gap-2"><UserPlus size={20} /> المنضمون الجدد</div>
                               <span className="bg-emerald-600 text-white px-3 py-1 rounded-full text-[11px]">{movementReport.newJoiners.length}</span>
                            </h4>
                            <div className="space-y-3 max-h-[400px] overflow-y-auto pr-2 custom-scrollbar">
                              {movementReport.newJoiners.map((e, i) => (
                                <div key={i} className="p-5 bg-white border border-slate-100 rounded-[1.5rem] shadow-sm hover:border-emerald-200 transition-all group">
                                  <p className="text-sm font-black text-slate-800">{e.nameEng}</p>
                                  <p className="text-[10px] text-slate-400 font-bold uppercase">{e.location}</p>
                                </div>
                              ))}
                            </div>
                          </div>
                          <div className="space-y-5">
                            <h4 className="text-sm font-black text-rose-600 flex items-center justify-between bg-rose-50 p-4 rounded-2xl">
                               <div className="flex items-center gap-2"><UserMinus size={20} /> المغادرون</div>
                               <span className="bg-rose-600 text-white px-3 py-1 rounded-full text-[11px]">{movementReport.leavers.length}</span>
                            </h4>
                            <div className="space-y-3 max-h-[400px] overflow-y-auto pr-2 custom-scrollbar">
                              {movementReport.leavers.map((e, i) => (
                                <div key={i} className="p-5 bg-white border border-slate-100 rounded-[1.5rem] shadow-sm hover:border-rose-200 transition-all group">
                                  <p className="text-sm font-black text-slate-800">{e.nameEng}</p>
                                  <p className="text-[10px] text-slate-400 font-bold uppercase">{e.lastLocation}</p>
                                </div>
                              ))}
                            </div>
                          </div>
                          <div className="space-y-5">
                            <h4 className="text-sm font-black text-amber-600 flex items-center justify-between bg-amber-50 p-4 rounded-2xl">
                               <div className="flex items-center gap-2"><ArrowRightLeft size={20} /> تنقلات المواقع</div>
                               <span className="bg-amber-600 text-white px-3 py-1 rounded-full text-[11px]">{movementReport.locationSwaps.length}</span>
                            </h4>
                            <div className="space-y-3 max-h-[400px] overflow-y-auto pr-2 custom-scrollbar">
                              {movementReport.locationSwaps.map((s, i) => (
                                <div key={i} className="p-5 bg-white border border-slate-100 rounded-[1.5rem] shadow-sm hover:border-amber-200 transition-all group">
                                  <p className="text-sm font-black text-slate-800">{s.employee.nameEng}</p>
                                  <div className="flex items-center gap-2 mt-2 text-[10px] font-bold">
                                     <span className="text-slate-400 line-through">{s.oldLocation}</span>
                                     <ArrowRight size={10} className="text-amber-500" />
                                     <span className="text-amber-600">{s.newLocation}</span>
                                  </div>
                                </div>
                              ))}
                            </div>
                          </div>
                        </div>
                      )}
                    </div>
                  )}

                  {activeTab === 'lifecycle' && (
                    <div className="space-y-6 animate-in fade-in duration-500">
                      <div className="grid md:grid-cols-3 gap-4 mb-6">
                        <div className="p-4 bg-emerald-50 border border-emerald-100 rounded-2xl">
                          <p className="text-[10px] font-black uppercase text-emerald-600 mb-1">المستقرون</p>
                          <p className="text-2xl font-black">{staffLifecycle.filter(i => i.status === 'Stable').length}</p>
                        </div>
                        <div className="p-4 bg-amber-50 border border-amber-100 rounded-2xl">
                          <p className="text-[10px] font-black uppercase text-amber-600 mb-1">المعاد تعيينهم (العودة)</p>
                          <p className="text-2xl font-black">{staffLifecycle.filter(i => i.status === 'Returned').length}</p>
                        </div>
                        <div className="p-4 bg-rose-50 border border-rose-100 rounded-2xl">
                          <p className="text-[10px] font-black uppercase text-rose-600 mb-1">غادروا ولم يعودوا</p>
                          <p className="text-2xl font-black">{staffLifecycle.filter(i => i.status === 'Left-No-Return').length}</p>
                        </div>
                      </div>

                      <div className="space-y-4">
                        {staffLifecycle.filter(i => i.name.toLowerCase().includes(searchTerm.toLowerCase()) || i.badge.includes(searchTerm)).map((item, idx) => (
                          <div key={idx} className="bg-white border border-slate-100 p-6 rounded-[2rem] shadow-sm hover:border-indigo-200 transition-all flex flex-col md:flex-row md:items-center justify-between gap-6">
                            <div className="flex items-center gap-4">
                              <div className={`p-3 rounded-2xl ${item.status === 'Stable' ? 'bg-emerald-100 text-emerald-600' : item.status === 'Returned' ? 'bg-amber-100 text-amber-600' : 'bg-rose-100 text-rose-600'}`}>
                                {item.status === 'Stable' ? <UserCheck size={20}/> : item.status === 'Returned' ? <Repeat size={20}/> : <UserX size={20}/>}
                              </div>
                              <div>
                                <h4 className="font-black text-slate-800">{item.name}</h4>
                                <p className="text-xs text-slate-400 font-bold uppercase tracking-widest">{item.badge}</p>
                              </div>
                            </div>
                            <div className="flex-1 flex items-center justify-center gap-2">
                              {item.presence.map((p, i) => (
                                <div key={i} title={p || 'Absent'} className={`w-3 h-3 rounded-full ${p ? 'bg-indigo-500' : 'bg-slate-100 border border-slate-200'}`}></div>
                              ))}
                            </div>
                            <div className="text-right">
                              <p className={`text-[10px] font-black uppercase px-3 py-1 rounded-full inline-block mb-1 ${item.status === 'Stable' ? 'bg-emerald-100 text-emerald-700' : item.status === 'Returned' ? 'bg-amber-100 text-amber-700' : 'bg-rose-100 text-rose-700'}`}>
                                {item.status === 'Stable' ? 'نشط/مستقر' : item.status === 'Returned' ? 'عاد خلال السنة' : 'غادر ولم يعد'}
                              </p>
                              <p className="text-[9px] text-slate-400 font-medium">آخر ظهور: {MONTH_NAMES[item.lastSeen-1].ar} {MONTH_NAMES[item.lastSeen-1].year}</p>
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}

                  {activeTab === 'yearly' && (
                    <div className="space-y-12 animate-in fade-in duration-500">
                      <div className="bg-slate-50 rounded-[2.5rem] p-10 border border-slate-100 relative overflow-hidden">
                        <div className="absolute top-0 right-0 w-64 h-64 bg-indigo-500/5 rounded-full -mr-20 -mt-20 blur-3xl"></div>
                        <h3 className="text-sm font-black text-slate-900 uppercase mb-10 flex items-center justify-between relative">
                          <div className="flex items-center gap-3"><TrendingUp size={20} className="text-indigo-500"/> توجهات القوى العاملة 2025-2026</div>
                        </h3>
                        <div className="h-[350px] w-full relative">
                          <ResponsiveContainer width="100%" height="100%">
                            <AreaChart data={yearlyStats}>
                              <defs>
                                <linearGradient id="colorCount" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#6366f1" stopOpacity={0.2}/><stop offset="95%" stopColor="#6366f1" stopOpacity={0}/></linearGradient>
                              </defs>
                              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                              <XAxis dataKey="monthName" axisLine={false} tickLine={false} tick={{fontSize: 11, fontWeight: 'bold'}} />
                              <YAxis axisLine={false} tickLine={false} tick={{fontSize: 11, fontWeight: 'bold'}} />
                              <Tooltip contentStyle={{ borderRadius: '20px', border: 'none', boxShadow: '0 20px 25px -5px rgb(0 0 0 / 0.1)' }} />
                              <Area type="monotone" dataKey="count" stroke="#6366f1" strokeWidth={5} fillOpacity={1} fill="url(#colorCount)" />
                            </AreaChart>
                          </ResponsiveContainer>
                        </div>
                      </div>
                      <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-6">
                        {yearlyStats.map((stat, i) => (
                          <div key={i} className="p-6 bg-white border border-slate-100 rounded-[2rem] shadow-sm hover:shadow-md transition-all cursor-pointer group" onClick={() => setSelectedMonth(stat.month)}>
                            <div className="flex justify-between items-center mb-5">
                              <span className="text-sm font-black text-slate-800">{stat.monthNameAr} {stat.year}</span>
                              <span className={`text-[11px] font-black px-3 py-1 rounded-full ${stat.variance >= 0 ? 'bg-emerald-100 text-emerald-700' : 'bg-rose-100 text-rose-700'}`}>
                                {stat.variance > 0 ? '+' : ''}{stat.variance}
                              </span>
                            </div>
                            <div className="grid grid-cols-3 gap-2 text-center">
                               <div className="p-2 bg-slate-50 rounded-xl"><p className="text-[9px] font-bold text-slate-400 uppercase mb-1">Total</p><p className="text-sm font-black text-slate-800">{stat.count}</p></div>
                               <div className="p-2 bg-emerald-50 rounded-xl"><p className="text-[9px] font-bold text-emerald-600 uppercase mb-1">In</p><p className="text-sm font-black text-emerald-600">{stat.joiners}</p></div>
                               <div className="p-2 bg-rose-50 rounded-xl"><p className="text-[9px] font-bold text-rose-600 uppercase mb-1">Out</p><p className="text-sm font-black text-rose-600">{stat.leavers}</p></div>
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}
                </div>
              </div>
            </div>
          </div>
        ) : (
          <div className="flex flex-col items-center justify-center min-h-[500px] bg-white rounded-[4rem] border-4 border-dashed border-slate-100 animate-pulse">
            <div className="bg-slate-50 p-10 rounded-full mb-6 text-slate-200"><FileSpreadsheet size={64} /></div>
            <p className="text-slate-400 font-black text-lg">بانتظار رفع ملفات الإكسل (2025 - 2026)</p>
            <p className="text-slate-300 text-sm mt-2 font-medium text-center leading-loose">
              Workforce Planning System<br/>
              <span className="font-black">This report was prepared by Inspector Layla Alotaibi</span>
            </p>
          </div>
        )}
      </main>
    </div>
  );
};

export default App;
