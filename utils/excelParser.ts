
import * as XLSX from 'xlsx';
import { Employee, AuditResult, LocationMatch } from '../types';

function cleanMergedEmployeeData(rawName: string, existingBadge: string, existingJob: string, existingLoc: string) {
  let name = rawName.trim();
  let badge = existingBadge;
  let job = existingJob;
  let loc = existingLoc;

  const badgeMatch = name.match(/\d{4,8}/); 
  if (badgeMatch && (!badge || badge === "N/A" || badge === "")) {
    badge = badgeMatch[0];
    name = name.replace(badge, "").trim();
  }

  const jobKeywords = [
    "CLEANER", "SUPERVISOR", "MANAGER", "DRIVER", "OPERATOR", 
    "TECHNICIAN", "FOREMAN", "HOUSEKEEPER", "WORKER", "HELPER", 
    "ATTENDANT", "LABOR", "COORDINATOR", "LEADMAN", "TEAM LEADER"
  ];
  
  if (!job || job === "UNKNOWN" || job === "" || job === "N/A") {
    for (const keyword of jobKeywords) {
      if (name.toUpperCase().includes(keyword)) {
        job = keyword;
        const regex = new RegExp(keyword, "gi");
        name = name.replace(regex, "").trim();
        break;
      }
    }
  }

  name = name.replace(/\s{2,}/g, ' ')
             .replace(/^[\s\-\/]+|[\s\-\/]+$/g, '')
             .trim();

  return { name, badge, job, loc };
}

/**
 * دالة ذكية لتحديد الجنس بناءً على الاسم في حال غياب البيانات في الإكسل
 */
function guessGenderByName(name: string): string {
  const upperName = name.toUpperCase();
  // قائمة بأسماء أو نهايات أسماء غالباً ما تكون إناث
  const femalePatterns = [
    "MAISA", "NOURA", "SARA", "FATIMA", "MARIAM", "AISHIA", "LINDA", "JANE", 
    "HESSA", "REEM", "AMAL", "MONA", "LATIFA", "JAWAHER", "NADA", "ARWA"
  ];
  
  // الأسماء التي تنتهي بـ A أو AH أو IA غالباً إناث في السياق العربي/الإنجليزي
  if (femalePatterns.some(p => upperName.startsWith(p)) || 
      upperName.split(' ')[0].endsWith("A") || 
      upperName.split(' ')[0].endsWith("AH") ||
      upperName.split(' ')[0].endsWith("IA")) {
    return "FEMALE";
  }
  
  return "MALE"; // الافتراضي ذكر في حال عدم التأكد
}

export async function auditExcelFile(file: File, month: number): Promise<AuditResult> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        
        let actualOnSite = 0;
        let controlSheetName = "";
        const expectedJobCounts: Record<string, number> = {};
        const processedSheets: string[] = [];

        // البحث عن ورقة التحكم (التي تحتوي على الأعداد الإجمالية)
        for (const name of workbook.SheetNames) {
          const sheet = workbook.Sheets[name];
          const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" }) as any[][];
          
          let tableHeaderRow = -1;
          for (let r = 0; r < Math.min(jsonData.length, 100); r++) {
            const rowStr = (jsonData[r] || []).map(c => (c ?? "").toString().toUpperCase());
            if (rowStr.some(c => c && (
              c.includes("ACTUAL ON SITE") || 
              c.includes("JOB TITLE") || 
              c.includes("DESIGNATION") || 
              c.includes("CATEGORY") || 
              c.includes("NUMBER OF STAFF")
            ))) {
              tableHeaderRow = r;
              controlSheetName = name;
              break;
            }
          }

          if (tableHeaderRow !== -1) {
            const headers = jsonData[tableHeaderRow].map(h => (h ?? "").toString().toUpperCase().trim());
            const jobTitleCol = headers.findIndex(h => h && (h.includes("JOB TITLE") || h.includes("DESIGNATION") || h.includes("CATEGORY")));
            const actualCol = headers.findIndex(h => h && (h.includes("ACTUAL ON SITE") || h.includes("NUMBER OF STAFF") || h.includes("COUNT")));

            for (let r = tableHeaderRow + 1; r < jsonData.length; r++) {
              const row = jsonData[r];
              if (!row || row.length === 0) continue;
              
              const jobTitle = (row[jobTitleCol] ?? "").toString().trim();
              if (!jobTitle || jobTitle.toUpperCase().includes("TOTAL")) {
                if (jobTitle.toUpperCase().includes("TOTAL") && actualCol !== -1) {
                   const val = parseInt(row[actualCol]);
                   if (!isNaN(val)) actualOnSite = val;
                }
                break;
              }

              if (actualCol !== -1) {
                const count = parseInt(row[actualCol]) || 0;
                expectedJobCounts[jobTitle.toUpperCase()] = count;
              }
            }
            break; 
          }
        }

        const rawEmployees: Employee[] = [];

        for (const name of workbook.SheetNames) {
          if (name === controlSheetName) continue;
          
          const sheet = workbook.Sheets[name];
          const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" }) as any[][];
          
          let headerIdx = -1;
          for (let r = 0; r < Math.min(jsonData.length, 40); r++) {
            const rowStr = (jsonData[r] || []).map(c => (c ?? "").toString().toUpperCase());
            if (rowStr.some(c => c && (c.includes("NAME") || c.includes("EMP#") || c.includes("BADGE") || c.includes("الاسم")))) {
              headerIdx = r;
              break;
            }
          }

          if (headerIdx === -1) continue;

          processedSheets.push(name);
          const headers = jsonData[headerIdx].map(h => (h ?? "").toString().toUpperCase().trim());
          const getIdx = (terms: string[]) => headers.findIndex(h => h && terms.some(t => h.includes(t)));
          
          const idx = {
            nameEng: getIdx(["NAME (ENG)", "EMPLOYEE NAME", "STAFF NAME", "NAME"]),
            nameAr: getIdx(["NAME (AR)", "الاسم", "الأسم"]),
            job: getIdx(["JOB TITLE", "POSITION", "JOB", "ROLE", "DESIGNATION", "المسمى"]),
            nat: getIdx(["NATIONALITY", "الجنسية"]),
            gender: getIdx(["GENDER", "SEX", "M/F", "النوع", "الجنس"]),
            id: getIdx(["ID#", "IQAMA", "PASSPORT", "هوية"]),
            badge: getIdx(["EMP#", "BADGE", "STAFF NO", "FILE", "CODE", "رقم الموظف", "الرقم الوظيفي", "S.NO"]),
            mrn: getIdx(["MRN", "MEDICAL", "ملف طبي"]),
            loc: getIdx(["LOCATION", "ZONE", "AREA", "SITE", "الموقع"]),
            comp: getIdx(["COMPANY", "SPONSOR", "الشركة"])
          };

          for (let r = headerIdx + 1; r < jsonData.length; r++) {
            const row = jsonData[r];
            if (!row || idx.nameEng === -1 || !row[idx.nameEng]) continue;

            const rawNameEng = row[idx.nameEng].toString().trim();
            if (!rawNameEng || rawNameEng.toUpperCase().includes("TOTAL") || rawNameEng.toUpperCase() === "NAME (ENG)") continue;

            const val = (index: number) => (index !== -1 && row[index] !== undefined && row[index] !== null) ? row[index].toString().trim() : "";

            // --- منطق تحديد الجنس المحدث ---
            let finalGender = "MALE";
            
            // 1. الأولوية للعمود G (فهرس 6)
            const colGValue = row[6] ? row[6].toString().toUpperCase().trim() : "";
            
            // 2. فحص عمود الجنس المستخرج من العنوان (Header)
            const headerGenderValue = val(idx.gender).toUpperCase();

            if (colGValue === "F" || colGValue === "FEMALE" || headerGenderValue === "F" || headerGenderValue === "FEMALE") {
              finalGender = "FEMALE";
            } else if (colGValue === "M" || colGValue === "MALE" || headerGenderValue === "M" || headerGenderValue === "MALE") {
              finalGender = "MALE";
            } else {
              // 3. إذا كان فارغاً، استخدم التخمين الذكي بناءً على الاسم (لحل مشكلة Maisa وأخواتها)
              finalGender = guessGenderByName(rawNameEng);
            }

            let initialBadge = val(idx.badge);
            let initialJob = val(idx.job).toUpperCase();
            let initialLoc = val(idx.loc) || name;
            let mrnValue = val(idx.mrn);
            let idNum = val(idx.id);
            let nationalityVal = val(idx.nat);

            const cleaned = cleanMergedEmployeeData(rawNameEng, initialBadge, initialJob, initialLoc);

            rawEmployees.push({
              id: `${file.name}-${name}-${r}`,
              nameEng: cleaned.name,
              nameAr: val(idx.nameAr),
              nationality: nationalityVal || "Unknown",
              gender: finalGender,
              idNumber: idNum,
              badgeNo: cleaned.badge || mrnValue || idNum || "N/A",
              empId: cleaned.badge || idNum || "",
              company: val(idx.comp),
              jobTitle: cleaned.job || "UNKNOWN",
              mrn: mrnValue,
              location: cleaned.loc,
              shift: "",
              month,
              sheetName: name,
              fileName: file.name
            });
          }
        }

        const employeeMap: Record<string, Employee> = {};
        const seenMap: Record<string, string[]> = {};
        const rosterJobCounts: Record<string, number> = {};

        rawEmployees.forEach(emp => {
          const identifier = (emp.mrn || emp.badgeNo || emp.nameEng.toUpperCase()).replace(/^0+/, '').replace(/\s+/g, '');
          if (!seenMap[identifier]) seenMap[identifier] = [];
          seenMap[identifier].push(emp.sheetName);
          if (!employeeMap[identifier]) {
            employeeMap[identifier] = emp;
          }
        });

        const uniqueEmployees = Object.values(employeeMap);
        uniqueEmployees.forEach(e => {
           const j = e.jobTitle.toUpperCase();
           rosterJobCounts[j] = (rosterJobCounts[j] || 0) + 1;
        });

        const locationMatches: LocationMatch[] = [];
        Object.entries(expectedJobCounts).forEach(([job, expected]) => {
          const found = rosterJobCounts[job] || 0;
          locationMatches.push({
            location: job,
            expected,
            found,
            variance: found - expected
          });
        });

        const totalExpectedFromSummary = Object.values(expectedJobCounts).reduce((a, b) => a + b, 0);
        const finalTarget = actualOnSite || totalExpectedFromSummary;

        resolve({
          fileName: file.name,
          area: "Workforce Audit",
          category: "Operations",
          actualOnSite: finalTarget,
          masterTarget: 0,
          calculatedCount: uniqueEmployees.length,
          difference: uniqueEmployees.length - finalTarget,
          status: uniqueEmployees.length === finalTarget ? 'PASS' : 'FAIL',
          processedSheets,
          locationAnalysis: Object.entries(rosterJobCounts).map(([location, count]) => ({ location, count })),
          locationMatches,
          duplicates: Object.entries(seenMap)
            .filter(([_, sheets]) => sheets.length > 1)
            .map(([id, sheets]) => {
              const emp = employeeMap[id];
              return {
                identifier: id,
                name: emp?.nameEng || "Unknown",
                sheets: Array.from(new Set(sheets)),
                jobTitle: emp?.jobTitle || "Unknown",
                mrn: emp?.mrn || ""
              };
            }) as any[],
          employees: uniqueEmployees
        });
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(new Error("File reading failed"));
    reader.readAsArrayBuffer(file);
  });
}
