/************************************
 * GLOBAL VARIABLES & CONSTANTS
 ************************************/
const SERVICE_PACKAGE_RULES = {
  'Dental Services': ['dental', 'orthodontic'],
  'Physiotherapy': ['physio'],
  'Other OP Services': ['physio', 'diet', 'occupational', 'speech']
  // Consultation: custom handling below
};
const DATE_KEYS = ['Date', 'On'];
const MONTHS = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"];

let reportData = null;
let eligData = null;
const usedEligibilities = new Set();

// DOM Elements
const reportInput = document.getElementById("reportFileInput");
const eligInput = document.getElementById("eligibilityFileInput");
const processBtn = document.getElementById("processBtn");
const exportInvalidBtn = document.getElementById("exportInvalidBtn");
const status = document.getElementById("uploadStatus");
const resultsContainer = document.getElementById("results");

/************************************
 * UPDATE PROCESS BUTTON STATE
 ************************************/
function updateProcessButtonState() {
  const hasEligibility = !!eligData;
  const hasReportData = !!reportData;
  processBtn.disabled = !(hasEligibility && hasReportData);
  exportInvalidBtn.disabled = !(hasEligibility && hasReportData);
}

/************************************
 * DATE HANDLING UTILITIES
 ************************************/
let lastReportWasCSV = false;
const DateHandler = {
  parse: function(input, options = {}) {
    const preferMDY = !!options.preferMDY;
    if (!input) return null;
    if (input instanceof Date) return isNaN(input) ? null : input;
    if (typeof input === 'number') return this._parseExcelDate(input);

    const cleanStr = input.toString().trim().replace(/[,.]/g, '');
    const parsed = this._parseStringDate(cleanStr, preferMDY) || new Date(cleanStr);
    if (isNaN(parsed)) return null;
    return parsed;
  },
  format: function(date) {
    if (!(date instanceof Date) || isNaN(date)) return '';
    const d = date.getUTCDate().toString().padStart(2, '0');
    const m = (date.getUTCMonth() + 1).toString().padStart(2, '0');
    const y = date.getUTCFullYear();
    return `${d}/${m}/${y}`;
  },
  isSameDay: function(date1, date2) {
    if (!date1 || !date2) return false;
    return date1.getUTCDate() === date2.getUTCDate() &&
           date1.getUTCMonth() === date2.getUTCMonth() &&
           date1.getUTCFullYear() === date2.getUTCFullYear();
  },
  _parseExcelDate: function(serial) {
    const utcDays = Math.floor(serial) - 25569;
    const ms = utcDays * 86400 * 1000;
    const date = new Date(ms);
    return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
  },
  _parseStringDate: function(dateStr, preferMDY = false) {
    if (dateStr.includes(' ')) dateStr = dateStr.split(' ')[0];
    const dmyMdyMatch = dateStr.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
    if (dmyMdyMatch) {
      const part1 = parseInt(dmyMdyMatch[1], 10);
      const part2 = parseInt(dmyMdyMatch[2], 10);
      const year = parseInt(dmyMdyMatch[3], 10);

      if (part1 > 12 && part2 <= 12) {
        return new Date(Date.UTC(year, part2 - 1, part1));
      } else if (part2 > 12 && part1 <= 12) {
        return new Date(Date.UTC(year, part1 - 1, part2));
      } else {
        if (preferMDY) {
          return new Date(Date.UTC(year, part1 - 1, part2));
        } else {
          return new Date(Date.UTC(year, part2 - 1, part1));
        }
      }
    }
    const textMatch = dateStr.match(/^(\d{1,2})[\/\- ]([a-z]{3,})[\/\- ](\d{2,4})$/i);
    if (textMatch) {
      const monthIndex = MONTHS.indexOf(textMatch[2].toLowerCase().substr(0, 3));
      if (monthIndex >= 0) return new Date(Date.UTC(textMatch[3], monthIndex, textMatch[1]));
    }
    const isoMatch = dateStr.match(/^(\d{4})[\/\-](\d{2})[\/\-](\d{2})$/);
    if (isoMatch) return new Date(Date.UTC(isoMatch[1], isoMatch[2] - 1, isoMatch[3]));
    return null;
  }
};

/************************************
 * NORMALIZERS
 ************************************/
function normalizeMemberID(id) {
  if (!id) return '';
  // Remove all spaces, leading zeros, set to lowercase
  return String(id).replace(/\s+/g, '').replace(/^0+/, '').toLowerCase();
}

function normalizeClinician(name) {
  if (!name) return '';
  return name.trim().toLowerCase().replace(/\s+/g, ' ');
}

/************************************
 * HEADER DETECTION / DATA NORMALIZATION
 ************************************/
function normalizeReportData(rawData) {
  const isInsta = rawData[0]?.hasOwnProperty('Pri. Claim No');
  const isOdoo = rawData[0]?.hasOwnProperty('Pri. Claim ID');
  return rawData.map(row => {
    let memberID =
      isInsta ? row['Pri. Patient Insurance Card No'] :
      isOdoo ? row['Pri. Member ID'] :
      row['PatientCardID'];
    if (!memberID) memberID = '';
    return {
      claimID: isInsta ? row['Pri. Claim No'] || '' :
                isOdoo ? row['Pri. Claim ID'] || '' :
                row['ClaimID'] || '',
      memberID: normalizeMemberID(memberID),
      claimDate: isInsta ? row['Encounter Date'] || '' :
                 isOdoo ? row['Adm/Reg. Date'] || '' :
                 row['ClaimDate'] || '',
      clinician: isInsta ? row['Clinician License'] || '' :
                 isOdoo ? row['Admitting License'] || '' :
                 row['Clinician License'] || '',
      department: isInsta ? row['Department'] || '' :
                  isOdoo ? row['Admitting Department'] || '' :
                  row['Clinic'] || '',
      packageName: isInsta ? row['Pri. Payer Name'] || '' :
                   row['Insurance Company'] || '',
      insuranceCompany: isOdoo ? row['Pri. Plan Type'] || '' :
                       isInsta ? row['Pri. Payer Name'] || '' :
                       row['Insurance Company'] || '',
      claimStatus: isInsta ? row['Codification Status'] || '' :
                   isOdoo ? row['Codification Status'] || '' :
                   row['VisitStatus'] || ''
    };
  });
}

/************************************
 * ELIGIBILITY MATCHING
 ************************************/
function prepareEligibilityMap(eligData) {
  const eligMap = new Map();
  eligData.forEach(e => {
    const rawID = e['Card Number / DHA Member ID'] ||
                  e['Card Number'] ||
                  e['_5'] ||
                  e['MemberID'] ||
                  e['Member ID'] ||
                  e['Patient Insurance Card No'];
    if (!rawID) return;
    const memberID = normalizeMemberID(rawID);
    if (!eligMap.has(memberID)) eligMap.set(memberID, []);
    const eligRecord = {
      'Eligibility Request Number': e['Eligibility Request Number'],
      'Card Number / DHA Member ID': rawID,
      'Answered On': e['Answered On'],
      'Ordered On': e['Ordered On'],
      'Status': e['Status'],
      'Clinician': e['Clinician'],
      'Payer Name': e['Payer Name'],
      'Service Category': e['Service Category'],
      'Package Name': e['Package Name']
    };
    eligMap.get(memberID).push(eligRecord);
  });
  return eligMap;
}

function findEligibilityForClaim(eligMap, claimDate, memberID, claimClinicians = []) {
  const normalizedID = normalizeMemberID(memberID);
  const eligList = eligMap.get(normalizedID) || [];
  if (!eligList.length) {
    console.warn(`No eligibility match for Member ID: ${memberID} (normalized: ${normalizedID})`);
    return null;
  }
  console.log(`Eligibility candidates for ${normalizedID}:`, eligList);
  for (const elig of eligList) {
    const eligDate = DateHandler.parse(elig["Answered On"]);
    console.log(`Claim Date: ${claimDate} (${DateHandler.format(claimDate)}), Eligibility Date: ${elig["Answered On"]} (${DateHandler.format(eligDate)})`);
    if (!DateHandler.isSameDay(claimDate, eligDate)) {
      console.log("Date mismatch for claim vs eligibility");
      continue;
    }
    const eligClinician = (elig.Clinician || '').trim();
    if (eligClinician && claimClinicians.length && !claimClinicians.includes(eligClinician)) {
      console.log("Clinician mismatch for claim vs eligibility");
      continue;
    }
    const serviceCategory = (elig['Service Category'] || '').trim();
    const consultationStatus = (elig['Consultation Status'] || '').trim();
    const department = (elig.Department || elig.Clinic || '').toLowerCase();
    const categoryCheck = isServiceCategoryValid(serviceCategory, consultationStatus, department);
    if (!categoryCheck.valid) {
      console.log("Service category mismatch for claim vs eligibility");
      continue;
    }
    if ((elig.Status || '').trim().toLowerCase() !== 'eligible') {
      console.log(`Eligibility status is not 'eligible' (was: '${elig.Status}')`);
      continue;
    }
    console.log("Valid eligibility match found.");
    return elig;
  }
  console.warn(`No full eligibility match for memberID: ${memberID} (${normalizedID}) on date: ${claimDate}`);
  return null;
}

/************************************
 * VALIDATION
 ************************************/
function isServiceCategoryValid(serviceCategory, consultationStatus, rawPackage) {
  if (!serviceCategory) return { valid: true };
  const category = serviceCategory.trim().toLowerCase();
  const pkgRaw = rawPackage || '';
  const pkg = pkgRaw.toLowerCase();

  if (category === 'consultation' && consultationStatus?.toLowerCase() === 'elective') {
    const disallowed = ['dental', 'physio', 'diet', 'occupational', 'speech'];
    if (disallowed.some(term => pkg.includes(term))) {
      return {
        valid: false,
        reason: `Consultation (Elective) cannot include restricted types. Found: "${pkgRaw}"`
      };
    }
    return { valid: true };
  }

  const allowedKeywords = SERVICE_PACKAGE_RULES[serviceCategory];
  if (allowedKeywords && allowedKeywords.length > 0) {
    if (pkg && !allowedKeywords.some(keyword => pkg.includes(keyword))) {
      return {
        valid: false,
        reason: `${serviceCategory} requires related package. Found: "${pkgRaw}"`
      };
    }
  }
  return { valid: true };
}

/************************************
 * UI RENDERING & STATUS
 ************************************/
function renderResults(results, eligMap) {
  resultsContainer.innerHTML = '';
  if (!results || results.length === 0) {
    resultsContainer.innerHTML = '<div class="no-results">No claims to display</div>';
    return;
  }
  // ... modal & table rendering as before ...
}

/************************************
 * PARSING XLSX/CSV
 ************************************/
async function parseExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = function(e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const allRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
        let headerRow = 0, foundHeaders = false;
        while (headerRow < allRows.length && !foundHeaders) {
          const row = allRows[headerRow].map(c => String(c).trim());
          const nonEmptyCells = row.filter(c => c !== '');
          foundHeaders = (nonEmptyCells.length >= 3);
          if (!foundHeaders) headerRow++;
        }
        if (!foundHeaders) headerRow = 0;
        const headers = allRows[headerRow].map(h => String(h).trim());
        const dataRows = allRows.slice(headerRow + 1);
        const jsonData = dataRows.map(row => {
          const obj = {};
          headers.forEach((header, index) => obj[header] = row[index] || '');
          return obj;
        });
        resolve(jsonData);
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(file);
  });
}

async function parseCsvFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = function(e) {
      try {
        const text = e.target.result;
        const workbook = XLSX.read(text, { type: 'string' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const allRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
        let headerRowIndex = -1;
        for (let i = 0; i < 5; i++) {
          const row = allRows[i];
          if (!row) continue;
          const joined = row.join(',').toLowerCase();
          if (joined.includes('claim') && joined.includes('id')) {
            headerRowIndex = i;
            break;
          }
        }
        if (headerRowIndex === -1) throw new Error("Could not detect header row in CSV");
        const headers = allRows[headerRowIndex];
        const dataRows = allRows.slice(headerRowIndex + 1);
        const rawParsed = dataRows.map(row => {
          const obj = {};
          headers.forEach((header, index) => obj[header] = row[index] || '');
          return obj;
        });
        const seen = new Set();
        const uniqueRows = [];
        const claimIdHeader = headers.find(h => h.toLowerCase().includes('claim'));
        rawParsed.forEach(row => {
          const claimID = row[claimIdHeader];
          if (claimID && !seen.has(claimID)) {
            seen.add(claimID);
            uniqueRows.push(row);
          }
        });
        resolve(uniqueRows);
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsText(file);
  });
}

/************************************
 * CLAIM VALIDATION
 ************************************/
function validateReportClaims(reportData, eligMap) {
  if (!reportData) return [];
  return reportData.map(row => {
    if (!row.claimID) return null;
    const memberID = row.memberID;
    const claimDateRaw = row.claimDate;
    const claimDate = DateHandler.parse(claimDateRaw, { preferMDY: lastReportWasCSV });
    const formattedDate = DateHandler.format(claimDate);

    const isVVIP = memberID.startsWith('(vvip)');
    if (isVVIP) {
      return {
        claimID: row.claimID,
        memberID,
        encounterStart: formattedDate,
        packageName: row.packageName || '',
        provider: row.provider || '',
        clinician: row.clinician || '',
        serviceCategory: '',
        consultationStatus: '',
        status: 'VVIP',
        claimStatus: row.claimStatus || '',
        remarks: ['VVIP member, eligibility check bypassed'],
        finalStatus: 'valid',
        fullEligibilityRecord: null
      };
    }
    const hasLeadingZero = memberID.match(/^0+\d+$/);
    const eligibility = findEligibilityForClaim(eligMap, claimDate, memberID, [row.clinician]);
    let status = 'invalid';
    const remarks = [];

    if (!eligibility) {
      remarks.push(`No matching eligibility found for ${memberID} on ${formattedDate}`);
      console.warn('Eligibility not found:', { memberID, claimDateRaw, formattedDate, clinician: row.clinician });
    } else if ((eligibility.Status || '').trim().toLowerCase() !== 'eligible') {
      remarks.push(`Eligibility status: ${eligibility.Status}`);
      console.log("Eligibility status not eligible:", eligibility.Status);
    } else {
      const serviceCategory = eligibility['Service Category']?.trim() || '';
      const consultationStatus = eligibility['Consultation Status']?.trim()?.toLowerCase() || '';
      const matchesCategory = isServiceCategoryValid(serviceCategory, consultationStatus, row.department || row.clinic).valid;
      if (!matchesCategory) {
        remarks.push(`Invalid for category: ${serviceCategory}, department: ${row.department || row.clinic}`);
        console.log('Service category mismatch for claim:', serviceCategory, row.department || row.clinic);
      } else if (!hasLeadingZero) {
        status = 'valid';
      }
    }

    return {
      claimID: row.claimID,
      memberID,
      encounterStart: formattedDate,
      packageName: eligibility?.['Package Name'] || row.packageName || '',
      provider: eligibility?.['Payer Name'] || row.provider || '',
      clinician: eligibility?.['Clinician'] || row.clinician || '',
      serviceCategory: eligibility?.['Service Category'] || '',
      consultationStatus: eligibility?.['Consultation Status'] || '',
      status: eligibility?.Status || '',
      claimStatus: row.claimStatus || '',
      remarks,
      finalStatus: status,
      fullEligibilityRecord: eligibility
    };
  }).filter(r => r);
}

/************************************
 * EXPORT FUNCTIONALITY
 ************************************/
function exportInvalidEntries(results) {
  const invalidEntries = results.filter(r => r && r.finalStatus === 'invalid');
  if (invalidEntries.length === 0) {
    alert('No invalid entries to export.');
    return;
  }
  const exportData = invalidEntries.map(entry => ({
    'Claim ID': entry.claimID,
    'Member ID': entry.memberID,
    'Encounter Date': entry.encounterStart,
    'Package Name': entry.packageName || '',
    'Provider': entry.provider || '',
    'Clinician': entry.clinician || '',
    'Service Category': entry.serviceCategory || '',
    'Consultation Status': entry.consultationStatus || '',
    'Eligibility Status': entry.status || '',
    'Final Status': entry.finalStatus,
    'Remarks': entry.remarks.join('; ')
  }));

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(exportData);
  XLSX.utils.book_append_sheet(wb, ws, 'Invalid Claims');
  XLSX.writeFile(wb, `invalid_claims_${new Date().toISOString().slice(0,10)}.xlsx`);
}

/************************************
 * EVENT HANDLERS
 ************************************/
reportInput.addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  try {
    updateStatus("Loading report file...");
    lastReportWasCSV = file.name.toLowerCase().endsWith('.csv');
    const rawData = lastReportWasCSV ? await parseCsvFile(file) : await parseExcelFile(file);
    reportData = normalizeReportData(rawData).filter(r => r.claimID && String(r.claimID).trim() !== '');
    console.log("--- Report File Uploaded ---", file.name, reportData);
    updateStatus(`Loaded ${reportData.length} report rows`);
    updateProcessButtonState();
  } catch (error) {
    updateStatus("Error loading report file");
    console.error('Report file error:', error);
    processBtn.disabled = true;
    exportInvalidBtn.disabled = true;
  }
});

eligInput.addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  try {
    updateStatus("Loading eligibility file...");
    eligData = await parseExcelFile(file);
    console.log("--- Eligibility File Uploaded ---", file.name, eligData);
    updateStatus(`Loaded ${eligData.length} eligibility records`);
    updateProcessButtonState();
  } catch (error) {
    updateStatus("Error loading eligibility file");
    console.error('Eligibility file error:', error);
    processBtn.disabled = true;
    exportInvalidBtn.disabled = true;
  }
});

processBtn.addEventListener('click', async () => {
  if (!eligData) {
    updateStatus('Missing eligibility file');
    alert('Please upload eligibility file first');
    return;
  }
  if (!reportData) {
    updateStatus('Missing report file');
    alert('Please upload patient report file');
    return;
  }
  try {
    updateStatus('Processing...');
    usedEligibilities.clear();
    const eligMap = prepareEligibilityMap(eligData);
    console.log("--- Eligibility Map Keys ---", Array.from(eligMap.keys()));
    let results = validateReportClaims(reportData, eligMap);
    console.log("--- Pre-filtered Validation Results ---"); results.forEach((r, idx) => console.log(`[${idx}]`, r));
    window.lastValidationResults = results;
    renderResults(results, eligMap);
    updateStatus(`Processed ${results.length} claims`);
    if (results.length === 0) {
      console.warn("No claims processed! Check input files and mapping logic.");
      status.innerHTML = '<span style="color:red;">Troubleshooting: No processed claims. See console for details.</span>';
    }
  } catch (error) {
    updateStatus('Processing failed');
    resultsContainer.innerHTML = `<div class="error">${error.message}</div>`;
    console.error('Processing error:', error);
  }
});

exportInvalidBtn.addEventListener('click', () => {
  if (!window.lastValidationResults) {
    alert('Please run the validation first.');
    return;
  }
  exportInvalidEntries(window.lastValidationResults);
});

document.addEventListener('DOMContentLoaded', () => {
  updateProcessButtonState();
  updateStatus('Ready to process files');
  console.log("--- Eligibility Checker Initialized ---");
});
