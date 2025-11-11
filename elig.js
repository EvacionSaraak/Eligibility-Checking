/*******************************
 * elig.js (updated)
 *
 * Adds a "Invalid only" result-window-only checkbox (enabled by default).
 * The script uses window.lastValidationResults (the last processed results)
 * and a stored eligMap to re-render the displayed results when the checkbox
 * changes. The checkbox filters only the displayed results (does not change
 * exported data or the underlying processed results).
 *******************************/

const SERVICE_PACKAGE_RULES = {
  'Dental Services': ['dental', 'orthodontic'],
  'Physiotherapy': ['physio'],
  'Other OP Services': ['physio', 'diet', 'occupational', 'speech'],
  'Consultation': []  // Special handling below
};
const DATE_KEYS = ['Date', 'On'];
const MONTHS = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"];

// Application state
let xlsData = null;        // parsed & normalized report rows
let eligData = null;       // eligibility sheet as array of arrays (raw) — NOTE: keep as raw rows for header detection
let rawParsedReport = null; // raw parsed sheet result (header detection output)
const usedEligibilities = new Set();
let lastReportWasCSV = false;

// Keep last eligibility map so UI filters can re-render without rebuilding the map
let lastEligMap = null;

// DOM Elements (lookups performed in initializeEventListeners)
let reportInput, eligInput, processBtn, exportInvalidBtn, statusEl, resultsContainer, filterCheckbox, filterStatus, pasteTextarea, pasteBtn, invalidOnlyCheckbox;

/*************************
 * DATE HANDLING UTILITIES (From first script)
 *************************/
const DateHandler = {
  parse: function(input, options = {}) {
    const preferMDY = !!options.preferMDY;
    if (!input) return null;
    if (input instanceof Date) return isNaN(input) ? null : input;
    if (typeof input === 'number') return this._parseExcelDate(input);

    const cleanStr = input.toString().trim().replace(/[,.]/g, '');
    const parsed = this._parseStringDate(cleanStr, preferMDY) || new Date(cleanStr);
    if (isNaN(parsed)) {
      console.warn('Unrecognized date:', input);
      return null;
    }
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
    // Return UTC midnight
    return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
  },

  // PATCHED: Always parse string dates as UTC
  _parseStringDate: function(dateStr, preferMDY = false) {
    if (dateStr.includes(' ')) {
      dateStr = dateStr.split(' ')[0];
    }
    // Matches DD/MM/YYYY or MM/DD/YYYY (ambiguous). We'll disambiguate using preferMDY flag
    const dmyMdyMatch = dateStr.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
    if (dmyMdyMatch) {
      const part1 = parseInt(dmyMdyMatch[1], 10);
      const part2 = parseInt(dmyMdyMatch[2], 10);
      const year = parseInt(dmyMdyMatch[3], 10);

      if (part1 > 12 && part2 <= 12) {
        return new Date(Date.UTC(year, part2 - 1, part1)); // dmy
      } else if (part2 > 12 && part1 <= 12) {
        return new Date(Date.UTC(year, part1 - 1, part2)); // mdy (rare)
      } else {
        if (preferMDY) {
          return new Date(Date.UTC(year, part1 - 1, part2)); // MM/DD/YYYY UTC
        } else {
          return new Date(Date.UTC(year, part2 - 1, part1)); // DD/MM/YYYY UTC
        }
      }
    }

    // Matches 30-Jun-2025 or 30 Jun 2025
    const textMatch = dateStr.match(/^(\d{1,2})[\/\- ]([a-z]{3,})[\/\- ](\d{2,4})$/i);
    if (textMatch) {
      const monthIndex = MONTHS.indexOf(textMatch[2].toLowerCase().substr(0, 3));
      if (monthIndex >= 0) return new Date(Date.UTC(textMatch[3], monthIndex, textMatch[1]));
    }

    // ISO: 2025-07-01
    const isoMatch = dateStr.match(/^(\d{4})[\/\-](\d{2})[\/\-](\d{2})$/);
    if (isoMatch) return new Date(Date.UTC(isoMatch[1], isoMatch[2] - 1, isoMatch[3]));
    return null;
  }
};

/*****************************
 * DATA NORMALIZATION FUNCTIONS (from first)
 *****************************/
function normalizeMemberID(id) {
  if (!id) return "";
  return String(id).replace(/\D/g, "").trim();
}

function normalizeClinician(name) {
  if (!name) return '';
  return name.trim().toLowerCase().replace(/\s+/g, ' ');
}

/*************************
 * Header detection (from first)
 *************************/
function findHeaderRowFromArrays(allRows, maxScan = 10) {
  if (!Array.isArray(allRows) || allRows.length === 0) { return { headerRowIndex: -1, headers: [], rows: [] }; }

  // tokens that commonly appear in the header rows for the supported file types
  const tokens = [
    'pri. claim no', 'pri claim no', 'claimid', 'claim id', 'pri. claim id', 'pri claim id',
    'center name', 'card number', 'card number / dha member id', 'member id', 'patientcardid',
    'pri. patient insurance card no', 'institution', 'facility id', 'mr no.', 'pri. claim id'
  ];

  const scanLimit = Math.min(maxScan, allRows.length);
  let bestIndex = 0;
  let bestScore = 0;

  for (let i = 0; i < scanLimit; i++) {
    const row = allRows[i] || [];
    const joined = row.map(c => (c === null || c === undefined) ? '' : String(c)).join(' ').toLowerCase();

    let score = 0;
    for (const t of tokens) { if (joined.includes(t)) score++; }

    // prefer a row that contains multiple token hits; tie-breaker: earlier row wins
    if (score > bestScore) {
      bestScore = score;
      bestIndex = i;
    }
  }

  // If we found no meaningful header row, default to first row (index 0)
  const headerRowIndex = bestScore > 0 ? bestIndex : 0;
  const rawHeaderRow = allRows[headerRowIndex] || [];

  // normalize headers (trim strings)
  const headers = rawHeaderRow.map(h => (h === null || h === undefined) ? '' : String(h).trim());

  // assemble data rows (everything after headerRowIndex)
  const dataRows = allRows.slice(headerRowIndex + 1);

  // convert to array of objects using detected headers
  const rows = dataRows.map(rowArray => {
    const obj = {};
    for (let c = 0; c < headers.length; c++) {
      const key = headers[c] || `Column${c+1}`;
      obj[key] = rowArray[c] === undefined || rowArray[c] === null ? '' : rowArray[c];
    }
    return obj;
  });
  console.log(headers);
  return { headerRowIndex, headers, rows };
}

/*******************************
 * ELIGIBILITY MATCHING FUNCTIONS (kept & robust)
 *******************************/
function prepareEligibilityMap(rawSheetArray) {
  if (!Array.isArray(rawSheetArray) || rawSheetArray.length === 0) return new Map();

  // Find the row that contains "Eligibility Request Number"
  let headerRowIndex = rawSheetArray.findIndex(row => 
    Array.isArray(row) && row.some(cell => String(cell || '').trim().toLowerCase().includes('eligibility request number'))
  );

  // If not found and input is array-of-objects, we'll handle in other path below
  if (headerRowIndex === -1 && Array.isArray(rawSheetArray[0])) {
    // try looser detection
    headerRowIndex = rawSheetArray.findIndex(row => Array.isArray(row) && row.some(cell => String(cell || '').trim() !== ''));
  }

  // If array-of-arrays input, convert rows to objects using headers
  if (Array.isArray(rawSheetArray[0])) {
    if (headerRowIndex === -1) {
      console.warn("No eligibility header row detected in array-of-arrays; returning empty map");
      return new Map();
    }

    const headers = rawSheetArray[headerRowIndex].map(h => String(h || '').trim());
    const eligMap = new Map();

    for (let i = headerRowIndex + 1; i < rawSheetArray.length; i++) {
      const row = rawSheetArray[i];
      if (!Array.isArray(row)) continue;

      // Skip mostly empty rows
      const blankOrJunkCount = row.filter((v, idx) => {
        const key = headers[idx] || '';
        return v === undefined || v === null || v === '' || key.startsWith('_') || key.toLowerCase().includes('policy');
      }).length;
      if (blankOrJunkCount > headers.length / 2) continue;

      const record = {};
      headers.forEach((h, idx) => record[h] = row[idx] !== undefined ? row[idx] : '');

      // Normalize member ID
      const idCandidates = [
        'Card Number / DHA Member ID', 'Card Number', 'MemberID', 'Member ID',
        'Patient Insurance Card No', 'Policy1', 'Policy 1', 'PatientCardID'
      ];
      let rawMemberID = '';
      for (const k of idCandidates) {
        if (Object.prototype.hasOwnProperty.call(record, k) && record[k]) {
          rawMemberID = String(record[k]).trim();
          break;
        }
      }
      if (!rawMemberID) continue;
      const memberID = normalizeMemberID(rawMemberID);
      if (!memberID) continue;

      if (!eligMap.has(memberID)) eligMap.set(memberID, []);
      eligMap.get(memberID).push(record);
    }

    console.log("Elig Map (arrays) built successfully. Members:", eligMap.size);
    return eligMap;
  }

  // Otherwise assume array-of-objects (headers already mapped)
  const eligMap = new Map();
  const idCandidatesObj = ['Card Number / DHA Member ID', 'Card Number', '_5', 'MemberID', 'Member ID', 'Patient Insurance Card No', 'PatientCardID'];

  rawSheetArray.forEach(e => {
    if (!e || typeof e !== 'object') return;
    let rawMemberID = '';
    for (const k of idCandidatesObj) {
      if (Object.prototype.hasOwnProperty.call(e, k) && e[k]) {
        rawMemberID = String(e[k]).trim();
        break;
      }
    }
    if (!rawMemberID) return;
    const memberID = normalizeMemberID(rawMemberID);
    if (!memberID) return;

    if (!eligMap.has(memberID)) eligMap.set(memberID, []);
    eligMap.get(memberID).push(e);
  });

  console.log("Elig Map (objects) built successfully. Members:", eligMap.size);
  return eligMap;
}

function checkClinicianMatch(claimClinicians, eligClinician) {
  if (!eligClinician || !claimClinicians?.length) return true;
  const normElig = normalizeClinician(eligClinician);
  return claimClinicians.some(c => normalizeClinician(c) === normElig);
}

function isServiceCategoryValid(serviceCategory, consultationStatus, rawPackage) {
  if (!serviceCategory) return { valid: true };

  const category = serviceCategory.trim().toLowerCase();
  const pkgRaw = rawPackage || '';
  const pkg = pkgRaw.toLowerCase();

  if (category === 'consultation' && consultationStatus?.toLowerCase() === 'elective') {
    const disallowed = ['dental', 'physio', 'diet', 'occupational', 'speech'];
    if (disallowed.some(term => pkg.includes(term))) {
      return { valid: false, reason: `Consultation (Elective) cannot include restricted service types. Found: "${pkgRaw}"` };
    }
    return { valid: true };
  }

  const allowedKeywords = SERVICE_PACKAGE_RULES[serviceCategory];
  if (allowedKeywords && allowedKeywords.length > 0) {
    if (pkg && !allowedKeywords.some(keyword => pkg.includes(keyword))) {
      return { valid: false, reason: `${serviceCategory} category requires related package. Found: "${pkgRaw}"` };
    }
  }

  return { valid: true };
}

/* Ensure lookups normalize the incoming member id before Map.get */
function findEligibilityForClaim(eligMap, claimDate, memberID, claimClinicians = []) {
  const normalizedID = normalizeMemberID(memberID || '');
  const eligList = eligMap.get(normalizedID) || [];
  if (!eligList.length) return null;
  console.log(`[Diagnostics] Searching eligibilities for member "${memberID}" (normalized: "${normalizedID}")`);
  console.log(`[Diagnostics] Claim date: ${claimDate} (${DateHandler.format(claimDate)}), Claim clinicians: ${JSON.stringify(claimClinicians)}`);
  for (const elig of eligList) {
    console.log(`[Diagnostics] Checking eligibility ${elig["Eligibility Request Number"] || "(unknown)"}:`);
    const eligDate = DateHandler.parse(elig["Answered On"]);
    if (!DateHandler.isSameDay(claimDate, eligDate)) {
      console.log(`  ❌ Date mismatch: claim ${DateHandler.format(claimDate)} vs elig ${DateHandler.format(eligDate)}`);
      continue;
    }
    const eligClinician = (elig.Clinician || '').trim();
    if (eligClinician && claimClinicians.length && !claimClinicians.includes(eligClinician)) {
      console.log(`  ❌ Clinician mismatch: claim clinicians ${JSON.stringify(claimClinicians)} vs elig clinician "${eligClinician}"`);
      continue;
    }
    const serviceCategory = (elig['Service Category'] || '').trim();
    const consultationStatus = (elig['Consultation Status'] || '').trim();
    const department = (elig.Department || elig.Clinic || '').toLowerCase();
    const categoryCheck = isServiceCategoryValid(serviceCategory, consultationStatus, department);
    if (!categoryCheck.valid) {
      console.log(`  ❌ Service category mismatch: claim dept "${department}" not valid for category "${serviceCategory}" / consult "${consultationStatus}"`);
      continue;
    }
    if ((elig.Status || '').toLowerCase() !== 'eligible') {
      console.log(`  ❌ Status mismatch: expected Eligible, got "${elig.Status}"`);
      continue;
    }
    console.log(`  ✅ Eligibility match found: ${elig["Eligibility Request Number"]}`);
    return elig;
  }
  console.log(`[Diagnostics] No matching eligibility passed all checks for member "${memberID}"`);
  return null;
}

/*************************
 * Normalize & validate report rows (use working version's normalizeReportData)
 *************************/
function normalizeReportData(rawData) {
  if (!rawData) return [];

  if (Array.isArray(rawData) && rawData.length > 0 && typeof rawData[0] === 'object' && !rawData.headers) {
    const sample = rawData[0];
    const isInsta = sample.hasOwnProperty('Pri. Claim No');
    const isOdoo = sample.hasOwnProperty('Pri. Claim ID');
    return rawData.map(row => {
      if (isInsta) {
        return {
          claimID: row['Pri. Claim No'] || '',
          memberID: row['Pri. Patient Insurance Card No'] || '',
          claimDate: row['Encounter Date'] || '',
          clinician: row['Clinician License'] || '',
          department: row['Department'] || '',
          packageName: row['Pri. Payer Name'] || '',
          insuranceCompany: row['Pri. Payer Name'] || '',
          claimStatus: row['Codification Status'] || ''
        };
      } else if (isOdoo) {
        return {
          claimID: row['Pri. Claim ID'] || '',
          memberID: row['Pri. Member ID'] || '',
          claimDate: row['Adm/Reg. Date'] || '',
          clinician: row['Admitting License'] || '',
          department: row['Admitting Department'] || '',
          insuranceCompany: row['Pri. Plan Type'] || '',
          claimStatus: row['Codification Status'] || ''
        };
      } else {
        return {
          claimID: row['ClaimID'] || '',
          memberID: row['PatientCardID'] || '',
          claimDate: row['ClaimDate'] || '',
          clinician: row['Clinician License'] || '',
          packageName: row['Insurance Company'] || '',
          insuranceCompany: row['Insurance Company'] || '',
          department: row['Clinic'] || '',
          claimStatus: row['VisitStatus'] || ''
        };
      }
    });
  }

  const rows = rawData.rows || [];
  const headers = rawData.headers || [];

  function getField(obj, candidates) {
    for (const k of candidates) {
      if (obj && Object.prototype.hasOwnProperty.call(obj, k) && obj[k] !== '' && obj[k] !== null && obj[k] !== undefined) return obj[k];
    }
    return '';
  }

  return rows.map(r => {
    const isInsta = !!(r['Pri. Claim No'] || r['Pri. Patient Insurance Card No']);
    const isOdoo = !!r['Pri. Claim ID'];

    if (isInsta) {
      return {
        claimID: r['Pri. Claim No'] || '',
        memberID: r['Pri. Patient Insurance Card No'] || '',
        claimDate: r['Encounter Date'] || '',
        clinician: r['Clinician License'] || '',
        department: r['Department'] || '',
        packageName: r['Pri. Payer Name'] || '',
        insuranceCompany: r['Pri. Payer Name'] || '',
        claimStatus: r['Codification Status'] || ''
      };
    } else if (isOdoo) {
      return {
        claimID: r['Pri. Claim ID'] || '',
        memberID: r['Pri. Member ID'] || '',
        claimDate: r['Adm/Reg. Date'] || '',
        clinician: r['Admitting License'] || '',
        department: r['Admitting Department'] || '',
        insuranceCompany: r['Pri. Plan Type'] || '',
        claimStatus: r['Codification Status'] || ''
      };
    } else {
      const out = {
        claimID: r['ClaimID'] || r['Pri. Claim No'] || r['Pri. Claim ID'] || getField(r, ['ClaimID','Pri. Claim No','Pri. Claim ID','Claim ID','Pri. Claim ID']) || '',
        memberID: r['Pri. Member ID'] || r['Pri. Patient Insurance Card No'] || r['PatientCardID'] || getField(r, ['PatientCardID','Patient Insurance Card No','Card Number / DHA Member ID']) || '',
        claimDate: r['Encounter Date'] || r['Adm/Reg. Date'] || r['ClaimDate'] || getField(r, ['Encounter Date','ClaimDate','Adm/Reg. Date','Date']) || '',
        clinician: r['Clinician License'] || r['Admitting License'] || r['OrderDoctor'] || getField(r, ['Clinician License','Clinician','Admitting License','OrderDoctor']) || '',
        department: r['Department'] || r['Clinic'] || r['Admitting Department'] || getField(r, ['Department','Clinic','Admitting Department']) || '',
        packageName: r['Pri. Payer Name'] || r['Insurance Company'] || r['Pri. Sponsor'] || getField(r, ['Pri. Payer Name','Insurance Company','Pri. Plan Type','Package','Pri. Sponsor']) || '',
        insuranceCompany: r['Pri. Payer Name'] || r['Insurance Company'] || getField(r, ['Payer Name','Insurance Company','Pri. Payer Name']) || '',
        claimStatus: r['Codification Status'] || r['VisitStatus'] || r['Status'] || getField(r, ['Codification Status','VisitStatus','Status','Claim Status']) || ''
      };

      if (!out.memberID) {
        for (const h of headers) {
          const val = r[h];
          if (val && String(h).toLowerCase().includes('card')) { out.memberID = val; break; }
        }
      }
      if (!out.claimID) {
        for (const h of headers) {
          const val = r[h];
          if (val && String(h).toLowerCase().includes('claim')) { out.claimID = val; break; }
        }
      }
      return out;
    }
  });
}

function validateReportClaims(reportDataArray, eligMap, reportType) {
  const results = [];
  for (let i = 0; i < reportDataArray.length; i++) {
    const row = reportDataArray[i];
    const claimID = String(row.claimID || '').trim();
    if (!claimID) continue;

    const rawMemberID = String(row.memberID || '').trim();
    if (!rawMemberID) continue;
    const memberID = normalizeMemberID(rawMemberID);

    let insurance = (row.insuranceCompany || '').trim();
    const claimDate = DateHandler.parse(row.claimDate, { preferMDY: lastReportWasCSV });
    if (!claimDate) continue;
    const formattedDate = DateHandler.format(claimDate);

    if (memberID.startsWith('(VVIP)')) {
      results.push({ claimID, memberID, encounterStart: formattedDate, status: 'VVIP', finalStatus: 'valid', remarks: ['VVIP member, eligibility check bypassed'], fullEligibilityRecord: null });
      continue;
    }

    const eligibility = findEligibilityForClaim(eligMap, claimDate, memberID, [row.clinician]);
    let finalStatus = 'invalid', remarks = [];
    if (!eligibility) remarks.push(`No matching eligibility found for ${memberID} on ${formattedDate}`);
    else if (eligibility.Status?.toLowerCase() === 'eligible') {
      const categoryCheck = isServiceCategoryValid(eligibility['Service Category'], eligibility['Consultation Status'], (row.department || '').toLowerCase());
      if (categoryCheck.valid) finalStatus = 'valid';
      else remarks.push(categoryCheck.reason || 'Service category mismatch');
    } else remarks.push(`Eligibility status: ${eligibility.Status}`);

    results.push({
      claimID, memberID, encounterStart: formattedDate,
      packageName: eligibility?.['Package Name'] || row.packageName || '',
      provider: insurance,
      clinician: eligibility?.Clinician || row.clinician || '',
      serviceCategory: eligibility?.['Service Category'] || '',
      consultationStatus: eligibility?.['Consultation Status'] || '',
      status: eligibility?.Status || '',
      claimStatus: row.claimStatus || '',
      remarks, finalStatus, fullEligibilityRecord: eligibility
    });
  }
  return results;
}

/*********************
 * Diagnostic logger used when a claim has no matching eligibility (from first)
 *********************/
function logNoEligibilityMatch(sourceType, claimSummary, memberID, parsedClaimDate, claimClinicians, eligMap) {
  try {
    const normalizedID = normalizeMemberID(memberID);
    const eligList = eligMap.get(normalizedID) || [];
    console.groupCollapsed(`[Diagnostics] No eligibility match (${sourceType}) — member: "${memberID}" (normalized: "${normalizedID}")`);
    console.log('Claim / row summary:', claimSummary);
    console.log('Parsed claim date object:', parsedClaimDate, 'Formatted:', DateHandler.format(parsedClaimDate));
    console.log('Claim clinicians:', claimClinicians || []);
    if (!eligList || eligList.length === 0) {
      console.warn('No eligibility records found for this member ID in eligMap.');
    } else {
      console.log(`Found ${eligList.length} eligibility record(s) for member "${memberID}":`);
      eligList.forEach((e, i) => {
        const answeredOnRaw = e['Answered On'] || e['Ordered On'] || '';
        const answeredOnParsed = DateHandler.parse(answeredOnRaw);
        console.log(`#${i+1}`, {
          'Eligibility Request Number': e['Eligibility Request Number'],
          'Answered On (raw)': answeredOnRaw,
          'Answered On (parsed)': answeredOnParsed,
          'Ordered On': e['Ordered On'],
          'Status': e['Status'],
          'Clinician': e['Clinician'],
          'Payer Name': e['Payer Name'],
          'Service Category': e['Service Category'],
          'Package Name': e['Package Name'],
          'Used': usedEligibilities.has(e['Eligibility Request Number'])
        });
      });
    }
    console.groupEnd();
  } catch (err) {
    console.error('Error in logNoEligibilityMatch diagnostic logger:', err);
  }
}

/*********************
 * FILE PARSING FUNCTIONS (kept from second script)
 *********************/
async function parseExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = function(e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const allRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

        // Helper: detect likely title rows
        function isLikelyTitleRow(row) {
          const emptyCount = row.filter(c => String(c).trim() === '').length;
          return emptyCount > 4; // skip if more than 4 empty cells
        }

        // Detect header row dynamically
        let headerRow = 0;
        let foundHeaders = false;

        while (headerRow < allRows.length && !foundHeaders) {
          const currentRow = allRows[headerRow].map(c => String(c).trim());

          // Skip likely title rows
          if (isLikelyTitleRow(currentRow)) {
            headerRow++;
            continue;
          }

          // Check for known headers
          if (currentRow.some(cell => cell.includes('Pri. Claim No')) ||
              currentRow.some(cell => cell.includes('Pri. Claim ID')) ||
              currentRow.some(cell => cell.includes('Card Number / DHA Member ID'))) {
            foundHeaders = true;
            break;
          }

          // Fallback: treat row with >= 3 non-empty cells as header
          const nonEmptyCells = currentRow.filter(c => c !== '');
          if (nonEmptyCells.length >= 3) {
            foundHeaders = true;
            break;
          }
          headerRow++;
        }

        // Default to first row if none detected
        if (!foundHeaders) headerRow = 0;

        // Trim headers
        const headers = allRows[headerRow].map(h => String(h).trim());
        console.log(`Headers: ${headers}`);

        // Extract data rows
        const dataRows = allRows.slice(headerRow + 1);

        // Map rows to objects
        const jsonData = dataRows.map(row => {
          const obj = {};
          headers.forEach((header, index) => {
            obj[header] = row[index] || '';
          });
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

function parseCsvFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = function(e) {
      try {
        const text = e.target.result;
        const workbook = XLSX.read(text, { type: 'string' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const allRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

        // Try to detect header row in first few rows
        let headerRowIndex = -1;
        for (let i = 0; i < Math.min(5, allRows.length); i++) {
          const row = allRows[i];
          if (!row) continue;
          const joined = row.join(',').toLowerCase();
          if (joined.includes('pri. claim no') || joined.includes('claimid') || joined.includes('claim id')) {
            headerRowIndex = i;
            break;
          }
        }

        if (headerRowIndex === -1) {
          const detection = findHeaderRowFromArrays(allRows, 20);
          resolve(detection);
          return;
        }

        const headers = allRows[headerRowIndex].map(h => String(h || '').trim());
        const dataRows = allRows.slice(headerRowIndex + 1);
        const rows = dataRows.map(row => {
          const obj = {};
          headers.forEach((header, index) => { obj[header] = row[index] || ''; });
          return obj;
        });

        // Deduplicate based on claim ID if possible
        const claimIdHeader = headers.find(h =>
          h.toLowerCase().replace(/\s+/g, '') === 'claimid' ||
          h.toLowerCase().includes('claim')
        );

        if (!claimIdHeader) {
          resolve({ headerRowIndex, headers, rows });
          return;
        }

        const seen = new Set();
        const uniqueRows = [];
        rows.forEach(row => {
          const claimID = row[claimIdHeader];
          if (claimID && !seen.has(claimID)) { seen.add(claimID); uniqueRows.push(row); }
        });

        resolve({ headerRowIndex, headers, rows: uniqueRows });
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsText(file);
  });
}

function parseCsvText(text) {
  return new Promise((resolve, reject) => {
    try {
      const clean = (text || '').replace(/^\uFEFF/, '');
      const wb = XLSX.read(clean, { type: 'string' });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const allRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
      const detection = findHeaderRowFromArrays(allRows, 20);
      resolve(detection);
    } catch (err) {
      reject(err);
    }
  });
}
// Insert this function into elig.js (for example right after the escapeHtml function).
// It defines summarizeAndDisplayCounts which was being called when eligibility/report
// files are loaded. This prevents the ReferenceError and updates the status element.
function summarizeAndDisplayCounts() {
  try {
    const eligCount = Array.isArray(eligData) ? eligData.length : 0;

    // Ensure xlsData exists; if not but rawParsedReport exists try to normalize it now
    if ((!Array.isArray(xlsData) || xlsData.length === 0) && rawParsedReport) {
      try {
        const normalized = normalizeReportData(rawParsedReport);
        xlsData = normalized.filter(r => r && r.claimID && String(r.claimID).trim() !== '');
      } catch (e) {
        console.warn('summarizeAndDisplayCounts: failed to normalize report for counting', e);
      }
    }

    const claimCount = Array.isArray(xlsData) ? xlsData.length : 0;

    if (statusEl) {
      statusEl.textContent = `Loaded ${eligCount} eligibilities, ${claimCount} claims — Ready to process files`;
    }
  } catch (err) {
    console.error('summarizeAndDisplayCounts error', err);
  }
}

/*************************
 * UI RENDERING FUNCTIONS (kept from second but updated with Bootstrap)
 *************************/
function escapeHtml(s) {
  if (s === null || s === undefined) return '';
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#x27;');
}

/* Helper to get displayed results according to "Invalid only" checkbox */
function getDisplayedResultsFromStored(results) {
  const raw = results || window.lastValidationResults || [];
  const invalidOnly = (invalidOnlyCheckbox && invalidOnlyCheckbox.checked) ? true : false;
  if (!invalidOnly) return raw;
  return raw.filter(r => r && r.finalStatus === 'invalid');
}

/*************************
 * renderResults (updated): color entire rows etc.
 *************************/
function renderResults(results, eligMap) {
  if (!resultsContainer) return;
  resultsContainer.innerHTML = '';

  if (!results || results.length === 0) {
    resultsContainer.innerHTML = '<div class="text-muted">No claims to display</div>';
    return;
  }

  // Container for table with Bootstrap responsive wrapper
  const tableContainer = document.createElement('div');
  tableContainer.className = 'table-responsive analysis-results';

  // Build table using Bootstrap table classes plus existing shared-table class
  const table = document.createElement('table');
  table.className = 'table table-sm table-striped table-hover shared-table';

  // Header
  const thead = document.createElement('thead');
  thead.innerHTML = `
    <tr>
      <th>Claim ID</th>
      <th>Member ID</th>
      <th>Encounter Date</th>
      <th>Package</th>
      <th>Provider</th>
      <th>Clinician</th>
      <th>Service Category</th>
      <th>Status</th>
      <th class="wrap-col">Remarks</th>
      <th>Details</th>
    </tr>
  `;
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  const statusCounts = { valid: 0, invalid: 0, unknown: 0 };
  let processedRows = 0;

  // mapping from our semantic finalStatus -> Bootstrap contextual table class
  const finalStatusToBootstrap = {
    valid: 'table-success',
    invalid: 'table-danger',
    unknown: 'table-warning'
  };

  results.forEach((result, index) => {
    if (!result.memberID || result.memberID.toString().trim() === '') return;
    const statusToCheck = (result.claimStatus || result.status || result.fullEligibilityRecord?.Status || '')
      .toString()
      .trim()
      .toLowerCase();
    if (statusToCheck === 'not seen') return;

    if (result.finalStatus && statusCounts.hasOwnProperty(result.finalStatus)) {
      statusCounts[result.finalStatus]++;
    }

    const row = document.createElement('tr');

    // Add semantic class (keeps legacy colors from tables.css) and Bootstrap contextual class
    const finalStatus = (result.finalStatus || '').toString().toLowerCase();
    if (finalStatus) {
      row.classList.add(finalStatus);
      const bs = finalStatusToBootstrap[finalStatus];
      if (bs) row.classList.add(bs);
    }

    // Tag-based coloring: detect provider keywords for Daman/Thiqa and apply tag classes
    const provider = (result.provider || result.insuranceCompany || '').toString().toLowerCase();
    if (provider.includes('daman')) {
      row.classList.add('daman-only');
    } else if (provider.includes('thiqa')) {
      row.classList.add('thiqa-only');
    }

    if ((result.finalStatus || '').toLowerCase() === 'vvip' || (result.status || '').toString().toLowerCase() === 'vvip') {
      row.classList.add('selected');
    }

    const statusBadge = result.status
      ? `<span class="badge ${result.status.toString().toLowerCase() === 'eligible' ? 'bg-success' : 'bg-danger'}">${escapeHtml(result.status)}</span>`
      : '';

    const remarksHTML = result.remarks && result.remarks.length > 0
      ? result.remarks.map(r => `<div>${escapeHtml(r)}</div>`).join('')
      : '<div class="source-note">No remarks</div>';

    let detailsCellHtml = '<div class="source-note">N/A</div>';
    if (result.fullEligibilityRecord?.['Eligibility Request Number']) {
      detailsCellHtml = `<button class="btn btn-sm btn-outline-primary eligibility-details" data-index="${index}" data-claimdate="${escapeHtml(result.encounterStart)}">${escapeHtml(result.fullEligibilityRecord['Eligibility Request Number'])}</button>`;
    } else if (eligMap && typeof eligMap.get === 'function' && (eligMap.get(result.memberID) || []).length) {
      // include claim date so "View All" can highlight rows relative to this claim
      detailsCellHtml = `<button class="btn btn-sm btn-outline-secondary show-all-eligibilities" data-member="${escapeHtml(result.memberID)}" data-claimdate="${escapeHtml(result.encounterStart)}">View All</button>`;
    }

    row.innerHTML = `
      <td>${escapeHtml(result.claimID)}</td>
      <td>${escapeHtml(result.memberID)}</td>
      <td>${escapeHtml(result.encounterStart)}</td>
      <td class="description-col">${escapeHtml(result.packageName)}</td>
      <td class="description-col">${escapeHtml(result.provider)}</td>
      <td class="description-col">${escapeHtml(result.clinician)}</td>
      <td class="description-col">${escapeHtml(result.serviceCategory)}</td>
      <td class="description-col">${statusBadge}</td>
      <td class="wrap-col">${remarksHTML}</td>
      <td>${detailsCellHtml}</td>
    `;

    tbody.appendChild(row);
    processedRows++;
  });

  table.appendChild(tbody);
  tableContainer.appendChild(table);
  resultsContainer.appendChild(tableContainer);

  // Summary counts shown above the table using Bootstrap badges
  const summary = document.createElement('div');
  summary.className = 'loaded-count mb-2';
  summary.innerHTML = `
    Processed ${processedRows} claims:
    <span class="badge bg-success ms-2">${statusCounts.valid} valid</span>
    <span class="badge bg-secondary ms-1">${statusCounts.unknown} unknown</span>
    <span class="badge bg-danger ms-1">${statusCounts.invalid} invalid</span>
  `;
  resultsContainer.prepend(summary);

  // Ensure modal wiring for details and view-all buttons
  initEligibilityModal(results, lastEligMap);

  // Small accessibility hint: make results container focusable and move focus
  resultsContainer.setAttribute('tabindex', '-1');
  resultsContainer.focus();
}

/*************************
 * initEligibilityModal (Bootstrap-flavored) - unchanged
 *************************/
// Replace the existing initEligibilityModal function with this version,
// and add the helper generateAndSendDebugLog below. This adds a "Send debug log"
// button to the modal header which gathers relevant context and downloads the log
// (and copies to clipboard) so you can paste or attach it to an issue/email.

function initEligibilityModal(results, eligMap) {
  // Ensure modal exists (Bootstrap-flavored modal markup)
  if (!document.getElementById("modalOverlay")) {
    const modalHtml = `
      <div id="modalOverlay" class="modal" tabindex="-1" aria-hidden="true">
        <div class="modal-dialog modal-xl modal-dialog-centered">
          <div class="modal-content">
            <div class="modal-header d-flex align-items-center">
              <h5 class="modal-title me-auto">Eligibility Details</h5>
              <!-- Debug button to generate a diagnostic log -->
              <button type="button" class="btn btn-sm btn-outline-info me-2" id="modalDebugBtn" title="Generate debug log for this modal" style="display:none;">
                <i class="bi bi-bug-fill"></i> Send debug log
              </button>
              <button type="button" class="btn-close" id="modalCloseBtn" aria-label="Close"></button>
            </div>
            <div class="modal-body p-0">
              <div id="modalTable" class="p-3" style="overflow:auto; max-height:70vh;"></div>
            </div>
          </div>
        </div>
      </div>
    `;
    document.body.insertAdjacentHTML("beforeend", modalHtml);

    // Wire up close behavior
    const overlay = document.getElementById("modalOverlay");
    const closeBtn = document.getElementById("modalCloseBtn");
    closeBtn.addEventListener('click', hideModal);
    overlay.addEventListener('click', function (e) {
      if (e.target === overlay) hideModal();
    });
    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape') {
        const ov = document.getElementById('modalOverlay');
        if (ov && ov.style.display && ov.style.display !== 'none') hideModal();
      }
    });

    // Debug button handler (calls helper to build log and download/send it)
    const debugBtn = document.getElementById('modalDebugBtn');
    debugBtn.addEventListener('click', () => {
      const ctx = window.__elig_current_debug || null;
      generateAndSendDebugLog(ctx, results, eligMap);
    });
  }

  // DETAILS button: show single eligibility record
  document.querySelectorAll(".eligibility-details").forEach(btn => {
    // remove previous handlers to avoid double-binding
    btn.onclick = null;
    btn.addEventListener('click', function () {
      const index = parseInt(this.dataset.index, 10);
      const result = results[index];
      if (!result?.fullEligibilityRecord) return;
      const record = result.fullEligibilityRecord;

      // derive claim date from the result (formatted string)
      const claimDateStr = this.dataset.claimdate || result.encounterStart || '';
      const claimDate = claimDateStr ? DateHandler.parse(claimDateStr) : null;

      // set debug context so the debug button knows what to include
      window.__elig_current_debug = {
        mode: 'single',
        member: result.memberID,
        claimDate: claimDateStr || '',
        record,
        resultIndex: index
      };

      // make debug button visible
      const debugBtn = document.getElementById('modalDebugBtn');
      if (debugBtn) debugBtn.style.display = '';

      // pass claimDate into formatEligibilityDetails so it can highlight date rows
      document.getElementById("modalTable").innerHTML = formatEligibilityDetails(record, result.memberID, claimDate);
      showModal();
    });
  });

  // VIEW ALL button: show all eligibilities for member, highlight by claim date if provided
  document.querySelectorAll(".show-all-eligibilities").forEach(btn => {
    btn.onclick = null;
    btn.addEventListener('click', function () {
      const member = this.dataset.member;
      // dataset.claimdate may be present (passed from renderResults)
      const claimDateStr = this.dataset.claimdate || '';
      const claimDate = claimDateStr ? DateHandler.parse(claimDateStr) : null;

      const list = (typeof eligMap.get === 'function') ? (eligMap.get(member) || []) : [];
      const modalTable = document.getElementById("modalTable");

      // set debug context to include member/list + claim date for "send debug"
      window.__elig_current_debug = {
        mode: 'list',
        member,
        claimDate: claimDateStr || '',
        listSnapshot: list.slice(0, 200) // limit size to avoid huge logs
      };

      // make debug button visible
      const debugBtn = document.getElementById('modalDebugBtn');
      if (debugBtn) debugBtn.style.display = '';

      if (!list.length) {
        modalTable.innerHTML = `<div class="p-3">No eligibilities found for <strong>${escapeHtml(member)}</strong></div>`;
        showModal();
        return;
      }

      // Build a Bootstrap-styled table of eligibilities and highlight rows:
      let html = `<h6 class="px-3 pt-3">Eligibilities for ${escapeHtml(member)}</h6>
        <div class="table-responsive px-3 pb-3">
          <table class="table table-sm table-striped table-bordered mb-0">
            <thead class="table-light">
              <tr>
                <th style="min-width:38px">#</th>
                <th>Request No</th>
                <th>Answered On</th>
                <th>Status</th>
                <th>Clinician</th>
                <th>Service Category</th>
                <th>Package Name</th>
              </tr>
            </thead>
            <tbody>`;

      list.forEach((rec, idx) => {
        const answeredOnRaw = rec['Answered On'] || rec['Ordered On'] || '';
        const eligDate = DateHandler.parse(answeredOnRaw);
        let trClass = '';
        if (claimDate && eligDate) {
          if (DateHandler.isSameDay(claimDate, eligDate)) trClass = 'table-warning';
          else trClass = 'table-danger';
        }

        html += `<tr class="${trClass}">
          <td>${idx + 1}</td>
          <td>${escapeHtml(rec['Eligibility Request Number'] || '')}</td>
          <td>${escapeHtml(answeredOnRaw || '')}</td>
          <td>${escapeHtml(rec['Status'] || '')}</td>
          <td>${escapeHtml(rec['Clinician'] || '')}</td>
          <td>${escapeHtml(rec['Service Category'] || '')}</td>
          <td>${escapeHtml(rec['Package Name'] || '')}</td>
        </tr>`;
      });

      html += `</tbody></table></div>`;
      modalTable.innerHTML = html;
      showModal();
    });
  });

  // Helper to show the modal (manage display manually)
  function showModal() {
    const overlay = document.getElementById("modalOverlay");
    if (!overlay) return;
    overlay.style.display = 'flex';
    overlay.setAttribute('aria-hidden', 'false');
    setTimeout(() => overlay.classList.add('show'), 10);
    const focusable = overlay.querySelectorAll('button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])');
    if (focusable.length) focusable[0].focus();
  }

  // ensure hideModal exists (fallback)
  if (typeof hideModal !== 'function') {
    window.hideModal = function () {
      const overlay = document.getElementById("modalOverlay");
      if (!overlay) return;
      overlay.classList.remove('show');
      overlay.style.display = 'none';
      overlay.setAttribute('aria-hidden', 'true');
    };
  }
}

/* Helper: gather debug info and prompt download + copy to clipboard.
   ctx is the object we set on window.__elig_current_debug (mode/member/claimDate/record/listSnapshot).
   results and eligMap are passed for extra context. */
function generateAndSendDebugLog(ctx, results, eligMap) {
  try {
    const timestamp = new Date().toISOString();
    const env = {
      timestamp,
      pageUrl: window.location.href,
      userAgent: navigator.userAgent,
      platform: navigator.platform,
      viewport: { width: window.innerWidth, height: window.innerHeight }
    };

    const payload = {
      env,
      context: ctx || null,
      lastValidationResultsCount: Array.isArray(window.lastValidationResults) ? window.lastValidationResults.length : 0,
      lastEligMapSize: (lastEligMap && typeof lastEligMap.size === 'number') ? lastEligMap.size : (eligMap && typeof eligMap.size === 'number' ? eligMap.size : null),
      // include a small sample of lastValidationResults if present
      lastValidationSample: (window.lastValidationResults && Array.isArray(window.lastValidationResults)) ? window.lastValidationResults.slice(0,50) : []
    };

    // If ctx includes member, include elig entries for that member (limit to 200)
    if (ctx && ctx.member && (typeof eligMap?.get === 'function')) {
      const memberKey = normalizeMemberID(ctx.member);
      const memberEntries = eligMap.get(memberKey) || [];
      payload.memberEligibilities = memberEntries.slice(0,200);
    }

    const text = JSON.stringify(payload, null, 2);

    // Trigger download
    const blob = new Blob([text], { type: 'application/json' });
    const filename = `eligibility-debug-${timestamp.replace(/[:.]/g,'-')}.json`;
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);

    // Copy to clipboard for convenience if allowed
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).catch(() => {/* ignore */});
    }

    // Provide user feedback in modal
    const modalTable = document.getElementById('modalTable');
    if (modalTable) {
      const notice = document.createElement('div');
      notice.className = 'alert alert-success mt-2';
      notice.textContent = `Debug log prepared and downloaded as ${filename}. A sample was copied to clipboard. Attach this file to your issue.`;
      modalTable.prepend(notice);
      setTimeout(() => { if (notice.parentNode) notice.remove(); }, 8000);
    }
  } catch (err) {
    console.error('Failed to generate debug log', err);
    alert('Failed to create debug log: ' + (err && err.message ? err.message : String(err)));
  }
}
function hideModal() { const overlay = document.getElementById("modalOverlay"); if (overlay) overlay.style.display = "none"; }

function formatEligibilityDetails(record, memberID) {
  if (!record) return '<div>No details</div>';

  // Header with member and status badge
  const status = (record.Status || '').toString();
  const statusClass = status.toLowerCase() === 'eligible' ? 'status-badge eligible' : 'status-badge ineligible';
  let html = `<div class="mb-2"><strong>Member:</strong> ${escapeHtml(memberID)} <span class="${statusClass}" style="margin-left:8px;">${escapeHtml(status)}</span></div>`;

  // Start details table
  html += '<table class="eligibility-details"><tbody>';

  // Render known important fields first (if present) in a predictable order
  const preferredKeys = [
    'Eligibility Request Number', 'Card Number / DHA Member ID', 'Answered On', 'Ordered On',
    'Status', 'Clinician', 'Payer Name', 'Service Category', 'Package Name'
  ];

  const used = new Set();

  preferredKeys.forEach(key => {
    if (Object.prototype.hasOwnProperty.call(record, key)) {
      const raw = record[key];
      if (raw === undefined || raw === null || raw === '') return;
      used.add(key);
      let disp = raw;
      if (typeof raw === 'string' && (key.includes('Date') || key.toLowerCase().includes('answered') || key.toLowerCase().includes('ordered'))) {
        const parsed = DateHandler.parse(raw);
        disp = parsed ? DateHandler.format(parsed) : raw;
      }
      html += `<tr><th>${escapeHtml(key)}</th><td>${escapeHtml(String(disp))}</td></tr>`;
    }
  });

  // Render any remaining fields
  Object.keys(record).forEach(key => {
    if (used.has(key)) return;
    const raw = record[key];
    if (raw === undefined || raw === null || raw === '') return;
    let disp = raw;
    if (typeof raw === 'string' && (key.includes('Date') || key.toLowerCase().includes('answered') || key.toLowerCase().includes('ordered'))) {
      const parsed = DateHandler.parse(raw);
      disp = parsed ? DateHandler.format(parsed) : raw;
    }
    html += `<tr><th>${escapeHtml(key)}</th><td>${escapeHtml(String(disp))}</td></tr>`;
  });

  html += '</tbody></table>';
  return html;
}

/*************************
 * Export invalid entries (kept)
 *************************/
function exportInvalidEntries(results) {
  const invalidEntries = (results || []).filter(r => r && r.finalStatus === 'invalid');
  if (!invalidEntries.length) { alert('No invalid entries to export.'); return; }
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
    'Remarks': (entry.remarks || []).join('; ')
  }));
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(exportData);
  XLSX.utils.book_append_sheet(wb, ws, 'Invalid Claims');
  XLSX.writeFile(wb, `invalid_claims_${new Date().toISOString().slice(0,10)}.xlsx`);
}

/*************************
 * Handlers & initialization (kept, adjusted to use fixed functions)
 *************************/
async function handleFileUpload(event, type) {
  const file = event.target.files && event.target.files[0];
  if (!file) return;
  try {
    updateStatus(`Loading ${type} file...`);
    if (type === 'eligibility') {
      const reader = new FileReader();
      reader.onload = function(ev) {
        try {
          const data = new Uint8Array(ev.target.result);
          const wb = XLSX.read(data, { type: 'array' });
          const sheet = wb.Sheets[wb.SheetNames[0]];
          // IMPORTANT: read as array-of-arrays (header: 1) so prepareEligibilityMap can detect header row
          const allRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
          eligData = allRows; // now an array-of-arrays (raw rows)
          updateStatus(`Loaded ${Array.isArray(eligData) ? eligData.length : 0} eligibility rows (raw)`);
          updateProcessButtonState();
          // If report already present, summarize counts
          if (eligData && (rawParsedReport || xlsData)) summarizeAndDisplayCounts();
        } catch (err) {
          console.error('Eligibility read error', err);
          updateStatus('Error loading eligibility file');
        }
      };
      reader.onerror = () => updateStatus('Error loading eligibility file');
      reader.readAsArrayBuffer(file);
      return;
    }
    if (type === 'report') {
      lastReportWasCSV = file.name.toLowerCase().endsWith('.csv');
      const parsed = await (file.name.toLowerCase().endsWith('.csv') ? parseCsvFile(file) : parseExcelFile(file));
      rawParsedReport = parsed;
      const normalized = normalizeReportData(parsed);
      xlsData = normalized.filter(r => r && r.claimID && String(r.claimID).trim() !== '');
      if (!xlsData || xlsData.length === 0) {
        console.warn('Report file contained no recognizable claim rows');
      }
      updateStatus(`Loaded ${xlsData.length} report rows`);
      updateProcessButtonState();
      // If eligibility file already present, summarize counts
      if (eligData && (rawParsedReport || xlsData)) summarizeAndDisplayCounts();
      return;
    }
  } catch (err) {
    console.error('File load error:', err);
    updateStatus(`Error loading ${type} file`);
  }
}

async function handlePasteCsvClick() {
  if (!pasteTextarea) return alert('Paste area not found');
  const text = pasteTextarea.value;
  if (!text || !text.trim()) return alert('Please paste CSV text before clicking Load');
  try {
    updateStatus('Parsing pasted CSV...');
    const parsed = await parseCsvText(text);
    lastReportWasCSV = true;
    rawParsedReport = parsed;
    const normalized = normalizeReportData(parsed);
    xlsData = normalized.filter(r => r && r.claimID && String(r.claimID).trim() !== '');
    updateStatus(`Loaded ${xlsData.length} rows from pasted CSV`);
    updateProcessButtonState();
    if (eligData && (rawParsedReport || xlsData)) summarizeAndDisplayCounts();
  } catch (err) {
    console.error('Error parsing pasted CSV:', err);
    updateStatus('Error parsing pasted CSV');
    alert('Failed to parse pasted CSV');
  }
}

function updateProcessButtonState() {
  const hasEligibility = Array.isArray(eligData) && eligData.length > 0;
  const hasReport = Array.isArray(xlsData) && xlsData.length > 0;
  if (processBtn) processBtn.disabled = !(hasEligibility && hasReport);
  if (exportInvalidBtn) exportInvalidBtn.disabled = !(hasEligibility && hasReport);
}

async function handleProcessClick() {
  try {
    if (!eligData) {
      updateStatus('Processing stopped: Eligibility file missing');
      return;
    }
    if (!xlsData || !xlsData.length) {
      updateStatus('Processing stopped: Report file missing');
      return;
    }

    updateStatus('Processing...');

    usedEligibilities.clear();
    const eligMap = prepareEligibilityMap(eligData);
    lastEligMap = eligMap; // save for UI filtering

    // Detect report type automatically
    let reportType = 'Clinicpro'; // default
    const firstRow = xlsData[0];
    if (firstRow) {
      if ('Pri. Claim No' in firstRow) reportType = 'Insta';
      else if ('Pri. Claim ID' in firstRow) reportType = 'Odoo';
    }

    // Validate claims
    const results = validateReportClaims(xlsData, eligMap, reportType);

    // Apply optional filter if checkbox is checked (by insurance company only)
    let outputResults = results;
    if (filterCheckbox && filterCheckbox.checked) {
      outputResults = results.filter(r => {
        const insurance = (r.insuranceCompany || r.provider || r.packageName || '').toString().toLowerCase();
        return insurance.includes('daman') || insurance.includes('thiqa');
      });
    }

    window.lastValidationResults = outputResults;

    // Determine what to display based on the "Invalid only" checkbox (enabled by default)
    const displayedResults = getDisplayedResultsFromStored(outputResults);

    // Render results
    renderResults(displayedResults, eligMap);

    updateStatus(`Processed ${outputResults.length} claims successfully`);

  } catch (err) {
    console.error('Processing stopped due to error:', err);
  }
}

function updateStatus(msg) { if (statusEl) statusEl.textContent = msg || 'Ready'; }

function onFilterToggle() {
  if (!filterStatus) return;
  const on = filterCheckbox && filterCheckbox.checked;
  filterStatus.textContent = on ? 'ON' : 'OFF';
  filterStatus.classList.toggle('active', on);

  if (!window.lastValidationResults) return;

  // Start from the processed results, then apply the Daman/Thiqa filter (if any), then the Invalid-only filter
  let base = window.lastValidationResults.slice();

  if (on) {
    base = base.filter(r => {
      const provider = (r.provider || r.insuranceCompany || r.packageName || r['Payer Name'] || r['Insurance Company'] || '').toString().toLowerCase();
      return provider.includes('daman') || provider.includes('thiqa');
    });
  }

  // Apply invalid-only display filter if active
  const displayed = getDisplayedResultsFromStored(base);

  // Use the last eligMap if available
  const eligMap = lastEligMap || (eligData ? prepareEligibilityMap(eligData) : new Map());
  renderResults(displayed, eligMap);
}

function initializeEventListeners() {
  reportInput = document.getElementById('reportFileInput');
  eligInput = document.getElementById('eligibilityFileInput');
  processBtn = document.getElementById('processBtn');
  exportInvalidBtn = document.getElementById('exportInvalidBtn');
  statusEl = document.getElementById('uploadStatus');
  resultsContainer = document.getElementById('results');
  filterCheckbox = document.getElementById('filterDamanThiqa');
  filterStatus = document.getElementById('filterStatus');
  pasteTextarea = document.getElementById('pasteCsvTextarea');
  pasteBtn = document.getElementById('pasteCsvBtn');
  invalidOnlyCheckbox = document.getElementById('filterInvalidOnly');

  if (eligInput) eligInput.addEventListener('change', (e) => handleFileUpload(e, 'eligibility'));
  if (reportInput) reportInput.addEventListener('change', (e) => handleFileUpload(e, 'report'));
  if (processBtn) processBtn.addEventListener('click', handleProcessClick);
  if (exportInvalidBtn) exportInvalidBtn.addEventListener('click', () => exportInvalidEntries(window.lastValidationResults || []));
  if (filterCheckbox) filterCheckbox.addEventListener('change', onFilterToggle);

  // invalid-only checkbox: enabled by default, re-render when toggled
  if (invalidOnlyCheckbox) {
    invalidOnlyCheckbox.checked = true;
    invalidOnlyCheckbox.addEventListener('change', () => {
      if (!window.lastValidationResults) return;
      const base = window.lastValidationResults.slice();
      // If Daman/Thiqa filter is on, respect it too
      let preFiltered = base;
      if (filterCheckbox && filterCheckbox.checked) {
        preFiltered = base.filter(r => {
          const provider = (r.provider || r.insuranceCompany || r.packageName || '').toString().toLowerCase();
          return provider.includes('daman') || provider.includes('thiqa');
        });
      }
      const displayed = getDisplayedResultsFromStored(preFiltered);
      const eligMap = lastEligMap || (eligData ? prepareEligibilityMap(eligData) : new Map());
      renderResults(displayed, eligMap);
    });
  }

  if (pasteBtn) pasteBtn.addEventListener('click', handlePasteCsvClick);

  if (filterStatus) onFilterToggle();
}

document.addEventListener('DOMContentLoaded', () => {
  initializeEventListeners();
  updateStatus('Ready to process files');
});
