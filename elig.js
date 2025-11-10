/*******************************
 * elig-fixed.js
 *
 * Report-only script: fixed core functions (from working script)
 * - retains report-only UI/behavior from second script
 * - uses robust DateHandler, normalization, and eligibility logic from the first (working) script
 * - no XML mode included
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
let eligData = null;       // eligibility sheet as array of objects
let rawParsedReport = null; // raw parsed sheet result (header detection output)
const usedEligibilities = new Set();
let lastReportWasCSV = false;

// DOM Elements (lookups performed in initializeEventListeners)
let reportInput, eligInput, processBtn, exportInvalidBtn, statusEl, resultsContainer, filterCheckbox, filterStatus, pasteTextarea, pasteBtn;

/*************************
 * DATE HANDLING UTILITIES (From first script)
 *************************/
const DateHandler = {
  parse: function(input, options = {}) {
    const preferMDY = !!options.preferMDY;
    if (!input && input !== 0) return null;
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
    // Return UTC midnight
    return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
  },

  _parseStringDate: function(dateStr, preferMDY = false) {
    if (!dateStr || typeof dateStr !== 'string') return null;
    if (dateStr.includes(' ')) {
      dateStr = dateStr.split(' ')[0];
    }
    // Matches DD/MM/YYYY or MM/DD/YYYY (ambiguous). We'll disambiguate using preferMDY flag
    const dmyMdyMatch = dateStr.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
    if (dmyMdyMatch) {
      const part1 = parseInt(dmyMdyMatch[1], 10);
      const part2 = parseInt(dmyMdyMatch[2], 10);
      let year = parseInt(dmyMdyMatch[3], 10);
      if (year < 100) year += 2000; // handle 2-digit years robustly

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
      const day = parseInt(textMatch[1], 10);
      let year = parseInt(textMatch[3], 10);
      if (year < 100) year += 2000;
      const monthIndex = MONTHS.indexOf(textMatch[2].toLowerCase().substr(0, 3));
      if (monthIndex >= 0) return new Date(Date.UTC(year, monthIndex, day));
    }

    // ISO: 2025-07-01
    const isoMatch = dateStr.match(/^(\d{4})[\/\-](\d{2})[\/\-](\d{2})$/);
    if (isoMatch) return new Date(Date.UTC(parseInt(isoMatch[1],10), parseInt(isoMatch[2],10) - 1, parseInt(isoMatch[3],10)));
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
    if (score > bestScore) { bestScore = score; bestIndex = i; }
  }

  const headerRowIndex = bestScore > 0 ? bestIndex : 0;
  const rawHeaderRow = allRows[headerRowIndex] || [];
  const headers = rawHeaderRow.map(h => (h === null || h === undefined) ? '' : String(h).trim());
  const dataRows = allRows.slice(headerRowIndex + 1);
  const rows = dataRows.map(rowArray => {
    const obj = {};
    for (let c = 0; c < headers.length; c++) {
      const key = headers[c] || `Column${c+1}`;
      obj[key] = rowArray[c] === undefined || rowArray[c] === null ? '' : rowArray[c];
    }
    return obj;
  });
  return { headerRowIndex, headers, rows };
}

/*******************************
 * ELIGIBILITY MATCHING FUNCTIONS (from first)
 *******************************/
function prepareEligibilityMap(eligArray) {
    if (!eligArray || eligArray.length === 0) return new Map();
    const headers = eligArray[0];
    const eligMap = new Map();
    eligArray.forEach((record, index) => {
        if (index === 0) return;
        const normalizedRecord = {};
        for (const key in headers) {
            const headerName = headers[key].trim();
            if (headerName) normalizedRecord[headerName] = record[key] || '';
        }
        const rawMemberID = String(normalizedRecord['memberID'] || normalizedRecord['Policy1'] || normalizedRecord['Card Number / DHA Member ID'] || '').trim();
        if (!rawMemberID) return;
        const memberID = normalizeMemberID(rawMemberID);
        if (!memberID) return;
        eligMap.set(memberID, normalizedRecord);
    });
    console.log(eligMap);
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

function findEligibilityForClaim(eligMap, claimDate, memberID, claimClinicians = []) {
  const normalizedID = normalizeMemberID(memberID || '');
  const eligList = eligMap.get(normalizedID) || [];
  if (!eligList.length) return null;

  for (const elig of eligList) {
    const eligDate = DateHandler.parse(elig["Answered On"]);
    if (!eligDate) continue;
    if (!DateHandler.isSameDay(claimDate, eligDate)) continue;

    const eligClinician = (elig.Clinician || '').trim();
    const serviceCategory = (elig['Service Category'] || '').trim();
    const consultationStatus = (elig['Consultation Status'] || '').trim();
    const department = (elig.Department || elig.Clinic || '').toLowerCase();
    const categoryCheck = isServiceCategoryValid(serviceCategory, consultationStatus, department);

    if (!categoryCheck.valid) continue;
    if ((elig.Status || '').toLowerCase() !== 'eligible') continue;

    return elig;
  }
  return null;
}

/*************************
 * Normalize & validate report rows (use working version's normalizeReportData)
 *************************/
function normalizeReportData(rawData) {
  // rawData expected shape: { headers: [...], rows: [...] } (from findHeaderRowFromArrays)
  // But keep compatibility if rawData is an array (older shape)
  if (!rawData) return [];

  // Backwards compatibility: if function received array-of-objects (already parsed sheet_to_json)
  if (Array.isArray(rawData) && rawData.length > 0 && typeof rawData[0] === 'object' && !rawData.headers) {
    // trying to detect known formats from object keys
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

  // Otherwise assume rawData has headers + rows from detection function
  const rows = rawData.rows || [];
  const headers = rawData.headers || [];

  function getField(obj, candidates) {
    for (const k of candidates) {
      if (obj && Object.prototype.hasOwnProperty.call(obj, k) && obj[k] !== '' && obj[k] !== null && obj[k] !== undefined) return obj[k];
    }
    return '';
  }

  return rows.map(r => {
    // Use the same mapping logic as the working script's normalizeReportData
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
      // generic mapping with fallback header scanning
      const out = {
        claimID: r['ClaimID'] || r['Pri. Claim No'] || r['Pri. Claim ID'] || getField(r, ['ClaimID','Pri. Claim No','Pri. Claim ID','Claim ID','Pri. Claim ID']) || '',
        memberID: r['Pri. Member ID'] || r['Pri. Patient Insurance Card No'] || r['PatientCardID'] || getField(r, ['PatientCardID','Patient Insurance Card No','Card Number / DHA Member ID','Card Number','MemberID','Member ID']) || '',
        claimDate: r['Encounter Date'] || r['Adm/Reg. Date'] || r['ClaimDate'] || getField(r, ['Encounter Date','ClaimDate','Adm/Reg. Date','Date']) || '',
        clinician: r['Clinician License'] || r['Admitting License'] || r['OrderDoctor'] || getField(r, ['Clinician License','Clinician','Admitting License','OrderDoctor']) || '',
        department: r['Department'] || r['Clinic'] || r['Admitting Department'] || getField(r, ['Department','Clinic','Admitting Department']) || '',
        packageName: r['Pri. Payer Name'] || r['Insurance Company'] || r['Pri. Sponsor'] || getField(r, ['Pri. Payer Name','Insurance Company','Pri. Plan Type','Package','Pri. Sponsor']) || '',
        insuranceCompany: r['Pri. Payer Name'] || r['Insurance Company'] || getField(r, ['Payer Name','Insurance Company','Pri. Payer Name']) || '',
        claimStatus: r['Codification Status'] || r['VisitStatus'] || r['Status'] || getField(r, ['Codification Status','VisitStatus','Status','Claim Status']) || ''
      };

      // If still missing memberID / claimID — scan headers for likely column
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
function parseExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = function(e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const allRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
        const detection = findHeaderRowFromArrays(allRows, 20);
        resolve(detection);
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

/*************************
 * UI RENDERING FUNCTIONS (kept from second)
 *************************/
function escapeHtml(s) {
  if (s === null || s === undefined) return '';
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#x27;');
}

/*************************
 * summarizeAndDisplayCounts (kept)
 *************************/
function summarizeAndDisplayCounts() {
  try {
    const eligCount = Array.isArray(eligData) ? eligData.length : 0;
    // Ensure xlsData exists, if not but rawParsedReport exists try to normalize it now
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
 * renderResults and modal functions (kept from second)
 *************************/
function renderResults(results, eligMap) {
  if (!resultsContainer) return;
  resultsContainer.innerHTML = '';
  if (!results || results.length === 0) {
    resultsContainer.innerHTML = '<div class="no-results">No claims to display</div>';
    return;
  }
  const tableContainer = document.createElement('div');
  tableContainer.className = 'analysis-results';
  tableContainer.style.overflowX = 'auto';
  const table = document.createElement('table');
  table.className = 'shared-table';
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
  results.forEach((result, index) => {
    if (!result.memberID || result.memberID.trim() === '') return;
    const statusToCheck = (result.claimStatus || result.status || result.fullEligibilityRecord?.Status || '')
      .toString()
      .trim()
      .toLowerCase();
    if (statusToCheck === 'not seen') return;
    if (result.finalStatus && statusCounts.hasOwnProperty(result.finalStatus)) {
      statusCounts[result.finalStatus]++;
    }
    const row = document.createElement('tr');
    row.className = result.finalStatus;
    const statusBadge = result.status ? `<span class="status-badge ${result.status.toLowerCase() === 'eligible' ? 'eligible' : 'ineligible'}">${escapeHtml(result.status)}</span>` : '';
    const remarksHTML = result.remarks && result.remarks.length > 0 ? result.remarks.map(r => `<div>${escapeHtml(r)}</div>`).join('') : '<div class="source-note">No remarks</div>';
    let detailsCell = '<div class="source-note">N/A</div>';
    if (result.fullEligibilityRecord?.['Eligibility Request Number']) {
      detailsCell = `<button class="details-btn eligibility-details" data-index="${index}">${escapeHtml(result.fullEligibilityRecord['Eligibility Request Number'])}</button>`;
    } else if (eligMap && typeof eligMap.get === 'function' && (eligMap.get(result.memberID) || []).length) {
      detailsCell = `<button class="details-btn show-all-eligibilities" data-member="${escapeHtml(result.memberID)}">View All</button>`;
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
      <td>${detailsCell}</td>
    `;
    tbody.appendChild(row);
  });
  table.appendChild(tbody);
  tableContainer.appendChild(table);
  resultsContainer.appendChild(tableContainer);
  const summary = document.createElement('div');
  summary.className = 'loaded-count';
  summary.innerHTML = `
    Processed ${results.length} claims: 
    <span class="valid">${statusCounts.valid} valid</span>, 
    <span class="unknown">${statusCounts.unknown} unknown</span>, 
    <span class="invalid">${statusCounts.invalid} invalid</span>
  `;
  resultsContainer.prepend(summary);
  initEligibilityModal(results, eligMap);
}

function initEligibilityModal(results, eligMap) {
  if (!document.getElementById("modalOverlay")) {
    const modalHtml = `
      <div id="modalOverlay" style="display:none;position:fixed;z-index:9999;left:0;top:0;width:100vw;height:100vh;background:rgba(0,0,0,0.35);">
        <div id="modalContent" style="
          background:#fff;
          width:90%;
          max-width:1200px;
          max-height:90vh;
          overflow:auto;
          position:absolute;
          left:50%;
          top:50%;
          transform:translate(-50%,-50%);
          padding:20px;
          border-radius:8px;
          box-shadow:0 4px 24px rgba(0,0,0,0.2);
        ">
          <button id="modalCloseBtn" style="
            float:right;
            font-size:18px;
            padding:2px 10px;
            cursor:pointer;
          " aria-label="Close">&times;</button>
          <div id="modalTable"></div>
        </div>
      </div>
    `;
    document.body.insertAdjacentHTML("beforeend", modalHtml);
    document.getElementById("modalCloseBtn").onclick = hideModal;
    document.getElementById("modalOverlay").onclick = function(e) { if (e.target.id === "modalOverlay") hideModal(); };
  }
  document.querySelectorAll(".eligibility-details").forEach(btn => {
    btn.onclick = function() {
      const index = parseInt(this.dataset.index, 10);
      const result = results[index];
      if (!result?.fullEligibilityRecord) return;
      const record = result.fullEligibilityRecord;
      document.getElementById("modalTable").innerHTML = formatEligibilityDetails(record, result.memberID);
      document.getElementById("modalOverlay").style.display = "block";
    };
  });
  document.querySelectorAll(".show-all-eligibilities").forEach(btn => {
    btn.onclick = function() {
      const member = this.dataset.member;
      const list = (eligMap.get(member) || []);
      if (!list.length) {
        document.getElementById("modalTable").innerHTML = `<div>No eligibilities found for ${escapeHtml(member)}</div>`;
        document.getElementById("modalOverlay").style.display = "block";
        return;
      }
      let html = `<h3>Eligibilities for ${escapeHtml(member)}</h3><div style="overflow-x:auto;"><table style="width:100%;border-collapse:collapse;"><thead><tr><th>#</th><th>Request No</th><th>Answered On</th><th>Status</th><th>Clinician</th><th>Service Category</th><th>Package</th></tr></thead><tbody>`;
      list.forEach((rec, idx) => {
        html += `<tr>
          <td style="padding:6px;border-bottom:1px solid #eee">${idx+1}</td>
          <td style="padding:6px;border-bottom:1px solid #eee">${escapeHtml(rec['Eligibility Request Number']||'')}</td>
          <td style="padding:6px;border-bottom:1px solid #eee">${escapeHtml(rec['Answered On']||rec['Ordered On']||'')}</td>
          <td style="padding:6px;border-bottom:1px solid #eee">${escapeHtml(rec['Status']||'')}</td>
          <td style="padding:6px;border-bottom:1px solid #eee">${escapeHtml(rec['Clinician']||'')}</td>
          <td style="padding:6px;border-bottom:1px solid #eee">${escapeHtml(rec['Service Category']||'')}</td>
          <td style="padding:6px;border-bottom:1px solid #eee">${escapeHtml(rec['Package Name']||'')}</td>
        </tr>`;
      });
      html += `</tbody></table></div>`;
      document.getElementById("modalTable").innerHTML = html;
      document.getElementById("modalOverlay").style.display = "block";
    };
  });
}

function hideModal() { const overlay = document.getElementById("modalOverlay"); if (overlay) overlay.style.display = "none"; }

function formatEligibilityDetails(record, memberID) {
  if (!record) return '<div>No details</div>';
  let html = `<div style="margin-bottom:8px;"><strong>Member:</strong> ${escapeHtml(memberID)} <span style="margin-left:8px;" class="status-badge ${((record.Status||'').toLowerCase()==='eligible')?'eligible':'ineligible'}">${escapeHtml(record.Status||'')}</span></div>`;
  html += '<table style="width:100%;border-collapse:collapse;"><tbody>';
  Object.entries(record).forEach(([k,v]) => {
    if ((v === null || v === undefined || v === '') && v !== 0) return;
    let disp = v;
    if (DATE_KEYS.some(dk => k.includes(dk)) || k.toLowerCase().includes('answered') || k.toLowerCase().includes('ordered')) {
      const p = DateHandler.parse(v);
      disp = p ? DateHandler.format(p) : v;
    }
    html += `<tr><th style="text-align:left;padding:6px;border-bottom:1px solid #eee;width:30%">${escapeHtml(k)}</th><td style="padding:6px;border-bottom:1px solid #eee">${escapeHtml(disp)}</td></tr>`;
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
          const json = XLSX.utils.sheet_to_json(sheet, { defval: '' });
          eligData = json;
          updateStatus(`Loaded ${Array.isArray(eligData) ? eligData.length : 0} eligibility records`);
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
        const insurance = (r.insuranceCompany || '').toString().toLowerCase();
        return insurance.includes('daman') || insurance.includes('thiqa');
      });
    }

    window.lastValidationResults = outputResults;

    // Render results
    renderResults(outputResults, eligMap);

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
  if (window.lastValidationResults) {
    const eligMap = eligData ? prepareEligibilityMap(eligData) : new Map();
    // Re-render with filtering honored by renderResults wrapper in handleProcessClick
    if (on) {
      const filtered = window.lastValidationResults.filter(r => {
        const provider = (r.provider || r.insuranceCompany || r.packageName || r['Payer Name'] || r['Insurance Company'] || '').toString().toLowerCase();
        return provider.includes('daman') || provider.includes('thiqa');
      });
      renderResults(filtered, eligMap);
    } else {
      renderResults(window.lastValidationResults, eligMap);
    }
  }
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

  if (eligInput) eligInput.addEventListener('change', (e) => handleFileUpload(e, 'eligibility'));
  if (reportInput) reportInput.addEventListener('change', (e) => handleFileUpload(e, 'report'));
  if (processBtn) processBtn.addEventListener('click', handleProcessClick);
  if (exportInvalidBtn) exportInvalidBtn.addEventListener('click', () => exportInvalidEntries(window.lastValidationResults || []));
  if (filterCheckbox) filterCheckbox.addEventListener('change', onFilterToggle);
  if (pasteBtn) pasteBtn.addEventListener('click', handlePasteCsvClick);

  if (filterStatus) onFilterToggle();
}

document.addEventListener('DOMContentLoaded', () => {
  initializeEventListeners();
  updateStatus('Ready to process files');
});
