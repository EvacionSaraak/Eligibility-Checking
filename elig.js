/*******************************
 * elig.js - robust parser & matcher
 * Replaces and improves previous file to handle xls/xlsx/csv + pasted CSV + header mapping
 *******************************/

const SERVICE_PACKAGE_RULES = {
  'Dental Services': ['dental', 'orthodontic'],
  'Physiotherapy': ['physio'],
  'Other OP Services': ['physio', 'diet', 'occupational', 'speech'],
  'Consultation': []
};
const DATE_KEYS = ['Date', 'On'];
const MONTHS = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"];

// Normalized rules for case-insensitive lookup
const NORMALIZED_SERVICE_PACKAGE_RULES = {};
Object.keys(SERVICE_PACKAGE_RULES).forEach(k => {
  NORMALIZED_SERVICE_PACKAGE_RULES[k.trim().toLowerCase()] = SERVICE_PACKAGE_RULES[k];
});

// App state
let eligData = null;
let xlsData = null;
let lastReportWasCSV = false;
const usedEligibilities = new Set();

// DOM references (initialized later)
let reportInput = null;
let eligInput = null;
let processBtn = null;
let exportInvalidBtn = null;
let statusEl = null;
let resultsContainer = null;
let filterCheckbox = null;
let filterStatus = null;
let pasteTextarea = null;
let pasteBtn = null;

/*************************
 * Utilities & Date funcs
 *************************/
function escapeHtml(s) {
  if (s === null || s === undefined) return '';
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#x27;');
}

function normalizeMemberID(id) {
  if (id === null || id === undefined) return '';
  return String(id).trim().replace(/^0+/, '');
}
function normalizeClinician(name) {
  if (!name) return '';
  return String(name).trim().toLowerCase().replace(/\s+/g, ' ');
}

const DateHandler = {
  parse(input, options = {}) {
    const preferMDY = !!options.preferMDY;
    if (!input && input !== 0) return null;
    if (input instanceof Date) return isNaN(input) ? null : input;
    if (typeof input === 'number') return this._parseExcelDate(input);
    const cleanStr = String(input).trim().replace(/[,.]/g, '');
    const parsed = this._parseStringDate(cleanStr, preferMDY) || new Date(cleanStr);
    if (isNaN(parsed)) return null;
    return parsed;
  },
  format(date) {
    if (!(date instanceof Date) || isNaN(date)) return '';
    const d = date.getUTCDate().toString().padStart(2, '0');
    const m = (date.getUTCMonth() + 1).toString().padStart(2, '0');
    const y = date.getUTCFullYear();
    return `${d}/${m}/${y}`;
  },
  isSameDay(a, b) {
    if (!a || !b) return false;
    return a.getUTCFullYear() === b.getUTCFullYear() &&
           a.getUTCMonth() === b.getUTCMonth() &&
           a.getUTCDate() === b.getUTCDate();
  },
  _parseExcelDate(serial) {
    const utcDays = Math.floor(serial) - 25569;
    const ms = utcDays * 86400 * 1000;
    const d = new Date(ms);
    return new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
  },
  _parseStringDate(dateStr, preferMDY = false) {
    if (!dateStr || typeof dateStr !== 'string') return null;
    if (dateStr.includes(' ')) dateStr = dateStr.split(' ')[0];
    // DD/MM/YYYY or MM/DD/YYYY
    const dmyMdy = dateStr.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
    if (dmyMdy) {
      const p1 = parseInt(dmyMdy[1], 10);
      const p2 = parseInt(dmyMdy[2], 10);
      let y = parseInt(dmyMdy[3], 10);
      if (y < 100) y += 2000;
      if (p1 > 12 && p2 <= 12) return new Date(Date.UTC(y, p2 - 1, p1));
      if (p2 > 12 && p1 <= 12) return new Date(Date.UTC(y, p1 - 1, p2));
      return preferMDY ? new Date(Date.UTC(y, p1 - 1, p2)) : new Date(Date.UTC(y, p2 - 1, p1));
    }
    // 30-Jun-2025 or 30 Jun 25
    const textual = dateStr.match(/^(\d{1,2})[\/\- ]([a-z]{3,})[\/\- ](\d{2,4})$/i);
    if (textual) {
      const day = parseInt(textual[1], 10);
      let year = parseInt(textual[3], 10);
      if (year < 100) year += 2000;
      const mon = MONTHS.indexOf(textual[2].toLowerCase().substr(0,3));
      if (mon >= 0) return new Date(Date.UTC(year, mon, day));
    }
    // ISO
    const iso = dateStr.match(/^(\d{4})[\/\-](\d{2})[\/\-](\d{2})$/);
    if (iso) {
      const y = parseInt(iso[1], 10);
      const mo = parseInt(iso[2], 10);
      const d = parseInt(iso[3], 10);
      return new Date(Date.UTC(y, mo - 1, d));
    }
    return null;
  }
};

/*****************************
 * Header mapping utilities
 *****************************/
// Map many possible header names to canonical keys used in the app.
const HEADER_SYNONYMS = {
  claimID: [
    /^pri\.?\s*claim\s*no$/i, /^pri\.?\s*claim\s*id$/i, /^claimid$/i, /^claim\s*id$/i,
    /^claim\s*number$/i, /^pri\.?\s*claim$/i, /^id$/i
  ],
  memberID: [
    /^card\s*number\s*\/\s*dha\s*member\s*id$/i, /^card\s*number$/i, /^patientcardid$/i,
    /^patient\s*insurance\s*card\s*no$/i, /^patient\s*insurance\s*card\s*id$/i,
    /^memberid$/i, /^member\s*id$/i, /^patientid$/i
  ],
  claimDate: [
    /^encounter\s*date$/i, /^claim\s*date$/i, /^adm\/reg\.\s*date$/i, /^visit\s*date$/i,
    /^date$/i, /^encounterdate$/i
  ],
  clinician: [
    /^clinician\s*license$/i, /^clinician$/i, /^admitting\s*license$/i, /^provider$/i, /^doctor$/i
  ],
  department: [
    /^department$/i, /^clinic$/i, /^admitting\s*department$/i, /^service\s*dept$/i, /^dept$/i
  ],
  packageName: [
    /^pri\.\s*payer\s*name$/i, /^pri\.\s*sponsor$/i, /^insurance\s*company$/i, /^package$/i,
    /^pri\.\s*plan\s*type$/i, /^pri\.\s*payer$/i
  ],
  insuranceCompany: [
    /^pri\.\s*payer\s*name$/i, /^insurance\s*company$/i, /^payer$/i, /^pri\.\s*plan\s*type$/i
  ],
  claimStatus: [
    /^codification\s*status$/i, /^visitstatus$/i, /^status$/i, /^claim\s*status$/i
  ]
};

// Normalize header string: trim, collapse spaces, lowercase
function normalizeHeaderKey(h) {
  if (h === null || h === undefined) return '';
  return String(h).trim().replace(/\s+/g,' ').toLowerCase();
}

// Given an array-of-arrays from sheet_to_json(header:1), detect header row index heuristically
function detectHeaderRow(allRows, maxScan = 15) {
  const rows = Array.isArray(allRows) ? allRows : [];
  for (let i = 0; i < Math.min(maxScan, rows.length); i++) {
    const row = rows[i] || [];
    const nonEmpty = row.filter(c => String(c).trim() !== '').length;
    if (nonEmpty >= 3) {
      // prefer row with known tokens
      const joined = row.join(' ').toLowerCase();
      if (joined.includes('pri') || joined.includes('claim') || joined.includes('card') || joined.includes('member') || joined.includes('patient')) {
        return i;
      }
    }
  }
  // fallback to first non-empty row
  for (let i=0;i<Math.min(maxScan, rows.length);i++){
    const row = rows[i] || [];
    const nonEmpty = row.filter(c => String(c).trim() !== '').length;
    if (nonEmpty>0) return i;
  }
  return 0;
}

// Map raw headers (from sheet) to canonical keys
function mapHeadersToCanonical(rawHeaders) {
  const mapped = {};
  rawHeaders.forEach((raw, idx) => {
    const key = normalizeHeaderKey(raw);
    let found = null;
    for (const [canon, patterns] of Object.entries(HEADER_SYNONYMS)) {
      for (const p of patterns) {
        if (p.test(key)) { found = canon; break; }
      }
      if (found) break;
    }
    // fallback heuristics
    if (!found) {
      if (key.includes('claim') && key.includes('id')) found = 'claimID';
      else if ((key.includes('card') && key.includes('no')) || key.includes('member') || key.includes('patientcard')) found = 'memberID';
      else if (key.includes('date')) found = 'claimDate';
      else if (key.includes('clin') || key.includes('doctor') || key.includes('provider')) found = 'clinician';
      else if (key.includes('clinic') || key.includes('department') || key.includes('dept')) found = 'department';
      else if (key.includes('payer') || key.includes('plan') || key.includes('insurance')) found = 'insuranceCompany';
    }
    mapped[idx] = found || null;
  });
  return mapped; // index -> canonical key (or null)
}

// Convert a row array to normalized object using header mapping
function rowArrayToNormalizedObject(rowArray, headerMap, rawHeaders) {
  const obj = {};
  for (let i=0;i<headerMap.length;i++) {
    const canon = headerMap[i];
    const rawHeader = rawHeaders[i] || (`Column${i+1}`);
    const rawVal = rowArray[i] === undefined || rowArray[i] === null ? '' : rowArray[i];
    if (canon) {
      obj[canon] = rawVal;
    } else {
      // also retain raw by header name for fallback
      obj[rawHeader] = rawVal;
    }
  }
  return obj;
}

/*****************************
 * Parsing: Excel (.xls/.xlsx) and CSV (.csv) & pasted CSV
 *****************************/
function parseExcelArrayBuffer(arrayBuffer) {
  const wb = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const allRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  return parseSheetRows(allRows);
}

function parseCsvTextString(text) {
  const wb = XLSX.read(text, { type: 'string' });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const allRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  return parseSheetRows(allRows);
}

// Core: takes array-of-arrays and returns array of normalized rows (claimID, memberID, etc.)
function parseSheetRows(allRows) {
  if (!Array.isArray(allRows) || allRows.length === 0) return [];

  const headerRowIndex = detectHeaderRow(allRows, 20);
  const rawHeaderRow = (allRows[headerRowIndex] || []).map(h => String(h).trim());
  const headerMap = []; // array of canonical keys in same order as rawHeaderRow
  const mapped = mapHeadersToCanonical(rawHeaderRow);
  for (let i=0;i<rawHeaderRow.length;i++) {
    headerMap[i] = mapped[i] || null;
  }

  // Build normalized objects for each subsequent row
  const dataRows = allRows.slice(headerRowIndex + 1);
  const normalizedRows = dataRows.map(rowArr => {
    const normObj = rowArrayToNormalizedObject(rowArr, headerMap, rawHeaderRow);
    // Fallback: if canonical keys missing, try to grab by raw header names typical for ClinicPro/Odoo/Insta
    // e.g. 'ClaimID' 'PatientCardID' etc
    // The rowObj may already include keys for raw headers, handled in normalizeReportData below
    return normObj;
  });

  // Return rows and raw headers to let downstream mapping normalize fields
  return {
    rows: normalizedRows,
    rawHeaders: rawHeaderRow
  };
}

/*****************************
 * File read helpers (FileReader) that call the parse functions
 *****************************/
function parseExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const parsed = parseExcelArrayBuffer(e.target.result);
        resolve(parsed);
      } catch (err) { reject(err); }
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(file);
  });
}

function parseCsvFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const parsed = parseCsvTextString(e.target.result);
        resolve(parsed);
      } catch (err) { reject(err); }
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsText(file);
  });
}

// parse pasted CSV text
function parseCsvText(text) {
  return new Promise((resolve, reject) => {
    try {
      const parsed = parseCsvTextString(text);
      resolve(parsed);
    } catch (err) { reject(err); }
  });
}

/*****************************
 * Normalize parsed rows to canonical report rows
 *****************************/
function normalizeParsedSheet(parsed) {
  // parsed: { rows: [ {claimID:..., memberID:...} OR raw header keys... ], rawHeaders: [...] }
  const rows = parsed.rows || [];
  const rawHeaders = parsed.rawHeaders || [];

  // For rows that already have canonical keys, return mapped object
  const normalized = rows.map(r => {
    // If r already had canonical keys, use them
    const out = {
      claimID: r.claimID || r['ClaimID'] || r['Pri. Claim No'] || r['Pri. Claim ID'] || '',
      memberID: r.memberID || r['PatientCardID'] || r['Patient Insurance Card No'] || r['Patient Insurance Card No'] || r['PatientInsuranceCardNo'] || r['Card Number / DHA Member ID'] || r['Card Number'] || r['Member ID'] || r['MemberID'] || '',
      claimDate: r.claimDate || r['Encounter Date'] || r['ClaimDate'] || r['Adm/Reg. Date'] || r['EncounterDate'] || '',
      clinician: r.clinician || r['Clinician License'] || r['Clinician'] || r['Admitting License'] || '',
      department: r.department || r['Department'] || r['Clinic'] || r['Admitting Department'] || '',
      packageName: r.packageName || r['Pri. Payer Name'] || r['Insurance Company'] || r['Pri. Plan Type'] || r['Package'] || '',
      insuranceCompany: r.insuranceCompany || r['Payer Name'] || r['Pri. Payer Name'] || r['Insurance Company'] || '',
      claimStatus: r.claimStatus || r['Codification Status'] || r['VisitStatus'] || r['Status'] || ''
    };
    // If claimID is empty, try by scanning raw headers -> raw header name like 'Column1' etc
    if (!out.claimID) {
      for (const h of rawHeaders) {
        const val = r[h];
        if (val && String(val).toLowerCase().includes('claim')) {
          out.claimID = val;
          break;
        }
      }
    }
    return out;
  });

  return normalized;
}

/*****************************
 * Eligibility map & matching
 *****************************/
function prepareEligibilityMap(eligArray) {
  const eligMap = new Map();
  if (!Array.isArray(eligArray)) return eligMap;
  eligArray.forEach(e => {
    const rawID =
      e['Card Number / DHA Member ID'] ||
      e['Card Number'] ||
      e['_5'] ||
      e['MemberID'] ||
      e['Member ID'] ||
      e['Patient Insurance Card No'] ||
      e['PatientCardID'] ||
      e['CardNumber'] ||
      e['Patient Insurance Card No'];

    if (rawID === undefined || rawID === null || rawID === '') return;
    const memberID = normalizeMemberID(rawID);
    if (!eligMap.has(memberID)) eligMap.set(memberID, []);
    eligMap.get(memberID).push({
      'Eligibility Request Number': e['Eligibility Request Number'] || e['Eligibility Request No'] || e['Request Number'] || '',
      'Card Number / DHA Member ID': rawID,
      'Answered On': e['Answered On'] || e['AnsweredOn'] || e['Answered Date'] || '',
      'Ordered On': e['Ordered On'] || e['OrderedOn'] || '',
      'Status': e['Status'] || '',
      'Clinician': e['Clinician'] || '',
      'Payer Name': e['Payer Name'] || e['PayerName'] || '',
      'Service Category': e['Service Category'] || '',
      'Package Name': e['Package Name'] || e['PackageName'] || '',
      'Department': e['Department'] || e['Clinic'] || ''
    });
  });
  return eligMap;
}

function checkClinicianMatch(claimClinicians, eligClinician) {
  if (!eligClinician || !claimClinicians?.length) return true;
  const normElig = normalizeClinician(eligClinician);
  return claimClinicians.some(c => normalizeClinician(c) === normElig);
}

function isServiceCategoryValid(serviceCategory, consultationStatus, rawPackage) {
  if (!serviceCategory) return { valid: true };
  const categoryLower = String(serviceCategory).trim().toLowerCase();
  const pkgRaw = rawPackage || '';
  const pkg = String(pkgRaw).toLowerCase();

  if (categoryLower === 'consultation' && consultationStatus?.toLowerCase() === 'elective') {
    const disallowed = ['dental', 'physio', 'diet', 'occupational', 'speech'];
    if (disallowed.some(term => pkg.includes(term))) {
      return { valid: false, reason: `Consultation (Elective) cannot include restricted service types. Found: "${pkgRaw}"` };
    }
    return { valid: true };
  }

  const allowedKeywords = NORMALIZED_SERVICE_PACKAGE_RULES[categoryLower];
  if (allowedKeywords && allowedKeywords.length > 0) {
    if (pkg && !allowedKeywords.some(keyword => pkg.includes(keyword))) {
      return { valid: false, reason: `${serviceCategory} category requires related package. Found: "${pkgRaw}"` };
    }
  }
  return { valid: true };
}

// Find eligibility record matching the claim (memberID normalized, same Answered On day, clinician match, category rules)
function findEligibilityForClaim(eligMap, claimDate, memberID, claimClinicians = []) {
  const normalizedID = normalizeMemberID(memberID);
  const eligList = eligMap.get(normalizedID) || [];
  if (!eligList.length) return null;

  const claimCliniciansFiltered = Array.isArray(claimClinicians) ? claimClinicians.filter(Boolean) : [];

  for (const elig of eligList) {
    const eligDate = DateHandler.parse(elig['Answered On'] || elig['Ordered On']);
    if (!DateHandler.isSameDay(claimDate, eligDate)) {
      // If no same day, continue; precise rule can be relaxed later if needed
      console.log(`Date mismatch for member ${memberID}: claim ${DateHandler.format(claimDate)} vs elig ${DateHandler.format(eligDate)}`);
      continue;
    }

    if (elig.Clinician && claimCliniciansFiltered.length && !checkClinicianMatch(claimCliniciansFiltered, elig.Clinician)) {
      console.log(`Clinician mismatch for member ${memberID}: claim clinicians ${JSON.stringify(claimCliniciansFiltered)} vs elig ${elig.Clinician}`);
      continue;
    }

    const serviceCategory = elig['Service Category'] || '';
    const consultationStatus = elig['Consultation Status'] || '';
    const dept = (elig.Department || elig.Clinic || '').toLowerCase();
    if (!isServiceCategoryValid(serviceCategory, consultationStatus, dept).valid) {
      console.log(`Service category mismatch for member ${memberID}`);
      continue;
    }

    if ((elig.Status || '').toLowerCase() !== 'eligible') {
      console.log(`Status not eligible for member ${memberID}: ${elig.Status}`);
      continue;
    }

    if (elig['Eligibility Request Number']) usedEligibilities.add(elig['Eligibility Request Number']);
    return elig;
  }
  return null;
}

/*****************************
 * Normalize report rows & validate
 *****************************/
function normalizeReportDataFromParsed(parsed) {
  // parsed: { rows: [...], rawHeaders: [...] }
  const normalized = normalizeParsedSheet(parsed);
  // The normalizeParsedSheet returns objects with canonical keys already; just postprocess
  return normalized.map(r => ({
    claimID: r.claimID || '',
    memberID: r.memberID || '',
    claimDate: r.claimDate || '',
    clinician: r.clinician || '',
    department: r.department || '',
    packageName: r.packageName || '',
    insuranceCompany: r.insuranceCompany || '',
    claimStatus: r.claimStatus || ''
  }));
}

function validateReportClaims(reportData, eligMap) {
  const results = reportData.map(row => {
    if (!row || !row.claimID || String(row.claimID).trim() === '') return null;

    const memberID = String(row.memberID || '').trim();
    const claimDateRaw = row.claimDate;
    const claimDate = DateHandler.parse(claimDateRaw, { preferMDY: lastReportWasCSV });
    const formattedDate = DateHandler.format(claimDate);

    const isVVIP = memberID.startsWith('(VVIP)');
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
        fullEligibilityRecord: null,
        insuranceCompany: row.insuranceCompany || ''
      };
    }

    const hasLeadingZero = /^0+\d+$/.test(memberID);
    const claimClinicians = row.clinician ? [row.clinician] : [];
    const eligibility = findEligibilityForClaim(eligMap, claimDate, memberID, claimClinicians);

    const remarks = [];
    let finalStatus = 'invalid';

    if (hasLeadingZero) remarks.push('Member ID has a leading zero; claim marked as invalid.');

    if (!eligibility) {
      remarks.push(`No matching eligibility found for ${memberID} on ${formattedDate}`);
    } else if ((eligibility.Status || '').toLowerCase() !== 'eligible') {
      remarks.push(`Eligibility status: ${eligibility.Status}`);
    } else {
      const serviceCategory = eligibility['Service Category']?.trim() || '';
      const consultationStatus = eligibility['Consultation Status']?.trim()?.toLowerCase() || '';
      const dept = (row.department || row.clinic || '').toLowerCase();
      if (!isServiceCategoryValid(serviceCategory, consultationStatus, dept).valid) {
        remarks.push(`Invalid for category: ${serviceCategory}, department: ${row.department || row.clinic}`);
      } else if (!hasLeadingZero) {
        finalStatus = 'valid';
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
      finalStatus,
      fullEligibilityRecord: eligibility,
      insuranceCompany: eligibility?.['Payer Name'] || row.insuranceCompany || row.packageName || ''
    };
  });

  return results.filter(r => r);
}

/*****************************
 * Rendering + modal
 *****************************/
function renderResults(results, eligMap) {
  if (!resultsContainer) return;
  resultsContainer.innerHTML = '';
  if (!results || results.length === 0) {
    resultsContainer.innerHTML = '<div class="no-results">No claims to display</div>';
    return;
  }

  const filterOn = filterCheckbox && filterCheckbox.checked;
  const table = document.createElement('table');
  table.className = 'shared-table';
  table.style.width = '100%';

  const thead = document.createElement('thead');
  thead.innerHTML = `<tr>
    <th>Claim ID</th><th>Member ID</th><th>Encounter Date</th><th>Clinician</th><th>Service Category</th><th>Status</th><th>Remarks</th><th>Details</th>
  </tr>`;
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  const counts = { valid:0, invalid:0, unknown:0 };

  results.forEach((res, i) => {
    if (!res || !res.memberID || String(res.memberID).trim() === '') return;

    if (filterOn) {
      const p = (res.insuranceCompany || res.packageName || '').toLowerCase();
      if (!p.includes('daman') && !p.includes('thiqa')) return;
    }

    if (res.finalStatus && counts.hasOwnProperty(res.finalStatus)) counts[res.finalStatus]++;

    const tr = document.createElement('tr');
    tr.className = res.finalStatus || '';

    const statusBadge = res.status ? `<span class="status-badge ${String(res.status).toLowerCase() === 'eligible' ? 'eligible' : 'ineligible'}">${escapeHtml(res.status)}</span>` : '';

    const remarksHTML = res.remarks && res.remarks.length ? res.remarks.map(r => `<div>${escapeHtml(r)}</div>`).join('') : '<div class="source-note">No remarks</div>';

    let detailsCell = '<div class="source-note">N/A</div>';
    if (res.fullEligibilityRecord?.['Eligibility Request Number']) {
      detailsCell = `<button class="details-btn eligibility-details" data-index="${i}">${escapeHtml(res.fullEligibilityRecord['Eligibility Request Number'])}</button>`;
    } else {
      const norm = normalizeMemberID(res.memberID);
      if (eligMap && typeof eligMap.get === 'function' && (eligMap.get(norm) || []).length) {
        detailsCell = `<button class="details-btn show-all-eligibilities" data-member="${escapeHtml(res.memberID)}">View All</button>`;
      }
    }

    tr.innerHTML = `
      <td>${escapeHtml(res.claimID)}</td>
      <td>${escapeHtml(res.memberID)}</td>
      <td>${escapeHtml(res.encounterStart)}</td>
      <td>${escapeHtml(res.clinician)}</td>
      <td>${escapeHtml(res.serviceCategory)}</td>
      <td>${statusBadge}</td>
      <td>${remarksHTML}</td>
      <td>${detailsCell}</td>
    `;
    tbody.appendChild(tr);
  });

  table.appendChild(tbody);
  resultsContainer.appendChild(table);

  const summary = document.createElement('div');
  summary.className = 'loaded-count';
  summary.innerHTML = `Processed ${results.length} claims: <span class="valid">${counts.valid} valid</span>, <span class="unknown">${counts.unknown} unknown</span>, <span class="invalid">${counts.invalid} invalid</span>`;
  resultsContainer.prepend(summary);

  initModalHandlers(results, eligMap);
}

function initModalHandlers(results, eligMap) {
  if (!document.getElementById('modalOverlay')) {
    const modalHtml = `
      <div id="modalOverlay" style="display:none;position:fixed;z-index:9999;left:0;top:0;width:100vw;height:100vh;background:rgba(0,0,0,0.35);">
        <div id="modalContent" style="background:#fff;width:90%;max-width:900px;max-height:90vh;overflow:auto;position:absolute;left:50%;top:50%;transform:translate(-50%,-50%);padding:20px;border-radius:8px;">
          <button id="modalCloseBtn" style="float:right;font-size:18px;padding:2px 10px;cursor:pointer;" aria-label="Close">&times;</button>
          <div id="modalTable"></div>
        </div>
      </div>`;
    document.body.insertAdjacentHTML('beforeend', modalHtml);
    document.getElementById('modalCloseBtn').onclick = hideModal;
    document.getElementById('modalOverlay').onclick = (e) => { if (e.target.id === 'modalOverlay') hideModal(); };
  }

  document.querySelectorAll('.details-btn').forEach(btn => {
    btn.onclick = function() {
      if (this.classList.contains('eligibility-details')) {
        const idx = parseInt(this.dataset.index, 10);
        const r = results[idx];
        if (!r?.fullEligibilityRecord) return;
        document.getElementById('modalTable').innerHTML = formatEligibilityDetails(r.fullEligibilityRecord, r.memberID);
        document.getElementById('modalOverlay').style.display = 'block';
      } else if (this.classList.contains('show-all-eligibilities')) {
        const member = this.dataset.member;
        const normalized = normalizeMemberID(member);
        const list = eligMap.get(normalized) || [];
        if (!list.length) {
          document.getElementById('modalTable').innerHTML = `<div>No eligibilities found for ${escapeHtml(member)}</div>`;
          document.getElementById('modalOverlay').style.display = 'block';
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
        document.getElementById('modalTable').innerHTML = html;
        document.getElementById('modalOverlay').style.display = 'block';
      }
    };
  });
}
function hideModal(){ const o = document.getElementById('modalOverlay'); if (o) o.style.display = 'none'; }

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

/*****************************
 * Export invalid entries
 *****************************/
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

/*****************************
 * Handlers & initialization
 *****************************/
async function handleFileUpload(e, type) {
  const file = e.target.files && e.target.files[0];
  if (!file) return;
  try {
    updateStatus(`Loading ${type} file...`);
    if (type === 'eligibility') {
      // eligibility XSLX -> read sheet rows and convert to array-of-objects via XLSX utils
      const parsed = await parseFileByExtension(file);
      // parsed: { rows: [...], rawHeaders: [...] }
      // For eligData, we want the raw sheet -> use xlsx utils sheet_to_json directly for reliability
      // Simpler: read file as arraybuffer and use XLSX to produce json with header row
      const reader = new FileReader();
      reader.onload = function(ev) {
        const data = new Uint8Array(ev.target.result);
        const wb = XLSX.read(data, { type: 'array' });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: '' });
        eligData = json;
        updateStatus(`Loaded ${eligData.length} eligibility records`);
        updateProcessButtonState();
      };
      reader.onerror = () => { updateStatus('Error loading eligibility file'); };
      reader.readAsArrayBuffer(file);
      return;
    } else if (type === 'report') {
      lastReportWasCSV = file.name.toLowerCase().endsWith('.csv');
      const parsed = await parseFileByExtension(file);
      xlsData = normalizeReportDataFromParsed(parsed).filter(r => r && r.claimID && String(r.claimID).trim() !== '');
      updateStatus(`Loaded ${xlsData.length} report rows`);
      updateProcessButtonState();
    }
  } catch (err) {
    console.error('File load error:', err);
    updateStatus(`Error loading ${type} file`);
  }
}

async function parseFileByExtension(file) {
  const name = file.name.toLowerCase();
  if (name.endsWith('.csv')) {
    return await parseCsvFile(file);
  } else {
    // .xlsx or .xls
    return await parseExcelFile(file);
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
    xlsData = normalizeReportDataFromParsed(parsed).filter(r => r && r.claimID && String(r.claimID).trim() !== '');
    updateStatus(`Loaded ${xlsData.length} rows from pasted CSV`);
    updateProcessButtonState();
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
  if (!eligData) { alert('Please upload eligibility file first'); return; }
  if (!xlsData || !xlsData.length) { alert('Please upload report file first'); return; }
  try {
    updateStatus('Processing...');
    usedEligibilities.clear();
    const eligMap = prepareEligibilityMap(eligData);
    const results = validateReportClaims(xlsData, eligMap);
    window.lastValidationResults = results;
    renderResults(results, eligMap);
    updateStatus(`Processed ${results.length} claims successfully`);
  } catch (err) {
    console.error('Processing error:', err);
    updateStatus('Processing failed');
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
    renderResults(window.lastValidationResults, eligMap);
  }
}

// Initialize DOM refs and event listeners after DOMContentLoaded
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

  // initialize filter UI
  if (filterStatus) onFilterToggle();
}

document.addEventListener('DOMContentLoaded', () => {
  initializeEventListeners();
  updateStatus('Ready to process files');
});
