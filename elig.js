/*******************************
 * elig.js - robust parser & matcher
 * Replaces and improves previous file to handle xls/xlsx/csv + pasted CSV + header mapping
 *
 * Modified: adds tracking of invalid-file errors and prints a console summary
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

// Error tracking for invalid files (new)
const invalidFileErrorCounts = new Map(); // key -> count

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

function recordInvalidFileError(type) {
  const key = String(type || 'Unknown error').trim() || 'Unknown error';
  invalidFileErrorCounts.set(key, (invalidFileErrorCounts.get(key) || 0) + 1);
  // Log occurrence for debugging
  console.warn(`Invalid file error recorded: ${key} (count=${invalidFileErrorCounts.get(key)})`);
  // Print a compact summary each time an invalid file error is recorded
  printInvalidFileErrorSummary();
}

function printInvalidFileErrorSummary() {
  if (invalidFileErrorCounts.size === 0) return;
  // Create array sorted by count desc
  const entries = Array.from(invalidFileErrorCounts.entries()).sort((a,b) => b[1] - a[1]);
  const [topType, topCount] = entries[0];
  console.group('Invalid Files Error Summary');
  console.log(`Most frequent error: "${topType}" occurred ${topCount} time(s)`);
  console.log('All error counts (sorted):');
  entries.forEach(([k,v]) => console.log(`  ${v} × ${k}`));
  console.groupEnd();
}

function expandScientificNotation(val) {
  if (val === null || val === undefined) return '';
  const s = String(val).trim();
  if (!/[eE]/.test(s)) return s;
  const m = s.match(/^([+-]?[\d]+(?:\.[\d]+)?)[eE]([+-]?\d+)$/);
  if (!m) return s;
  let mant = m[1];
  const exp = parseInt(m[2], 10);
  const negative = mant.startsWith('-');
  if (negative) mant = mant.slice(1);
  const parts = mant.split('.');
  let digits = parts.join('');
  const decimals = parts[1] ? parts[1].length : 0;
  let zerosToAdd = exp - decimals;
  if (zerosToAdd >= 0) {
    digits = digits + '0'.repeat(zerosToAdd);
    digits = digits.replace(/^0+/, '') || '0';
  } else {
    const pos = digits.length + zerosToAdd;
    let left = digits.slice(0, pos);
    let right = digits.slice(pos);
    if (left === '') left = '0';
    digits = left + '.' + right;
    digits = digits.replace(/^0+([1-9])/, '$1') || digits;
  }
  return negative ? '-' + digits : digits;
}

function normalizeMemberID(id) {
  if (id === null || id === undefined) return '';
  let s = String(id).trim();
  s = s.replace(/^\uFEFF/, '');
  if (/[eE]/.test(s)) s = expandScientificNotation(s);
  if (/^\d+\.\d+$/.test(s)) s = s.split('.')[0];
  s = s.replace(/^0+/, '');
  return s;
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
    const cleanStr = String(input).trim().replace(/\uFEFF/g,'');
    if (/^\d+(\.\d+)?$/.test(cleanStr) && !cleanStr.includes('-') && !cleanStr.includes('/')) {
      const n = Number(cleanStr);
      if (!isNaN(n) && n > 59) {
        return this._parseExcelDate(n);
      }
    }
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
    try {
      const floatSerial = Number(serial);
      if (isNaN(floatSerial)) return null;
      const utcDays = Math.floor(floatSerial) - 25569;
      const ms = utcDays * 86400 * 1000;
      const d = new Date(ms);
      return new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
    } catch (e) {
      return null;
    }
  },
  _parseStringDate(dateStr, preferMDY = false) {
    if (!dateStr || typeof dateStr !== 'string') return null;
    if (dateStr.includes(' ')) dateStr = dateStr.split(' ')[0];
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
    const textual = dateStr.match(/^(\d{1,2})[\/\- ]([a-z]{3,})[\/\- ](\d{2,4})$/i);
    if (textual) {
      const day = parseInt(textual[1], 10);
      let year = parseInt(textual[3], 10);
      if (year < 100) year += 2000;
      const mon = MONTHS.indexOf(textual[2].toLowerCase().substr(0,3));
      if (mon >= 0) return new Date(Date.UTC(year, mon, day));
    }
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

function normalizeHeaderKey(h) {
  if (h === null || h === undefined) return '';
  return String(h).trim().replace(/\s+/g,' ').toLowerCase();
}

function detectHeaderRow(allRows, maxScan = 15) {
  const rows = Array.isArray(allRows) ? allRows : [];
  for (let i = 0; i < Math.min(maxScan, rows.length); i++) {
    const row = rows[i] || [];
    const nonEmpty = row.filter(c => String(c).trim() !== '').length;
    if (nonEmpty >= 3) {
      const joined = row.join(' ').toLowerCase();
      if (joined.includes('pri') || joined.includes('claim') || joined.includes('card') || joined.includes('member') || joined.includes('patient')) {
        return i;
      }
    }
  }
  for (let i=0;i<Math.min(maxScan, rows.length);i++){
    const row = rows[i] || [];
    const nonEmpty = row.filter(c => String(c).trim() !== '').length;
    if (nonEmpty>0) return i;
  }
  return 0;
}

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
  return mapped;
}

function rowArrayToNormalizedObject(rowArray, headerMap, rawHeaders) {
  const obj = {};
  for (let i=0;i<headerMap.length;i++) {
    const canon = headerMap[i];
    const rawHeader = rawHeaders[i] || (`Column${i+1}`);
    const rawVal = rowArray[i] === undefined || rowArray[i] === null ? '' : rowArray[i];
    if (canon) {
      obj[canon] = rawVal;
    } else {
      obj[rawHeader] = rawVal;
    }
  }
  return obj;
}

/*****************************
 * Parsing: Excel (.xls/.xlsx) and CSV (.csv) & pasted CSV
 *****************************/
function parseExcelArrayBuffer(arrayBuffer) {
  try {
    const wb = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const allRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    if (!Array.isArray(allRows) || allRows.length === 0) {
      recordInvalidFileError('Empty or invalid Excel sheet');
      return { rows: [], rawHeaders: [] };
    }
    return parseSheetRows(allRows);
  } catch (err) {
    recordInvalidFileError(`Excel parse error: ${err && err.message ? err.message : err}`);
    throw err;
  }
}

function parseCsvTextString(text) {
  try {
    const clean = (text || '').replace(/^\uFEFF/, '');
    const wb = XLSX.read(clean, { type: 'string' });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const allRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    if (!Array.isArray(allRows) || allRows.length === 0) {
      recordInvalidFileError('Empty or invalid CSV content');
      return { rows: [], rawHeaders: [] };
    }
    return parseSheetRows(allRows);
  } catch (err) {
    recordInvalidFileError(`CSV parse error: ${err && err.message ? err.message : err}`);
    throw err;
  }
}

function parseSheetRows(allRows) {
  if (!Array.isArray(allRows) || allRows.length === 0) {
    recordInvalidFileError('No rows found in sheet');
    return { rows: [], rawHeaders: [] };
  }

  const headerRowIndex = detectHeaderRow(allRows, 20);
  const rawHeaderRow = (allRows[headerRowIndex] || []).map(h => String(h).trim());
  if (!rawHeaderRow || rawHeaderRow.length === 0) {
    recordInvalidFileError('Header row not detected');
    return { rows: [], rawHeaders: [] };
  }
  const mapped = mapHeadersToCanonical(rawHeaderRow);
  const headerMap = [];
  for (let i=0;i<rawHeaderRow.length;i++) {
    headerMap[i] = mapped[i] || null;
  }

  const dataRows = allRows.slice(headerRowIndex + 1);
  if (!dataRows || dataRows.length === 0) {
    recordInvalidFileError('No data rows after header');
  }
  const normalizedRows = dataRows.map(rowArr => {
    const normObj = rowArrayToNormalizedObject(rowArr, headerMap, rawHeaderRow);
    return normObj;
  });

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
      } catch (err) { 
        recordInvalidFileError(`File read/Excel parse failed: ${err && err.message ? err.message : err}`);
        reject(err); 
      }
    };
    reader.onerror = () => { 
      recordInvalidFileError('FileReader error reading Excel file');
      reject(reader.error);
    };
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
      } catch (err) { 
        recordInvalidFileError(`File read/CSV parse failed: ${err && err.message ? err.message : err}`);
        reject(err); 
      }
    };
    reader.onerror = () => {
      recordInvalidFileError('FileReader error reading CSV file');
      reject(reader.error);
    };
    reader.readAsText(file);
  });
}

// parse pasted CSV text
function parseCsvText(text) {
  return new Promise((resolve, reject) => {
    try {
      const parsed = parseCsvTextString(text);
      resolve(parsed);
    } catch (err) { 
      recordInvalidFileError(`Pasted CSV parse failed: ${err && err.message ? err.message : err}`);
      reject(err); 
    }
  });
}

/*****************************
 * Normalize parsed rows to canonical report rows
 *****************************/
function getField(obj, candidates) {
  for (const k of candidates) {
    if (obj && Object.prototype.hasOwnProperty.call(obj, k) && obj[k] !== '' && obj[k] !== null && obj[k] !== undefined) return obj[k];
  }
  return '';
}

function normalizeParsedSheet(parsed) {
  const rows = parsed.rows || [];
  const rawHeaders = parsed.rawHeaders || [];

  const normalized = rows.map(r => {
    const out = {
      claimID: r.claimID || getField(r, ['ClaimID','Pri. Claim No','Pri. Claim ID','Claim ID','Claim No']) || '',
      memberID: r.memberID || getField(r, ['PatientCardID','Patient Insurance Card No','PatientInsuranceCardNo','Card Number / DHA Member ID','Card Number','MemberID','Member ID']) || '',
      claimDate: r.claimDate || getField(r, ['Encounter Date','ClaimDate','Adm/Reg. Date','EncounterDate','Date']) || '',
      clinician: r.clinician || getField(r, ['Clinician License','Clinician','Admitting License']) || '',
      department: r.department || getField(r, ['Department','Clinic','Admitting Department']) || '',
      packageName: r.packageName || getField(r, ['Pri. Payer Name','Insurance Company','Pri. Plan Type','Package']) || '',
      insuranceCompany: r.insuranceCompany || getField(r, ['Payer Name','Insurance Company','Pri. Payer Name']) || '',
      claimStatus: r.claimStatus || getField(r, ['Codification Status','VisitStatus','Status','Claim Status']) || ''
    };
    if (!out.memberID) {
      for (const h of rawHeaders) {
        const val = r[h];
        if (val && String(h).toLowerCase().includes('card')) { out.memberID = val; break; }
      }
    }
    if (!out.claimID) {
      for (const h of rawHeaders) {
        const val = r[h];
        if (val && String(h).toLowerCase().includes('claim')) { out.claimID = val; break; }
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
    const rawID = normalizeMemberID(getField(e, [
      'Card Number / DHA Member ID','Card Number','_5','MemberID','Member ID','Patient Insurance Card No',
      'PatientCardID','CardNumber','Patient Insurance Card No'
    ]) || '');

    if (!rawID) return;
    const memberID = normalizeMemberID(rawID);
    if (!eligMap.has(memberID)) eligMap.set(memberID, []);
    eligMap.get(memberID).push({
      'Eligibility Request Number': getField(e, ['Eligibility Request Number','Eligibility Request No','Request Number']) || '',
      'Card Number / DHA Member ID': rawID,
      'Answered On': getField(e, ['Answered On','AnsweredOn','Answered Date']) || '',
      'Ordered On': getField(e, ['Ordered On','OrderedOn']) || '',
      'Status': getField(e, ['Status']) || '',
      'Clinician': getField(e, ['Clinician','Provider']) || '',
      'Payer Name': getField(e, ['Payer Name','PayerName']) || '',
      'Service Category': getField(e, ['Service Category']) || '',
      'Package Name': getField(e, ['Package Name','PackageName']) || '',
      'Department': getField(e, ['Department','Clinic']) || '',
      'Consultation Status': getField(e, ['Consultation Status','ConsultationStatus']) || ''
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

function findEligibilityForClaim(eligMap, claimDate, memberID, claimClinicians = []) {
  const normalizedID = normalizeMemberID(memberID);
  const eligList = eligMap.get(normalizedID) || [];
  if (!eligList.length) return null;

  const claimCliniciansFiltered = Array.isArray(claimClinicians) ? claimClinicians.filter(Boolean) : [];

  for (const elig of eligList) {
    const eligDate = DateHandler.parse(elig['Answered On'] || elig['Ordered On']);
    if (!DateHandler.isSameDay(claimDate, eligDate)) {
      continue;
    }

    if (elig.Clinician && claimCliniciansFiltered.length && !checkClinicianMatch(claimCliniciansFiltered, elig.Clinician)) {
      continue;
    }

    const serviceCategory = elig['Service Category'] || '';
    const consultationStatus = elig['Consultation Status'] || '';
    const dept = (elig.Department || elig.Clinic || '').toLowerCase();
    const svcCheck = isServiceCategoryValid(serviceCategory, consultationStatus, dept);
    if (!svcCheck.valid) {
      continue;
    }

    if ((elig.Status || '').toLowerCase() !== 'eligible') {
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
  const normalized = normalizeParsedSheet(parsed);
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

    const memberIDRaw = String(row.memberID || '').trim();
    const memberID = normalizeMemberID(memberIDRaw);
    const claimDateRaw = row.claimDate;
    const claimDate = DateHandler.parse(claimDateRaw, { preferMDY: lastReportWasCSV });
    const formattedDate = DateHandler.format(claimDate);

    const isVVIP = memberIDRaw.startsWith('(VVIP)');
    if (isVVIP) {
      return {
        claimID: row.claimID,
        memberID: memberIDRaw,
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

    const hasLeadingZero = /^0+\d+$/.test(memberIDRaw);
    const claimClinicians = row.clinician ? [row.clinician] : [];
    const eligibility = findEligibilityForClaim(eligMap, claimDate, memberIDRaw, claimClinicians);

    const remarks = [];
    let finalStatus = 'invalid';

    if (hasLeadingZero) remarks.push('Member ID has a leading zero; claim marked as invalid.');

    if (!eligibility) {
      remarks.push(`No matching eligibility found for ${memberIDRaw} on ${formattedDate}`);
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
      memberID: memberIDRaw,
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
        <div id="modalContent" style="background:#fff;width:90%;max-width:900px;max-height:90vh;overflow:auto;position:absolute;left:50%;top:50%;transform:translate(-50%,-50%);padding:20px;border-radius:6px;">
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
  let html = `<div style="margin-bottom:8px;"><strong>Member:</strong> ${escapeHtml(memberID)} <span style="margin-left:8px;" class="status-badge ${(String((record.Status||'')).toLowerCase()==='eligible')?'eligible':'ineligible'}">${escapeHtml(record.Status||'')}</span></div>`;
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
      const parsed = await parseFileByExtension(file);
      const reader = new FileReader();
      reader.onload = function(ev) {
        try {
          const data = new Uint8Array(ev.target.result);
          const wb = XLSX.read(data, { type: 'array' });
          const sheet = wb.Sheets[wb.SheetNames[0]];
          const json = XLSX.utils.sheet_to_json(sheet, { defval: '' });
          if (!Array.isArray(json) || json.length === 0) {
            recordInvalidFileError('Eligibility file contained no rows');
          }
          eligData = json;
          updateStatus(`Loaded ${eligData.length} eligibility records`);
          updateProcessButtonState();
        } catch (innerErr) {
          recordInvalidFileError(`Eligibility file read->json conversion failed: ${innerErr && innerErr.message ? innerErr.message : innerErr}`);
          updateStatus('Error loading eligibility file');
        }
      };
      reader.onerror = () => { 
        recordInvalidFileError('FileReader error loading eligibility file');
        updateStatus('Error loading eligibility file'); 
      };
      reader.readAsArrayBuffer(file);
      return;
    } else if (type === 'report') {
      lastReportWasCSV = file.name.toLowerCase().endsWith('.csv');
      const parsed = await parseFileByExtension(file);
      xlsData = normalizeReportDataFromParsed(parsed).filter(r => r && r.claimID && String(r.claimID).trim() !== '');
      if (!xlsData || xlsData.length === 0) {
        recordInvalidFileError('Report file contained no recognizable claim rows');
      }
      updateStatus(`Loaded ${xlsData.length} report rows`);
      updateProcessButtonState();
    }
  } catch (err) {
    console.error('File load error:', err);
    recordInvalidFileError(`File load error: ${err && err.message ? err.message : err}`);
    updateStatus(`Error loading ${type} file`);
  }
}

async function parseFileByExtension(file) {
  const name = file.name.toLowerCase();
  if (name.endsWith('.csv')) {
    return await parseCsvFile(file);
  } else {
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
    if (!xlsData || xlsData.length === 0) recordInvalidFileError('Pasted CSV contained no recognizable claim rows');
    updateStatus(`Loaded ${xlsData.length} rows from pasted CSV`);
    updateProcessButtonState();
  } catch (err) {
    console.error('Error parsing pasted CSV:', err);
    recordInvalidFileError(`Pasted CSV parse error: ${err && err.message ? err.message : err}`);
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
    // After a successful processing run, also print the invalid-file error summary (if any)
    if (invalidFileErrorCounts.size > 0) {
      console.info('Summary of invalid-file errors encountered during this session:');
      printInvalidFileErrorSummary();
    }
  } catch (err) {
    console.error('Processing error:', err);
    recordInvalidFileError(`Processing error: ${err && err.message ? err.message : err}`);
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
