/*******************************
 * elig.js (reordered)
 *
 * Reorganized for readability: constants & state, small utilities,
 * date handling, parsing utilities, eligibility map builders,
 * validation logic, rendering & modal, exports, handlers, and init.
 *
 * Behavior preserved; only function order changed.
 *******************************/

/* ===========================
   Version & Initialization
   =========================== */
const VERSION = '2026.02.15.35';
console.log(`✅ Eligibility Checker v${VERSION} loaded successfully`);

/* ===========================
   Constants & Application State
   =========================== */
const SERVICE_PACKAGE_RULES = {
  'Dental Services': ['dental', 'orthodontic'],
  'Physiotherapy': ['physio'],
  'Other OP Services': ['physio', 'diet', 'occupational', 'speech'],
  'Consultation': []  // Special handling below
};
const DATE_KEYS = ['Date', 'On'];
const MONTHS = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"];
const MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

// Application state
let xlsData = null;        // parsed & normalized report rows
let eligData = null;       // eligibility sheet as array of arrays (raw) — keep raw rows for header detection
let rawParsedReport = null; // raw parsed sheet result (header detection output)
let detectedReportType = 'Generic'; // detected report type before normalization
const usedEligibilities = new Set();
let lastReportWasCSV = false;

// Keep last eligibility map so UI filters can re-render without rebuilding the map
let lastEligMap = null;

// Option to remove leading zeroes from member IDs and claim IDs
let removeLeadingZeroes = false;

// DOM Elements (lookups performed in initializeEventListeners)
let reportInput, eligInput, processBtn, exportInvalidBtn, statusEl, resultsContainer, filterCheckbox, filterStatus, pasteTextarea, pasteBtn;
let removeZeroesCheckbox, removeZeroesStatus;

/* ===========================
   Small Utilities
   =========================== */
function escapeHtml(s) {
  if (s === null || s === undefined) return '';
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#x27;');
}

function normalizeMemberID(id) {
  if (!id) return "";
  let normalized = String(id).replace(/\D/g, "").trim();
  
  // Optionally remove leading zeroes if the option is enabled
  // Note: If input is all zeroes (e.g., "0000"), we keep at least one "0" 
  // to maintain a valid ID rather than returning an empty string
  if (removeLeadingZeroes && normalized.length > 0) {
    normalized = normalized.replace(/^0+/, '') || '0';
  }
  
  return normalized;
}

function normalizeClinician(name) {
  if (!name) return '';
  return name.trim().toLowerCase().replace(/\s+/g, ' ');
}

/**
 * Check if package names match with special handling for Thiqa/TC packages
 * @param {string} claimPackage - Package name from the claim
 * @param {string} eligPackage - Package name from the eligibility
 * @returns {boolean} - True if packages match
 */
function packageNamesMatch(claimPackage, eligPackage) {
  if (!claimPackage || !eligPackage) return false;
  
  const claimLower = claimPackage.trim().toLowerCase();
  const eligLower = eligPackage.trim().toLowerCase();
  
  // Direct match
  if (claimLower === eligLower) return true;
  
  // Special Thiqa/TC matching: if claim has "thiqa" and eligibility has "tc", consider it a match
  if (claimLower.includes('thiqa') && eligLower.includes('tc')) {
    return true;
  }
  
  return false;
}

/* ===========================
   Date handling (DateHandler)
   =========================== */
const DateHandler = {
  parse: function(input, options = {}) {
    const preferMDY = !!options.preferMDY;
    const debugLog = !!options.debugLog;
    if (!input) return null;
    if (input instanceof Date) return isNaN(input) ? null : input;
    if (typeof input === 'number') return this._parseExcelDate(input);

    // Check if string input is an Excel serial number BEFORE cleaning
    // This must happen first to preserve decimal precision (e.g., "46358.00013888889")
    // Otherwise, the period would be stripped, breaking Excel serial date conversion
    const inputStr = input.toString().trim();
    if (/^\d+\.?\d*$/.test(inputStr)) {
      const numericValue = parseFloat(inputStr);
      if (!isNaN(numericValue)) {
        return this._parseExcelDate(numericValue);
      }
    }

    const cleanStr = inputStr.replace(/[,.]/g, '');
    const parsed = this._parseStringDate(cleanStr, preferMDY, debugLog) || new Date(cleanStr);
    if (isNaN(parsed)) {
      // Removed console warning - only log via debugLog flag
      return null;
    }
    return parsed;
  },

  format: function(date) {
    if (!(date instanceof Date) || isNaN(date)) return '';
    const day = date.getUTCDate(); // No padding, just the number
    const monthIndex = date.getUTCMonth();
    const monthName = MONTH_NAMES[monthIndex];
    const year = date.getUTCFullYear();
    return `${day} ${monthName} ${year}`;
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
    // Get UTC components and return as-is (no swap logic)
    return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
  },

  _swapDayMonth: function(date) {
    // Helper function to swap day and month in a date
    // Used ONLY for Insta report claim dates which have backwards dates
    if (!date || isNaN(date.getTime())) return date;
    
    const day = date.getUTCDate();
    const month = date.getUTCMonth(); // 0-based
    
    // Only swap if both day and month are ambiguous (≤12)
    if (day <= 12 && month < 12) {
      return new Date(Date.UTC(
        date.getUTCFullYear(),
        day - 1,      // Use day as month (0-based)
        month + 1     // Use month+1 as day (1-based)
      ));
    }
    return date;
  },

  _normalizeTwoDigitYear: function(year) {
    // Handle 2-digit years: assume 2000s for years 0-99
    // This assumes medical records are from recent years (2000-2099)
    if (year < 100) {
      return year + 2000;
    }
    return year;
  },

  _parseStringDate: function(dateStr, preferMDY = false, debugLog = false) {
    if (!dateStr) return null;
    
    // Insta CSV sometimes includes timestamps like "13-02-2026 05:10:14"
    // Strip time portion if present (keep only date before first space)
    if (dateStr.includes(' ')) {
      if (debugLog) console.log(`    [Parse] Stripping timestamp: "${dateStr}" → "${dateStr.split(' ')[0]}"`);
      dateStr = dateStr.split(' ')[0];
    }

    // Try matching numeric date formats: DD/MM/YYYY, DD-MM-YYYY, etc.
    const dmyMdyMatch = dateStr.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
    if (dmyMdyMatch) {
      const part1 = parseInt(dmyMdyMatch[1], 10);
      const part2 = parseInt(dmyMdyMatch[2], 10);
      const rawYear = parseInt(dmyMdyMatch[3], 10);
      const year = this._normalizeTwoDigitYear(rawYear);
      
      if (debugLog) {
        console.log(`    [Parse] Numeric date matched: part1=${part1}, part2=${part2}, rawYear=${rawYear}, normalizedYear=${year}`);
      }
      
      if (part1 > 12 && part2 <= 12) {
        // Unambiguous DMY (day > 12, so first part must be day)
        if (debugLog) console.log(`    [Parse] Unambiguous DMY: day=${part1}, month=${part2}`);
        return new Date(Date.UTC(year, part2 - 1, part1)); // dmy
      } else if (part2 > 12 && part1 <= 12) {
        // Unambiguous MDY (second part > 12, so it must be day)
        if (debugLog) console.log(`    [Parse] Unambiguous MDY: month=${part1}, day=${part2}`);
        return new Date(Date.UTC(year, part1 - 1, part2)); // mdy
      } else {
        // Ambiguous case: both parts are <= 12, could be either day or month
        // Default to DMY (day/month/year) format which is standard
        if (debugLog) console.log(`    [Parse] Ambiguous case (both ≤12): using DMY format - day=${part1}, month=${part2}`);
        return new Date(Date.UTC(year, part2 - 1, part1)); // DMY format
      }
    }

    // Try matching text-based dates: "2-Dec-26", "12-Feb-2026", etc.
    const textMatch = dateStr.match(/^(\d{1,2})[\/\- ]([a-z]{3,})[\/\- ](\d{2,4})$/i);
    if (textMatch) {
      const day = parseInt(textMatch[1], 10);
      const monthName = textMatch[2].toLowerCase().substr(0, 3);
      const monthIndex = MONTHS.indexOf(monthName);
      const rawYear = parseInt(textMatch[3], 10);
      const year = this._normalizeTwoDigitYear(rawYear);
      
      if (debugLog) {
        console.log(`    [Parse] Text date matched: day=${day}, monthName=${monthName} (index=${monthIndex}), rawYear=${rawYear}, normalizedYear=${year}`);
      }
      
      if (monthIndex >= 0 && monthIndex < 12 && day >= 1 && day <= 31) {
        if (debugLog) console.log(`    [Parse] Final text date: day=${day}, month=${MONTHS[monthIndex]} (index=${monthIndex}), year=${year}`);
        return new Date(Date.UTC(year, monthIndex, day));
      }
    }

    // Try matching ISO format: YYYY-MM-DD or YYYY/MM/DD
    const isoMatch = dateStr.match(/^(\d{4})[\/\-](\d{2})[\/\-](\d{2})$/);
    if (isoMatch) {
      if (debugLog) console.log(`    [Parse] ISO format matched: ${isoMatch[1]}-${isoMatch[2]}-${isoMatch[3]}`);
      return new Date(Date.UTC(parseInt(isoMatch[1], 10), parseInt(isoMatch[2], 10) - 1, parseInt(isoMatch[3], 10)));
    }
    
    if (debugLog) console.log(`    [Parse] No date pattern matched!`);
    return null;
  }
};

/* ===========================
   Summary helper (used after files load)
   =========================== */
function summarizeAndDisplayCounts() {
  try {
    const eligCount = Array.isArray(eligData) ? eligData.length : 0;

    // Ensure xlsData exists; if not but rawParsedReport exists try to normalize it now
    if ((!Array.isArray(xlsData) || xlsData.length === 0) && rawParsedReport) {
      try {
        const normalized = normalizeReportData(rawParsedReport);
        xlsData = normalized.filter(r => r && r.claimID && String(r.claimID).trim() !== '');
      } catch (e) {
        // Removed console warning - fails silently
      }
    }

    const claimCount = Array.isArray(xlsData) ? xlsData.length : 0;

    if (statusEl) {
      statusEl.textContent = `Loaded ${eligCount} eligibilities, ${claimCount} claims — Ready to process files`;
    }
  } catch (err) {
    // Removed console error - fails silently
  }
}

/* ===========================
   Header detection helper (array-of-arrays)
   =========================== */
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

    if (score > bestScore) {
      bestScore = score;
      bestIndex = i;
    }
  }

  const headerRowIndex = bestScore > 0 ? bestIndex : 0;
  const rawHeaderRow = allRows[headerRowIndex] || [];
  
  // Build a mapping of non-empty headers to their column indices
  const headerMapping = [];
  for (let c = 0; c < rawHeaderRow.length; c++) {
    const headerValue = rawHeaderRow[c];
    const headerStr = (headerValue === null || headerValue === undefined || headerValue === '') 
      ? '' 
      : String(headerValue).trim();
    
    // Only include non-empty headers
    if (headerStr !== '') {
      headerMapping.push({ index: c, name: headerStr });
    }
  }
  
  const headers = headerMapping.map(h => h.name);
  const dataRows = allRows.slice(headerRowIndex + 1);

  const rows = dataRows.map(rowArray => {
    const obj = {};
    // Map only the non-empty header columns to the object
    for (const { index, name } of headerMapping) {
      obj[name] = rowArray[index] === undefined || rowArray[index] === null ? '' : rowArray[index];
    }
    return obj;
  });
  
  return { headerRowIndex, headers, rows };
}

/* ===========================
   File parsing helpers
   (Excel/CSV -> array-of-objects or array-of-arrays where needed)
   =========================== */
async function parseExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = function(e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const allRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
        resolve(allRows);
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
        resolve(allRows);
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
      resolve(allRows);
    } catch (err) {
      reject(err);
    }
  });
}

/* ===========================
   Eligibility map builders
   (Accepts array-of-arrays or array-of-objects)
   =========================== */
function prepareEligibilityMap(rawSheetArray) {
  if (!Array.isArray(rawSheetArray) || rawSheetArray.length === 0) return new Map();

  // If rows are arrays -> detect header and convert to objects
  if (Array.isArray(rawSheetArray[0])) {
    // find header row
    let headerRowIndex = rawSheetArray.findIndex(row =>
      Array.isArray(row) && row.some(cell => String(cell || '').trim().toLowerCase().includes('eligibility request number'))
    );
    if (headerRowIndex === -1) {
      headerRowIndex = rawSheetArray.findIndex(row => Array.isArray(row) && row.some(cell => String(cell || '').trim() !== ''));
    }
    if (headerRowIndex === -1) return new Map();

    const headers = (rawSheetArray[headerRowIndex] || []).map(h => String(h || '').trim());
    const eligMap = new Map();
    
    const totalRows = rawSheetArray.length - headerRowIndex - 1;
    console.log(`📥 Building eligibility map from ${totalRows} eligibility records...`);
    console.log(`📋 Header row found at index ${headerRowIndex}`);
    console.log(`📋 Column headers (first 15):`, headers.slice(0, 15).join(', '));
    
    // Show first 3 raw data rows to inspect actual file structure
    console.log(`\n🔍 RAW DATA INSPECTION - First 3 data rows from file:`);
    for (let inspectIdx = headerRowIndex + 1; inspectIdx < Math.min(headerRowIndex + 4, rawSheetArray.length); inspectIdx++) {
      const inspectRow = rawSheetArray[inspectIdx];
      if (!Array.isArray(inspectRow)) continue;
      const inspectRecord = {};
      headers.forEach((h, idx) => {
        if (inspectRow[idx] !== undefined && inspectRow[idx] !== null && inspectRow[idx] !== '') {
          inspectRecord[h] = inspectRow[idx];
        }
      });
      console.log(`   Row ${inspectIdx} (${Object.keys(inspectRecord).length} populated columns):`);
      // Show all column names and first few chars of values
      const sampleCols = Object.entries(inspectRecord).slice(0, 10);
      sampleCols.forEach(([col, val]) => {
        const valStr = String(val);
        const preview = valStr.length > 30 ? valStr.substring(0, 30) + '...' : valStr;
        console.log(`      "${col}": "${preview}"`);
      });
      if (Object.keys(inspectRecord).length > 10) {
        console.log(`      ... and ${Object.keys(inspectRecord).length - 10} more columns`);
      }
    }
    console.log(``);
    
    let recordCount = 0;
    let columnUsed = '';
    let skippedNoMemberID = 0;
    let skippedEmptyMemberID = 0;
    const firstSkippedSamples = [];

    for (let i = headerRowIndex + 1; i < rawSheetArray.length; i++) {
      const row = rawSheetArray[i];
      if (!Array.isArray(row)) continue;

      const record = {};
      headers.forEach((h, idx) => record[h] = row[idx] !== undefined ? row[idx] : '');

      const idCandidates = [
        'Card Number / DHA Member ID', 'Card Number', 'MemberID', 'Member ID',
        'Patient Insurance Card No', 'Policy1', 'Policy 1', 'PatientCardID'
      ];
      let rawMemberID = '';
      for (const k of idCandidates) {
        if (Object.prototype.hasOwnProperty.call(record, k) && record[k]) {
          rawMemberID = String(record[k]).trim();
          if (!columnUsed) columnUsed = k; // Track which column was used
          break;
        }
      }
      if (!rawMemberID) {
        skippedNoMemberID++;
        if (firstSkippedSamples.length < 10) {
          firstSkippedSamples.push({row: i, reason: 'No member ID found (card number is blank)', raw: rawMemberID});
        }
        continue;
      }
      const memberID = normalizeMemberID(rawMemberID);
      if (!memberID) {
        skippedEmptyMemberID++;
        if (firstSkippedSamples.length < 10) {
          firstSkippedSamples.push({row: i, reason: 'Empty after normalization', raw: rawMemberID});
        }
        continue;
      }

      // Log first eligibility with column info
      if (recordCount === 0) {
        console.log(`📋 Using "${columnUsed}" column for member identification`);
        console.log(`  Sample record:`, {[columnUsed]: rawMemberID, 'Request Number': record['Request Number'] || record['RequestNo'], Status: record['Status']});
      }
      
      // Log first 10 eligibilities to show mapping process
      recordCount++;
      if (recordCount <= 10) {
        console.log(`  Elig ${recordCount}: Raw="${rawMemberID}" → Normalized="${memberID}" → Map key="${memberID}"`);
      }

      if (!eligMap.has(memberID)) eligMap.set(memberID, []);
      eligMap.get(memberID).push(record);
    }

    // Show skip statistics
    const totalSkipped = skippedNoMemberID + skippedEmptyMemberID;
    if (totalSkipped > 0) {
      console.log(`\n⚠️ WARNING: ${totalSkipped} rows skipped during map building:`);
      console.log(`   No member ID found (card number is blank): ${skippedNoMemberID}`);
      console.log(`   Empty after normalization: ${skippedEmptyMemberID}`);
      if (firstSkippedSamples.length > 0) {
        console.log(`\n   First ${firstSkippedSamples.length} skipped rows:`);
        firstSkippedSamples.forEach(s => {
          console.log(`   Row ${s.row}: ${s.reason}, Raw="${s.raw}"`);
        });
      }
    }

    // Show ALL unique member IDs
    const allKeys = Array.from(eligMap.keys());
    console.log(`\n✅ Map built with ${eligMap.size} unique member IDs`);
    console.log(`   Total rows in file: ${totalRows}`);
    console.log(`   Successfully processed: ${recordCount}`);
    console.log(`   Skipped: ${totalSkipped}`);
    console.log(`   Full list: ${allKeys.join(', ')}`);
    return eligMap;
  }

  // Otherwise assume array-of-objects
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

  return eligMap;
}

/* ===========================
   Matching & Validation Utilities
   =========================== */

/**
 * Check if provider is Daman or Thiqa
 * @param {string} provider - Provider/insurance name
 * @returns {boolean} - True if Daman or Thiqa
 */
function isDamanOrThiqa(provider) {
  if (!provider) return false;
  const providerLower = provider.toString().toLowerCase();
  return providerLower.includes('daman') || providerLower.includes('thiqa');
}

/**
 * Find matching eligibility record for a claim
 * @param {Map} eligMap - Map of member IDs to eligibility records
 * @param {Date} claimDate - Claim date
 * @param {string} memberID - Member ID
 * @param {Array} claimClinicians - Array of clinician names/IDs
 * @param {string} provider - Provider/insurance name (used to filter diagnostic logging to Daman/Thiqa only)
 * @param {boolean} forceLog - Force diagnostic logging regardless of provider
 * @param {number} claimIndex - Index of the claim (for logging first 3)
 * @returns {Object|null} - Matching eligibility record or null
 */
function findEligibilityForClaim(eligMap, claimDate, memberID, claimClinicians = [], provider = '', forceLog = false, claimIndex = 0) {
  const normalizedID = normalizeMemberID(memberID || '');
  
  // Log lookup process for first 3 claims
  const shouldLog = claimIndex > 0 && claimIndex <= 3;
  if (shouldLog) {
    console.log(`🔍 CLAIM #${claimIndex} - Member ID Lookup:`);
    console.log(`  1️⃣ Raw from claim: "${memberID}"`);
    console.log(`  2️⃣ Normalized: "${normalizedID}"`);
    console.log(`  3️⃣ Lookup: eligibilityMap["${normalizedID}"]`);
  }
  
  const eligList = eligMap.get(normalizedID) || [];
  
  if (shouldLog) {
    console.log(`  4️⃣ Result: ${eligList.length > 0 ? `Found ${eligList.length} eligibilities` : 'undefined (NOT FOUND)'}`);
  }
  
  if (!eligList.length) {
    if (shouldLog) {
      if (eligMap.size === 0) {
        console.log(`  ❌ Eligibility map is EMPTY (no eligibilities loaded)`);
      } else {
        console.log(``);
        console.log(`  🔎 Exhaustive search through ALL ${eligMap.size} member IDs:`);
        
        const allMapKeys = Array.from(eligMap.keys());
        
        // Exact match (should be false if we're here)
        const exactMatch = allMapKeys.includes(normalizedID);
        console.log(`     • Exact match: ${exactMatch ? 'YES' : 'NO'}`);
        
        // Partial match - map ID contains claim ID
        const partialMatches = allMapKeys.filter(k => k.includes(normalizedID));
        if (partialMatches.length > 0) {
          console.log(`     • Partial (map contains "${normalizedID}"): ${partialMatches.slice(0, 3).join(', ')}`);
        } else {
          console.log(`     • Partial (map contains "${normalizedID}"): NO`);
        }
        
        // Reverse match - claim ID contains map ID
        const reverseMatches = allMapKeys.filter(k => normalizedID.includes(k));
        if (reverseMatches.length > 0) {
          console.log(`     • Reverse ("${normalizedID}" contains map ID): ${reverseMatches.slice(0, 3).join(', ')}`);
        } else {
          console.log(`     • Reverse ("${normalizedID}" contains map ID): NO`);
        }
        
        // Prefix match
        if (normalizedID.length >= 3) {
          const prefix = normalizedID.substring(0, 3);
          const prefixMatches = allMapKeys.filter(k => k.startsWith(prefix));
          if (prefixMatches.length > 0) {
            console.log(`     • Starts with "${prefix}": ${prefixMatches.slice(0, 3).join(', ')}`);
          } else {
            console.log(`     • Starts with "${prefix}": NO`);
          }
        }
        
        // Suffix match
        if (normalizedID.length >= 3) {
          const suffix = normalizedID.substring(normalizedID.length - 3);
          const suffixMatches = allMapKeys.filter(k => k.endsWith(suffix));
          if (suffixMatches.length > 0) {
            console.log(`     • Ends with "${suffix}": ${suffixMatches.slice(0, 3).join(', ')}`);
          } else {
            console.log(`     • Ends with "${suffix}": NO`);
          }
        }
        
        console.log(``);
        console.log(`  ❌ Member "${normalizedID}" not found in ANY form in the map`);
        
        const sampleIDs = allMapKeys.slice(0, 5);
        console.log(`  Sample IDs in map: ${sampleIDs.join(', ')}`);
      }
    }
    return null;
  }
  
  if (shouldLog) {
    const clinInfo = claimClinicians.length > 0 ? `, Clinicians: ${claimClinicians.join(', ')}` : '';
    console.log(`⚠️ CLAIM #${claimIndex}: Found ${eligList.length} eligibilities for member ${memberID}, Date: ${DateHandler.format(claimDate)}${clinInfo}`);
  }
  
  let eligIndex = 0;
  for (const elig of eligList) {
    eligIndex++;
    const eligDate = DateHandler.parse(elig["Answered On"], { preferMDY: false });
    
    if (shouldLog) {
      const eligNum = elig["Eligibility Request Number"] || "(unknown)";
      const eligDateStr = elig["Answered On"] || '(empty)';
      const status = elig.Status || '(empty)';
      const clinician = elig.Clinician || '(none)';
      
      if (eligDate) {
        const eligFormattedDate = DateHandler.format(eligDate);
        const matches = DateHandler.isSameDay(claimDate, eligDate);
        console.log(`   Elig ${eligIndex}/${eligList.length} #${eligNum}: Date=${eligFormattedDate}, Status="${status}", Clinician="${clinician}" → ${matches ? '✅ Date match' : '❌ Date mismatch'}`);
        if (!matches) {
          continue;
        }
      } else {
        console.log(`   Elig ${eligIndex}/${eligList.length} #${eligNum}: ❌ Failed to parse date from "${eligDateStr}"`);
        continue;
      }
    } else {
      // Non-logging path - just check date
      if (!DateHandler.isSameDay(claimDate, eligDate)) {
        continue;
      }
    }
    
    const eligClinician = (elig.Clinician || '').trim();
    if (eligClinician && claimClinicians.length && !claimClinicians.includes(eligClinician)) {
      if (shouldLog) {
        console.log(`      ❌ Clinician mismatch: has "${eligClinician}" but need: ${claimClinicians.join(', ')}`);
      }
      continue;
    }
    
    const serviceCategory = (elig['Service Category'] || '').trim();
    const consultationStatus = (elig['Consultation Status'] || '').trim();
    const department = (elig.Department || elig.Clinic || '').toLowerCase();
    const categoryCheck = isServiceCategoryValid(serviceCategory, consultationStatus, department);
    
    if (!categoryCheck.valid) {
      if (shouldLog) {
        console.log(`      ❌ Service category invalid: ${categoryCheck.reason}`);
      }
      continue;
    }
    
    if ((elig.Status || '').toLowerCase() !== 'eligible') {
      if (shouldLog) {
        console.log(`      ❌ Status "${elig.Status}" (must be "eligible")`);
      }
      continue;
    }
    
    if (shouldLog) {
      console.log(`      ✅ MATCH FOUND - Using this eligibility`);
    }
    
    if (usedEligibilities.has(elig['Eligibility Request Number'])) {
      if (shouldLog) {
        console.log(`      ⚠️ Note: Already used for another claim`);
      }
    } else {
      usedEligibilities.add(elig['Eligibility Request Number']);
    }
    
    return elig;
  }
  
  if (shouldLog) {
    console.log(`   ❌ No match found after checking all ${eligList.length} eligibilities (see rejection reasons above)`);
  }
  
  return null;
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
    const restrictedServices = {
      'dental': 'dental',
      'physio': 'physiotherapy',
      'diet': 'dietitian',
      'occupational': 'occupational therapy',
      'speech': 'speech therapy'
    };
    const foundService = Object.keys(restrictedServices).find(term => pkg.includes(term));
    if (foundService) {
      const serviceList = Object.values(restrictedServices).join(', ');
      return { valid: false, reason: `Elective consultations cannot include restricted services (${serviceList}). Package contains: "${pkgRaw}"` };
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

/* ===========================
   Report normalization & validation
   =========================== */

/**
 * Detect report type from raw parsed data (before normalization)
 * @param {Array|Object} rawData - Raw data from Excel/CSV parser
 * @returns {string} - 'Insta', 'Odoo', 'Combined', or 'Generic'
 */
function detectReportType(rawData) {
  if (!rawData) return 'Generic';
  
  // If it's array-of-arrays, convert to objects first to check headers
  if (Array.isArray(rawData) && rawData.length > 0 && Array.isArray(rawData[0])) {
    const detection = findHeaderRowFromArrays(rawData, 50);
    if (detection.rows && detection.rows.length > 0) {
      const sample = detection.rows[0];
      // Check for Combined report first (has Pri. Claim No + Visit Id or Total Amount)
      if (sample.hasOwnProperty('Pri. Claim No') && (sample.hasOwnProperty('Visit Id') || sample.hasOwnProperty('Total Amount'))) return 'Combined';
      if (sample.hasOwnProperty('Pri. Claim No')) return 'Insta';
      if (sample.hasOwnProperty('Pri. Claim ID')) return 'Odoo';
    }
    return 'Generic';
  }
  
  // If it's already an array of objects
  if (Array.isArray(rawData) && rawData.length > 0 && typeof rawData[0] === 'object' && !Array.isArray(rawData[0])) {
    const sample = rawData[0];
    // Check for Combined report first (has Pri. Claim No + Visit Id or Total Amount)
    if (sample.hasOwnProperty('Pri. Claim No') && (sample.hasOwnProperty('Visit Id') || sample.hasOwnProperty('Total Amount'))) return 'Combined';
    if (sample.hasOwnProperty('Pri. Claim No')) return 'Insta';
    if (sample.hasOwnProperty('Pri. Claim ID')) return 'Odoo';
    return 'Generic';
  }
  
  // If it has a {headers, rows} shape
  if (rawData.rows && Array.isArray(rawData.rows) && rawData.rows.length > 0) {
    const sample = rawData.rows[0];
    // Check for Combined report first (has Pri. Claim No + Visit Id or Total Amount)
    if (sample.hasOwnProperty('Pri. Claim No') && (sample.hasOwnProperty('Visit Id') || sample.hasOwnProperty('Total Amount'))) return 'Combined';
    if (sample.hasOwnProperty('Pri. Claim No')) return 'Insta';
    if (sample.hasOwnProperty('Pri. Claim ID')) return 'Odoo';
    return 'Generic';
  }
  
  return 'Generic';
}

function normalizeReportData(rawData) {
  if (!rawData) return [];

  // If the input is an array-of-arrays (what XLSX.utils.sheet_to_json(..., {header:1}) returns),
  // convert it into a { headers, rows } shape using the helper so downstream mapping can work.
  if (Array.isArray(rawData) && rawData.length > 0 && Array.isArray(rawData[0])) {
    const detection = findHeaderRowFromArrays(rawData, 50);
    // detection.headers is an array of header strings, detection.rows is array-of-objects keyed by headers
    rawData = {
      headers: detection.headers,
      rows: detection.rows
    };
  }

  // If rawData is an array of plain objects (not the {headers, rows} shape), handle that too.
  if (Array.isArray(rawData) && rawData.length > 0 && !rawData.headers && typeof rawData[0] === 'object' && !Array.isArray(rawData[0])) {
    const sample = rawData[0];
    const isCombined = sample.hasOwnProperty('Pri. Claim No') && (sample.hasOwnProperty('Visit Id') || sample.hasOwnProperty('Total Amount'));
    const isInsta = sample.hasOwnProperty('Pri. Claim No') && !isCombined;
    const isOdoo = sample.hasOwnProperty('Pri. Claim ID');
    return rawData.map(row => {
      if (isCombined) {
        return {
          claimID: row['Pri. Claim No'] || '',
          memberID: row['Pri. Patient Insurance Card No'] || '',
          claimDate: row['Encounter Date'] || '',
          clinician: row['Clinician License'] || '',
          department: row['Department'] || '',
          packageName: row['Pri. Plan Type'] || '',
          insuranceCompany: row['Pri. Plan Type'] || '',
          claimStatus: '',
          fileNo: row['Patient Code'] || '',
          admittingDoctor: ''
        };
      } else if (isInsta) {
        return {
          claimID: row['Pri. Claim No'] || '',
          memberID: row['Pri. Patient Insurance Card No'] || '',
          claimDate: row['Encounter Date'] || '',
          clinician: row['Clinician License'] || '',
          department: row['Department'] || '',
          packageName: row['Pri. Payer Name'] || '',
          insuranceCompany: row['Pri. Payer Name'] || '',
          claimStatus: row['Codification Status'] || '',
          fileNo: row['Patient Code'] || '',
          admittingDoctor: ''  // Insta reports don't have a separate Admitting Doctor column
        };
      } else if (isOdoo) {
        return {
          claimID: row['Pri. Claim ID'] || '',
          memberID: row['Pri. Member ID'] || '',
          claimDate: row['Adm/Reg. Date'] || '',
          clinician: row['Admitting License'] || '',
          department: row['Admitting Department'] || '',
          insuranceCompany: row['Pri. Plan Type'] || '',
          claimStatus: row['Codification Status'] || '',
          fileNo: row['MR No'] || '',
          admittingDoctor: row['Admitting Doctor'] || ''
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
          claimStatus: row['VisitStatus'] || '',
          fileNo: row['MR No'] || row['Patient Code'] || row['File No'] || row['FileNo'] || '',
          admittingDoctor: row['Admitting Doctor'] || row['Doctor'] || row['Physician'] || ''
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
    const isCombined = !!(r['Pri. Claim No'] && (r['Visit Id'] || r['Total Amount']));
    const isInsta = !!(r['Pri. Claim No']) && !isCombined;
    const isOdoo = !!r['Pri. Claim ID'];

    if (isCombined) {
      return {
        claimID: r['Pri. Claim No'] || '',
        memberID: r['Pri. Patient Insurance Card No'] || '',
        claimDate: r['Encounter Date'] || '',
        clinician: r['Clinician License'] || '',
        department: r['Department'] || '',
        packageName: r['Pri. Plan Type'] || '',
        insuranceCompany: r['Pri. Plan Type'] || '',
        claimStatus: '',
        fileNo: r['Patient Code'] || '',
        admittingDoctor: ''
      };
    } else if (isInsta) {
      return {
        claimID: r['Pri. Claim No'] || '',
        memberID: r['Pri. Patient Insurance Card No'] || '',
        claimDate: r['Encounter Date'] || '',
        clinician: r['Clinician License'] || '',
        department: r['Department'] || '',
        packageName: r['Pri. Plan Name'] || '',
        insuranceCompany: r['Pri. Plan Name'] || '',
        claimStatus: r['Codification Status'] || '',
        fileNo: r['Patient Code'] || '',
        admittingDoctor: ''  // Insta reports don't have a separate Admitting Doctor column
      };
    } else if (isOdoo) {
      return {
        claimID: r['Pri. Claim ID'] || '',
        memberID: r['Pri. Member ID'] || '',
        claimDate: r['Adm/Reg. Date'] || '',
        clinician: r['Admitting License'] || '',
        department: r['Admitting Department'] || '',
        insuranceCompany: r['Pri. Sponsor'] || '',
        claimStatus: r['Codification Status'] || '',
        fileNo: r['MR No'] || '',
        admittingDoctor: r['Admitting Doctor'] || ''
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
        claimStatus: r['Codification Status'] || r['VisitStatus'] || r['Status'] || getField(r, ['Codification Status','VisitStatus','Status','Claim Status']) || '',
        fileNo: r['MR No'] || r['Patient Code'] || getField(r, ['MR No','Patient Code','File No','FileNo']) || '',
        admittingDoctor: r['Admitting Doctor'] || getField(r, ['Admitting Doctor','Doctor','Physician']) || ''
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
  const seenClaimIDs = new Set(); // Track claim IDs to avoid duplicates
  let claimIndex = 0; // Track claim index for detailed logging (first 3 non-duplicate claims)
  
  // Log that validation started
  console.log(`\n📊 Starting validation with ${reportDataArray.length} rows, reportType="${reportType}"`);
  
  for (let i = 0; i < reportDataArray.length; i++) {
    const row = reportDataArray[i];
    const claimID = String(row.claimID || '').trim();
    if (!claimID) {
      if (i < 3) console.log(`  Row ${i}: Skipped - no claimID`);
      continue;
    }
    
    // Skip duplicate claim IDs - keep only first occurrence
    if (seenClaimIDs.has(claimID)) {
      if (i < 3) console.log(`  Row ${i}: Skipped - duplicate claimID ${claimID}`);
      continue;
    }
    seenClaimIDs.add(claimID);
    claimIndex++; // Increment for every non-duplicate claim

    const rawMemberID = String(row.memberID || '').trim();
    if (!rawMemberID) {
      if (claimIndex <= 3) console.log(`  Claim #${claimIndex} (${claimID}): Skipped - no memberID`);
      continue;
    }
    const memberID = normalizeMemberID(rawMemberID);

    let insurance = (row.insuranceCompany || '').trim();
    
    // For Insta CSV reports, dates are in DD/MM/YYYY format, not MM/DD/YYYY
    // So we should NOT use preferMDY even for CSV files
    let claimDate = DateHandler.parse(row.claimDate, { preferMDY: false });
    
    // Apply swap ONLY for Insta reports (claim dates are backwards in Insta CSV)
    // Eligibility dates are NOT swapped (they are already correct)
    let wasSwapped = false;
    if (reportType === 'Insta' && claimDate) {
      const originalDate = new Date(claimDate);
      const swappedDate = DateHandler._swapDayMonth(claimDate);
      if (swappedDate.getTime() !== originalDate.getTime()) {
        claimDate = swappedDate;
        wasSwapped = true;
      }
    }
    
    if (!claimDate) {
      if (claimIndex <= 3) console.log(`  Claim #${claimIndex} (${claimID}): Skipped - failed to parse date from "${row.claimDate}"`);
      continue;
    }
    const formattedDate = DateHandler.format(claimDate);

    if (memberID.startsWith('(VVIP)')) {
      if (claimIndex <= 3) console.log(`  Claim #${claimIndex} (${claimID}): VVIP member, skipping eligibility check`);
      results.push({ 
        claimID, memberID, encounterStart: formattedDate, 
        status: 'VVIP', finalStatus: 'valid', 
        remarks: ['VVIP member, eligibility check bypassed'], 
        fullEligibilityRecord: null,
        fileNo: row.fileNo || '',
        admittingDoctor: row.admittingDoctor || ''
      });
      continue;
    }

    // Check for leading zero in original memberID
    const hasLeadingZero = rawMemberID.match(/^0+\d+$/);
    
    if (claimIndex <= 3) {
      console.log(`  Claim #${claimIndex} (${claimID}): Processing - Member ${memberID}, Date ${formattedDate}`);
    }
    
    // Pass claimIndex to enable logging for first 3 claims
    const eligibility = findEligibilityForClaim(eligMap, claimDate, memberID, [row.clinician], insurance, false, claimIndex);
    let finalStatus = 'invalid', remarks = [];
    
    if (hasLeadingZero) {
      finalStatus = 'invalid';
      remarks.push('Member ID has a leading zero; claim marked as invalid.');
    } else if (!eligibility) {
      remarks.push(`No matching eligibility found for ${memberID} on ${formattedDate}`);
    } else if (eligibility.Status?.toLowerCase() === 'eligible') {
      const categoryCheck = isServiceCategoryValid(eligibility['Service Category'], eligibility['Consultation Status'], (row.department || '').toLowerCase());
      if (categoryCheck.valid) {
        // Validate package name match if both claim and eligibility have package names
        if (row.packageName && eligibility['Package Name']) {
          // Use special matching logic that handles Thiqa/TC packages
          if (!packageNamesMatch(row.packageName, eligibility['Package Name'])) {
            finalStatus = 'invalid';
            remarks.push(`Package name mismatch: claim has "${row.packageName}", eligibility has "${eligibility['Package Name']}"`);
          } else {
            finalStatus = 'valid';
          }
        } else {
          // No package name to validate, so it's valid based on other checks
          finalStatus = 'valid';
        }
      } else {
        remarks.push(categoryCheck.reason || 'Service category mismatch');
      }
    } else {
      remarks.push(`Eligibility status: ${eligibility.Status}`);
    }

    results.push({
      claimID, memberID, encounterStart: formattedDate,
      packageName: eligibility?.['Package Name'] || row.packageName || '',
      provider: insurance,
      clinician: eligibility?.Clinician || row.clinician || '',
      serviceCategory: eligibility?.['Service Category'] || '',
      consultationStatus: eligibility?.['Consultation Status'] || '',
      status: eligibility?.Status || '',
      claimStatus: row.claimStatus || '',
      remarks, finalStatus, fullEligibilityRecord: eligibility,
      fileNo: row.fileNo || '',
      admittingDoctor: row.admittingDoctor || ''
    });
  }
  return results;
}

/* ===========================
   Display helpers & rendering
   =========================== */
// Track which status filters are active
const activeStatusFilters = {
  valid: true,
  invalid: true,
  unknown: true
};

function getDisplayedResultsFromStored(results) {
  const raw = results || window.lastValidationResults || [];
  
  // Filter based on active status filters
  return raw.filter(r => {
    if (!r || !r.finalStatus) return false;
    const status = r.finalStatus.toLowerCase();
    // Show the result if its status filter is active
    return activeStatusFilters[status] === true;
  });
}

function renderResults(results, eligMap, totalResults = null) {
  if (!resultsContainer) return;
  resultsContainer.innerHTML = '';

  // totalResults is used for calculating total counts, while results is what's displayed
  // If not provided, assume all results are being displayed
  const allResults = totalResults || results;
  const displayedRows = results;

  if (!displayedRows || displayedRows.length === 0) {
    resultsContainer.innerHTML = '<div class="text-muted">No claims to display</div>';
    return;
  }

  const tableContainer = document.createElement('div');
  tableContainer.className = 'table-responsive analysis-results';

  const table = document.createElement('table');
  table.className = 'table table-sm table-striped table-hover shared-table';

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

  // Calculate total counts from ALL results (not just displayed ones)
  allResults.forEach((result) => {
    if (!result.memberID || result.memberID.toString().trim() === '') return;
    const statusToCheck = (result.claimStatus || result.status || result.fullEligibilityRecord?.Status || '')
      .toString()
      .trim()
      .toLowerCase();
    if (statusToCheck === 'not seen') return;
    if (result.finalStatus && statusCounts.hasOwnProperty(result.finalStatus)) statusCounts[result.finalStatus]++;
  });

  const finalStatusToBootstrap = {
    valid: 'table-success',
    invalid: 'table-danger',
    unknown: 'table-warning'
  };

  // Render only the displayed rows
  displayedRows.forEach((result, index) => {
    if (!result.memberID || result.memberID.toString().trim() === '') return;
    const statusToCheck = (result.claimStatus || result.status || result.fullEligibilityRecord?.Status || '')
      .toString()
      .trim()
      .toLowerCase();
    if (statusToCheck === 'not seen') return;

    const row = document.createElement('tr');

    const finalStatus = (result.finalStatus || '').toString().toLowerCase();
    if (finalStatus) {
      row.classList.add(finalStatus);
      const bs = finalStatusToBootstrap[finalStatus];
      if (bs) row.classList.add(bs);
    }

    const provider = (result.provider || result.insuranceCompany || result.packageName || '').toString().toLowerCase();
    if (provider.includes('daman')) row.classList.add('daman-only');
    else if (provider.includes('thiqa')) row.classList.add('thiqa-only');

    if ((result.finalStatus || '').toLowerCase() === 'vvip' || (result.status || '').toString().toLowerCase() === 'vvip') {
      row.classList.add('selected');
    }

    const statusBadge = result.status
      ? `<span class="badge ${result.status.toString().toLowerCase() === 'eligible' ? 'bg-success' : 'bg-danger'}">${escapeHtml(result.status)}</span>`
      : '';

    const remarksHTML = result.remarks && result.remarks.length > 0
      ? result.remarks.map(r => `<div>${escapeHtml(r)}</div>`).join('')
      : '<div class="source-note">No remarks</div>';

    // Build details button html without truncation
    let detailsCellHtml = '<div class="source-note">N/A</div>';
    if (result.fullEligibilityRecord && result.fullEligibilityRecord['Eligibility Request Number']) {
      // If a full eligibility record is attached to this result, show a primary "View details" button that opens the modal with the single record
      detailsCellHtml = `<button class="btn btn-sm btn-outline-primary eligibility-details" data-index="${index}" data-claimdate="${escapeHtml(result.encounterStart)}">View details</button>`;
    } else if (eligMap && typeof eligMap.get === 'function' && (eligMap.get(normalizeMemberID(result.memberID)) || []).length) {
      // Otherwise, if there are eligibilities in the map for this member, offer a secondary button to view all eligibilities for the member
      detailsCellHtml = `<button class="btn btn-sm btn-outline-secondary show-all-eligibilities" data-member="${escapeHtml(result.memberID)}" data-claimdate="${escapeHtml(result.encounterStart)}">View eligibilities</button>`;
      detailsCellHtml += ` <button class="btn btn-sm btn-outline-info show-diagnostics" data-index="${index}" title="Show diagnostic logging for this claim">
        <i class="bi bi-terminal"></i> Diagnostics
      </button>`;
    } else {
      // Even if no eligibility found, add diagnostics button to help debug why
      detailsCellHtml = `<button class="btn btn-sm btn-outline-info show-diagnostics" data-index="${index}" title="Show diagnostic logging for this claim">
        <i class="bi bi-terminal"></i> Diagnostics
      </button>`;
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

  // Update results title to show report type with displayed count
  const resultsTitle = document.getElementById('resultsTitle');
  if (resultsTitle) {
    const reportTypeText = detectedReportType === 'Insta' ? 'Insta Results' :
                           detectedReportType === 'Odoo' ? 'Odoo Results' :
                           detectedReportType === 'Combined' ? 'Combined Results' :
                           'Results';
    resultsTitle.textContent = `${displayedRows.length} ${reportTypeText}`;
  }

  // Update filter buttons to show counts (totals, never change)
  const filterValid = document.querySelector('#filterValid .badge');
  const filterInvalid = document.querySelector('#filterInvalid .badge');
  const filterUnknown = document.querySelector('#filterUnknown .badge');
  
  if (filterValid) filterValid.textContent = `Valid (${statusCounts.valid})`;
  if (filterInvalid) filterInvalid.textContent = `Invalid (${statusCounts.invalid})`;
  if (filterUnknown) filterUnknown.textContent = `Unknown (${statusCounts.unknown})`;

  initEligibilityModal(results, lastEligMap);

  resultsContainer.setAttribute('tabindex', '-1');
  resultsContainer.focus();
}

/* ===========================
   Modal, details rendering, debug utility
   =========================== */
function initEligibilityModal(results, eligMap) {
  if (!document.getElementById("modalOverlay")) {
    const modalHtml = `
      <div id="modalOverlay" class="modal" tabindex="-1" aria-hidden="true">
        <div class="modal-dialog modal-xl modal-dialog-centered">
          <div class="modal-content">
            <div class="modal-header d-flex align-items-center">
              <h5 class="modal-title me-auto">Eligibility Details</h5>
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

    const overlay = document.getElementById("modalOverlay");
    const closeBtn = document.getElementById("modalCloseBtn");
    closeBtn.addEventListener('click', hideModal);
    overlay.addEventListener('click', function (e) { if (e.target === overlay) hideModal(); });
    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape') {
        const ov = document.getElementById('modalOverlay');
        if (ov && ov.style.display && ov.style.display !== 'none') hideModal();
      }
    });

    const debugBtn = document.getElementById('modalDebugBtn');
    debugBtn.addEventListener('click', () => {
      const ctx = window.__elig_current_debug || null;
      generateAndSendDebugLog(ctx, results, eligMap);
    });
  }

  document.querySelectorAll(".eligibility-details").forEach(btn => {
    btn.onclick = null;
    btn.addEventListener('click', function () {
      const index = parseInt(this.dataset.index, 10);
      const result = results[index];
      if (!result?.fullEligibilityRecord) return;
      const record = result.fullEligibilityRecord;
      const claimDateStr = this.dataset.claimdate || result.encounterStart || '';
      const claimDate = claimDateStr ? DateHandler.parse(claimDateStr) : null;
      window.__elig_current_debug = { mode: 'single', member: result.memberID, claimDate: claimDateStr || '', record, resultIndex: index };
      const debugBtn = document.getElementById('modalDebugBtn'); if (debugBtn) debugBtn.style.display = '';
      document.getElementById("modalTable").innerHTML = formatEligibilityDetails(record, result.memberID, claimDate);
      showModal();
    });
  });

  document.querySelectorAll(".show-all-eligibilities").forEach(btn => {
    btn.onclick = null;
    btn.addEventListener('click', function () {
      const member = this.dataset.member;
      const claimDateStr = this.dataset.claimdate || '';
      const claimDate = claimDateStr ? DateHandler.parse(claimDateStr) : null;
      const list = (typeof eligMap.get === 'function') ? (eligMap.get(normalizeMemberID(member)) || []) : [];
      const modalTable = document.getElementById("modalTable");
      window.__elig_current_debug = { mode: 'list', member, claimDate: claimDateStr || '', listSnapshot: list.slice(0,200) };
      const debugBtn = document.getElementById('modalDebugBtn'); if (debugBtn) debugBtn.style.display = '';

      if (!list.length) {
        modalTable.innerHTML = `<div class="p-3">No eligibilities found for <strong>${escapeHtml(member)}</strong></div>`;
        showModal();
        return;
      }

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
        const formattedEligDate = eligDate ? DateHandler.format(eligDate) : answeredOnRaw;
        let trClass = '';
        if (claimDate && eligDate) {
          if (DateHandler.isSameDay(claimDate, eligDate)) trClass = 'table-warning';
          else trClass = 'table-danger';
        }
        html += `<tr class="${trClass}">
          <td>${idx + 1}</td>
          <td>${escapeHtml(rec['Eligibility Request Number'] || '')}</td>
          <td>${escapeHtml(formattedEligDate || '')}</td>
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

  // Add handler for diagnostics buttons
  document.querySelectorAll(".show-diagnostics").forEach(btn => {
    btn.onclick = null;
    btn.addEventListener('click', function () {
      const index = parseInt(this.dataset.index, 10);
      const result = results[index];
      if (!result) return;
      
      console.group(`🔍 [Manual Diagnostics] Claim: ${result.claimID}, Member: ${result.memberID}`);
      console.log('Claim Details:', {
        claimID: result.claimID,
        memberID: result.memberID,
        encounterDate: result.encounterStart,
        provider: result.provider,
        packageName: result.packageName,
        clinician: result.clinician,
        department: result.serviceCategory,
        finalStatus: result.finalStatus,
        remarks: result.remarks
      });
      
      // Re-run eligibility matching with forced logging
      const claimDate = DateHandler.parse(result.encounterStart);
      if (claimDate && eligMap) {
        console.log('Re-running eligibility matching with forced diagnostics...');
        const clinicians = result.clinician ? [result.clinician] : [];
        findEligibilityForClaim(eligMap, claimDate, result.memberID, clinicians, result.provider, true);
      } else {
        console.warn('Cannot re-run matching: missing claim date or eligibility map');
      }
      
      console.groupEnd();
      alert('Diagnostic logging has been written to the console. Press F12 to view the console.');
    });
  });

  function showModal() {
    const overlay = document.getElementById("modalOverlay");
    if (!overlay) return;
    overlay.style.display = 'flex';
    overlay.setAttribute('aria-hidden', 'false');
    setTimeout(() => overlay.classList.add('show'), 10);
    const focusable = overlay.querySelectorAll('button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])');
    if (focusable.length) focusable[0].focus();
  }
}

/* Debug log generator (used by modal debug button) */
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
      lastValidationSample: (window.lastValidationResults && Array.isArray(window.lastValidationResults)) ? window.lastValidationResults.slice(0,50) : []
    };

    if (ctx && ctx.member && eligMap && typeof eligMap.get === 'function') {
      const memberKey = normalizeMemberID(ctx.member);
      const memberEntries = eligMap.get(memberKey) || [];
      payload.memberEligibilities = memberEntries.slice(0,200);
    }

    const text = JSON.stringify(payload, null, 2);

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

    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).catch(() => {/* ignore */});
    }

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

/* Modal hide helper */
function hideModal() { const overlay = document.getElementById("modalOverlay"); if (overlay) overlay.style.display = "none"; }

/* Details formatter for a single eligibility record.
   Optional claimDate param can be used to colour date rows. */
function formatEligibilityDetails(record, memberID, claimDate) {
  if (!record) return '<div>No details</div>';

  const status = (record.Status || '').toString();
  const statusClass = status.toLowerCase() === 'eligible' ? 'status-badge eligible' : 'status-badge ineligible';
  let html = `<div class="mb-2"><strong>Member:</strong> ${escapeHtml(memberID)} <span class="${statusClass}" style="margin-left:8px;">${escapeHtml(status)}</span></div>`;

  html += '<table class="eligibility-details"><tbody>';

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
      let rowClass = '';
      if (typeof raw === 'string' && (key.includes('Date') || key.toLowerCase().includes('answered') || key.toLowerCase().includes('ordered'))) {
        const parsed = DateHandler.parse(raw);
        disp = parsed ? DateHandler.format(parsed) : raw;
        if (claimDate && parsed) {
          if (DateHandler.isSameDay(claimDate, parsed)) rowClass = 'table-warning';
          else rowClass = 'table-danger';
        }
      }
      html += `<tr class="${rowClass}"><th>${escapeHtml(key)}</th><td>${escapeHtml(String(disp))}</td></tr>`;
    }
  });

  Object.keys(record).forEach(key => {
    if (used.has(key)) return;
    const raw = record[key];
    if (raw === undefined || raw === null || raw === '') return;
    let disp = raw;
    let rowClass = '';
    if (typeof raw === 'string' && (key.includes('Date') || key.toLowerCase().includes('answered') || key.toLowerCase().includes('ordered'))) {
      const parsed = DateHandler.parse(raw);
      disp = parsed ? DateHandler.format(parsed) : raw;
      if (claimDate && parsed) {
        if (DateHandler.isSameDay(claimDate, parsed)) rowClass = 'table-warning';
        else rowClass = 'table-danger';
      }
    }
    html += `<tr class="${rowClass}"><th>${escapeHtml(key)}</th><td>${escapeHtml(String(disp))}</td></tr>`;
  });

  html += '</tbody></table>';
  return html;
}

/* ===========================
   Export helpers
   =========================== */
function exportInvalidEntries(results) {
  const invalidEntries = (results || []).filter(r => r && r.finalStatus === 'invalid');
  if (!invalidEntries.length) { alert('No invalid entries to export.'); return; }
  const exportData = invalidEntries.map((entry, index) => ({
    'FILE NO.': entry.fileNo || (index + 1),
    'CLAIM ID': entry.claimID || '',
    'VISIT ID': entry.claimID || '',
    'PHY LICENSE': entry.clinician || '',
    'DATE': entry.encounterStart || '',
    'MEMBER ID': entry.memberID || '',
    'ELIGIBILITY UNDER DOCTOR': entry.admittingDoctor || '',
    'ERROR': (entry.remarks || []).join('; '),
    'Reception': ''
  }));
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(exportData);
  XLSX.utils.book_append_sheet(wb, ws, 'Invalid Claims');
  XLSX.writeFile(wb, `invalid_claims_${new Date().toISOString().slice(0,10)}.xlsx`);
}

/* ===========================
   Event handlers & flow
   =========================== */
async function handleFileUpload(event, type) {
  const file = event.target.files && event.target.files[0];
  if (!file) return;
  try {
    updateStatus(`Loading ${type} file...`);
    if (type === 'eligibility') {
      // read as array-of-arrays so prepareEligibilityMap can detect header row
      const allRows = await parseExcelFile(file);
      eligData = allRows;
      updateStatus(`Loaded ${Array.isArray(eligData) ? eligData.length : 0} eligibility rows (raw)`);
      updateProcessButtonState();
      if (eligData && (rawParsedReport || xlsData)) summarizeAndDisplayCounts();
      return;
    }
    if (type === 'report') {
      lastReportWasCSV = file.name.toLowerCase().endsWith('.csv');
      const parsed = await (file.name.toLowerCase().endsWith('.csv') ? parseCsvFile(file) : parseExcelFile(file));
      rawParsedReport = parsed;
      
      // Detect report type BEFORE normalization
      detectedReportType = detectReportType(parsed);
      
      const normalized = normalizeReportData(parsed);
      xlsData = normalized.filter(r => r && r.claimID && String(r.claimID).trim() !== '');
      if (!xlsData || xlsData.length === 0) {
        // Reset report type if no valid data
        detectedReportType = 'Generic';
      }
      updateStatus(`Loaded ${xlsData.length} report rows`);
      updateProcessButtonState();
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
    
    // Detect report type BEFORE normalization
    detectedReportType = detectReportType(parsed);
    
    const normalized = normalizeReportData(parsed);
    xlsData = normalized.filter(r => r && r.claimID && String(r.claimID).trim() !== '');
    if (xlsData.length === 0) {
      // Reset report type if no valid data
      detectedReportType = 'Generic';
    }
    updateStatus(`Loaded ${xlsData.length} rows from pasted CSV`);
    updateProcessButtonState();
    if (eligData && (rawParsedReport || xlsData)) summarizeAndDisplayCounts();
  } catch (err) {
    console.error('Error parsing pasted CSV:', err);
    updateStatus('Error parsing pasted CSV');
    alert('Failed to parse pasted CSV');
    // Reset report type on error
    detectedReportType = 'Generic';
  }
}

async function handleProcessClick() {
  try {
    if (!eligData) { updateStatus('Processing stopped: Eligibility file missing'); return; }
    if (!xlsData || !xlsData.length) { updateStatus('Processing stopped: Report file missing'); return; }

    updateStatus('Processing...');
    usedEligibilities.clear();

    const eligMap = prepareEligibilityMap(eligData);
    lastEligMap = eligMap;

    // Use the report type that was detected during file upload (before normalization)
    const reportType = detectedReportType;
    
    // Validate that we detected a valid report type
    if (reportType === 'Generic') {
      const errorMsg = `❌ Unable to detect report type from uploaded file.\n\n` +
                       `Expected original columns in file:\n` +
                       `  • Insta report must have: "Pri. Claim No"\n` +
                       `  • Odoo report must have: "Pri. Claim ID"\n` +
                       `  • Combined report must have: "Pri. Claim No" and "Visit Id" (or "Total Amount")\n\n` +
                       `These columns were not found in your file.\n` +
                       `Please verify you're uploading a valid Insta, Odoo, or Combined export.`;
      console.error(errorMsg);
      updateStatus('Error: Unknown report type - see console for details');
      alert(`Cannot process report file.\n\n` +
            `This file does not appear to be a valid Insta, Odoo, or Combined report.\n\n` +
            `Expected columns in original file:\n` +
            `  • Insta reports must have: "Pri. Claim No"\n` +
            `  • Odoo reports must have: "Pri. Claim ID"\n` +
            `  • Combined reports must have: "Pri. Claim No" and "Visit Id" (or "Total Amount")\n\n` +
            `Please check your file and try again.`);
      return; // Stop processing
    }

    // Update results header title based on report type
    const resultsTitle = document.getElementById('resultsTitle');
    if (resultsTitle) {
      if (reportType === 'Insta') {
        resultsTitle.textContent = 'Insta Report Results';
      } else if (reportType === 'Odoo') {
        resultsTitle.textContent = 'Odoo Report Results';
      } else if (reportType === 'Combined') {
        resultsTitle.textContent = 'Combined Report Results';
      } else {
        resultsTitle.textContent = 'Results';
      }
    }

    const results = validateReportClaims(xlsData, eligMap, reportType);

    let outputResults = results;
    if (filterCheckbox && filterCheckbox.checked) {
      outputResults = results.filter(r => {
        const insurance = (r.insuranceCompany || r.provider || r.packageName || '').toString().toLowerCase();
        return insurance.includes('daman') || insurance.includes('thiqa');
      });
    }

    window.lastValidationResults = outputResults;
    const displayedResults = getDisplayedResultsFromStored(outputResults);
    renderResults(displayedResults, eligMap, outputResults);
    updateStatus(`Processed ${outputResults.length} claims successfully`);
  } catch (err) {
    console.error('Processing stopped due to error:', err);
  }
}

function updateProcessButtonState() {
  const hasEligibility = Array.isArray(eligData) && eligData.length > 0;
  const hasReport = Array.isArray(xlsData) && xlsData.length > 0;
  if (processBtn) processBtn.disabled = !(hasEligibility && hasReport);
  if (exportInvalidBtn) exportInvalidBtn.disabled = !(hasEligibility && hasReport);
}

function updateStatus(msg) { if (statusEl) statusEl.textContent = msg || 'Ready'; }

function onFilterToggle() {
  if (!filterStatus) return;
  const on = filterCheckbox && filterCheckbox.checked;
  filterStatus.textContent = on ? 'ON' : 'OFF';
  filterStatus.classList.toggle('active', on);
  if (!window.lastValidationResults) return;

  let base = window.lastValidationResults.slice();
  if (on) {
    base = base.filter(r => {
      const provider = (r.provider || r.insuranceCompany || r.packageName || r['Payer Name'] || r['Insurance Company'] || '').toString().toLowerCase();
      return provider.includes('daman') || provider.includes('thiqa');
    });
  }

  const displayed = getDisplayedResultsFromStored(base);
  const eligMap = lastEligMap || (eligData ? prepareEligibilityMap(eligData) : new Map());
  renderResults(displayed, eligMap, base);
}

function onRemoveZeroesToggle() {
  if (!removeZeroesStatus) return;
  const on = removeZeroesCheckbox && removeZeroesCheckbox.checked;
  removeZeroesStatus.textContent = on ? 'ON' : 'OFF';
  removeZeroesStatus.classList.toggle('active', on);
  
  // Update the global state
  removeLeadingZeroes = on;
  
  // If we have data, we need to rebuild the eligibility map and re-process
  if (eligData && xlsData) {
    updateStatus('Leading zeroes option changed. Click Process to re-check with new settings.');
    console.log(`🔧 Leading zeroes removal ${on ? 'ENABLED' : 'DISABLED'}. Re-processing required.`);
  }
}

/* ===========================
   Initialization
   =========================== */
function initializeEventListeners() {
  reportInput = document.getElementById('reportFileInput');
  eligInput = document.getElementById('eligibilityFileInput');
  processBtn = document.getElementById('processBtn');
  exportInvalidBtn = document.getElementById('exportInvalidBtn');
  statusEl = document.getElementById('uploadStatus');
  resultsContainer = document.getElementById('results');
  filterCheckbox = document.getElementById('filterDamanThiqa');
  filterStatus = document.getElementById('filterStatus');
  removeZeroesCheckbox = document.getElementById('removeLeadingZeroes');
  removeZeroesStatus = document.getElementById('removeZeroesStatus');
  pasteTextarea = document.getElementById('pasteCsvTextarea');
  pasteBtn = document.getElementById('pasteCsvBtn');

  if (eligInput) eligInput.addEventListener('change', (e) => handleFileUpload(e, 'eligibility'));
  if (reportInput) reportInput.addEventListener('change', (e) => handleFileUpload(e, 'report'));
  if (processBtn) processBtn.addEventListener('click', handleProcessClick);
  if (exportInvalidBtn) exportInvalidBtn.addEventListener('click', () => exportInvalidEntries(window.lastValidationResults || []));
  if (filterCheckbox) {
    filterCheckbox.checked = true;
    filterCheckbox.addEventListener('change', onFilterToggle);
  }
  if (removeZeroesCheckbox) {
    // Default to false (unchecked) to preserve leading zeroes by default
    // This maintains backward compatibility and avoids unexpected behavior changes
    removeZeroesCheckbox.checked = false;
    removeZeroesCheckbox.addEventListener('change', onRemoveZeroesToggle);
  }

  // Setup status filter buttons
  const filterValid = document.getElementById('filterValid');
  const filterInvalid = document.getElementById('filterInvalid');
  const filterUnknown = document.getElementById('filterUnknown');

  const setupStatusFilter = (btn, status) => {
    if (!btn) return;
    btn.addEventListener('click', () => {
      // Toggle the filter state
      activeStatusFilters[status] = !activeStatusFilters[status];
      
      // Update button appearance
      btn.classList.toggle('active', activeStatusFilters[status]);
      
      // Re-render results with updated filters
      if (!window.lastValidationResults) return;
      const base = window.lastValidationResults.slice();
      let preFiltered = base;
      if (filterCheckbox && filterCheckbox.checked) {
        preFiltered = base.filter(r => {
          const provider = (r.provider || r.insuranceCompany || r.packageName || '').toString().toLowerCase();
          return provider.includes('daman') || provider.includes('thiqa');
        });
      }
      const displayed = getDisplayedResultsFromStored(preFiltered);
      const eligMap = lastEligMap || (eligData ? prepareEligibilityMap(eligData) : new Map());
      renderResults(displayed, eligMap, preFiltered);
    });
  };

  setupStatusFilter(filterValid, 'valid');
  setupStatusFilter(filterInvalid, 'invalid');
  setupStatusFilter(filterUnknown, 'unknown');

  if (pasteBtn) pasteBtn.addEventListener('click', handlePasteCsvClick);
  if (filterStatus) onFilterToggle();
  if (removeZeroesStatus) onRemoveZeroesToggle();
}

document.addEventListener('DOMContentLoaded', () => {
  initializeEventListeners();
  updateStatus('Ready to process files');
});
