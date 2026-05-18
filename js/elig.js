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
// Service category validation rules
// Each category can have:
// - keywords: array of required keywords (department must contain at least one)
// - statusRules: optional object with rules based on consultation status
//   - Each status can have: 
//     - allowedDepartments: array of department terms (substring matching)
//     - wordBoundaryTerms: array of terms to match as whole words (regex boundary matching)
const SERVICE_PACKAGE_RULES = {
  'Dental Services': {
    keywords: ['dental', 'orthodontic']
  },
  'Physiotherapy': {
    keywords: ['physio']
  },
  'Other OP Services': {
    keywords: ['physio', 'diet', 'occupational', 'speech', 'orthop', 'family']
  },
  'Consultation': {
    keywords: [],  // No general keyword requirements
    statusRules: {
      'elective': {
        // ENT/Otolaryngology are explicitly allowed for elective consultations
        allowedDepartments: ['otolaryngology', 'ear nose throat', 'ear, nose & throat', 'ear, nose and throat'],
        wordBoundaryTerms: ['ent']  // Match 'ent' only as a whole word to avoid false matches
      }
    }
  }
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
let sourceFileName = '';   // Track the source report filename for export

// Keep last eligibility map so UI filters can re-render without rebuilding the map
let lastEligMap = null;

// Option to show/hide diagnostics buttons
let showDiagnosticsButtons = false;

// DOM Elements (lookups performed in initializeEventListeners)
let reportInput, eligInput, processBtn, exportInvalidBtn, statusEl, resultsContainer, filterCheckbox, filterStatus, pasteTextarea, pasteBtn;
let diagnosticsCheckbox, diagnosticsStatus;

/* ===========================
   Small Utilities
   =========================== */
function escapeHtml(s) {
  if (s === null || s === undefined) return '';
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#x27;');
}

function escapeRegex(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function normalizeMemberID(id) {
  if (!id) return "";
  let normalized = String(id).replace(/\D/g, "").trim();
  
  // Remove leading zeroes; if input is all zeroes (e.g., "0000"), keep at least one "0"
  if (normalized.length > 0) {
    normalized = normalized.replace(/^0+/, '') || '0';
  }
  
  return normalized;
}

function normalizeClinician(name) {
  if (!name) return '';
  return name.trim().toLowerCase().replace(/\s+/g, ' ');
}

// Package matching configuration - loaded from external JSON
let eligibilityMatchingConfig = null;

/**
 * Load eligibility matching rules from external JSON configuration
 * Falls back to hardcoded defaults if loading fails
 */
async function loadEligibilityMatchingConfig() {
  try {
    const response = await fetch('resources/eligibility_matching_rules.json');
    if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
    eligibilityMatchingConfig = await response.json();
    console.log('✅ Loaded eligibility matching configuration from JSON');
  } catch (error) {
    console.warn('⚠️ Failed to load eligibility matching config, using defaults:', error.message);
    // Fallback to hardcoded configuration
    eligibilityMatchingConfig = {
      classifications: {
        daman: ['silver', 'gold', 'bronze']
      },
      matchingRules: [
        { id: 'thiqa-tc', claimContains: ['thiqa'], eligibilityContains: ['tc'] },
        { id: 'thiqa-thiqa', claimContains: ['thiqa'], eligibilityContains: ['thiqa'] },
        { id: 'daman-basic', claimContainsAll: ['daman', 'basic'], eligibilityContainsAny: ['basic', 'abu dhabi'], displayTier: 'Daman Basic' },
        { id: 'daman-enhanced-forward', claimContainsAll: ['daman', 'enhanced'], eligibilityContainsAny: ['sahtak', 'enhanced', 'core-auh'], eligibilityContainsClassification: 'daman', displayTier: 'Daman Enhanced' },
        { id: 'daman-enhanced-reverse', eligibilityContainsAll: ['daman', 'enhanced'], claimContainsAny: ['sahtak', 'dge', 'core-auh'], claimContainsClassification: 'daman', claimSpecialCondition: { type: 'enhanced-without-daman' }, displayTier: 'Daman Enhanced' },
        { id: 'daman-mid', claimContainsAll: ['daman', 'mid'], eligibilityContainsAny: ['enhanced'], eligibilityContainsClassification: 'daman', displayTier: 'Daman Enhanced' },
        { id: 'daman-high-end', claimContainsAll: ['daman', 'high-end'], eligibilityContainsAny: ['enhanced'], eligibilityContainsClassification: 'daman', displayTier: 'Daman Enhanced' },
        { id: 'daman-key', claimContainsAll: ['daman', 'key'], eligibilityContainsClassification: 'daman', displayTier: 'Daman Key' },
        { id: 'daman-low-end', claimContainsAll: ['daman', 'low-end'], eligibilityContainsAny: ['enhanced-auh', 'core-auh', 'abu dhabi', 'nw uae', 'etihad nw', 'adnoc'], eligibilityContainsClassification: 'daman', displayTier: 'Daman Enhanced' },
        { id: 'daman-generic', claimPattern: '/(^|[_\\-])daman/i', claimExcludes: ['basic', 'enhanced', 'low-end', 'high-end'], eligibilityContainsClassification: 'daman' }
      ]
    };
  }
}

/**
 * Check if a string contains any classification keyword from the specified classification group
 * @param {string} text - Lowercase text to check
 * @param {string} classificationType - Type of classification (e.g., 'daman')
 * @returns {boolean} - True if contains a classification keyword
 */
function containsClassification(text, classificationType) {
  if (!eligibilityMatchingConfig || !eligibilityMatchingConfig.classifications) return false;
  const classifications = eligibilityMatchingConfig.classifications[classificationType] || [];
  return classifications.some(classification => text.includes(classification));
}

/**
 * Check if package names match using rules from external JSON configuration
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
  
  // If config not loaded, return false (will be loaded during initialization)
  if (!eligibilityMatchingConfig || !eligibilityMatchingConfig.matchingRules) {
    console.warn('⚠️ Eligibility matching config not loaded yet');
    return false;
  }
  
  // Apply each matching rule from configuration
  for (const rule of eligibilityMatchingConfig.matchingRules) {
    if (matchesRule(rule, claimLower, eligLower, claimPackage)) {
      return true;
    }
  }
  
  return false;
}

/**
 * Check if a specific rule matches the claim and eligibility packages
 * @param {Object} rule - Matching rule from configuration
 * @param {string} claimLower - Lowercase claim package name
 * @param {string} eligLower - Lowercase eligibility package name
 * @param {string} claimPackage - Original claim package name (for pattern matching)
 * @returns {boolean} - True if rule matches
 */
function matchesRule(rule, claimLower, eligLower, claimPackage) {
  // Check claim-side conditions
  if (rule.claimContains) {
    if (!rule.claimContains.some(term => claimLower.includes(term))) return false;
  }
  
  if (rule.claimContainsAll) {
    if (!rule.claimContainsAll.every(term => claimLower.includes(term))) return false;
  }
  
  if (rule.claimContainsAny) {
    if (!rule.claimContainsAny.some(term => claimLower.includes(term))) {
      // Check if classification matches instead
      if (!rule.claimContainsClassification || !containsClassification(claimLower, rule.claimContainsClassification)) {
        return false;
      }
    }
  }
  
  if (rule.claimExcludes) {
    if (rule.claimExcludes.some(term => claimLower.includes(term))) return false;
  }
  
  if (rule.claimPattern) {
    // Reconstruct regex from string (handle pattern like "/(^|[_\-])daman/i")
    const patternMatch = rule.claimPattern.match(/^\/(.+)\/([gimuy]*)$/);
    if (patternMatch) {
      const pattern = new RegExp(patternMatch[1], patternMatch[2]);
      if (!pattern.test(claimPackage)) return false;
    }
  }
  
  // Handle special claim conditions
  if (rule.claimSpecialCondition) {
    if (rule.claimSpecialCondition.type === 'enhanced-without-daman') {
      // Match 'enhanced' in claim only if 'daman' is NOT present
      const hasEnhanced = claimLower.includes('enhanced');
      const hasDaman = claimLower.includes('daman');
      if (hasEnhanced && hasDaman) return false;
      if (!hasEnhanced && !rule.claimContainsAny.some(term => claimLower.includes(term)) && !containsClassification(claimLower, rule.claimContainsClassification)) {
        return false;
      }
    }
  }
  
  // Check eligibility-side conditions
  if (rule.eligibilityContains) {
    if (!rule.eligibilityContains.some(term => eligLower.includes(term))) return false;
  }
  
  if (rule.eligibilityContainsAll) {
    if (!rule.eligibilityContainsAll.every(term => eligLower.includes(term))) return false;
  }
  
  if (rule.eligibilityContainsAny) {
    if (!rule.eligibilityContainsAny.some(term => eligLower.includes(term))) {
      // Check if classification matches instead
      if (!rule.eligibilityContainsClassification || !containsClassification(eligLower, rule.eligibilityContainsClassification)) {
        return false;
      }
    }
  }
  
  if (rule.eligibilityExcludes) {
    if (rule.eligibilityExcludes.some(term => eligLower.includes(term))) return false;
  }
  
  // If we reach here, all conditions passed
  return true;
}

/**
 * Normalize package name to its tier level for display purposes
 * Converts specific plan names (e.g., "Sahtak", "Core-AUH-LG") to generic tier names (e.g., "Daman Enhanced")
 * Uses the JSON configuration rules to determine tier mappings
 * @param {string} packageName - Original package name
 * @returns {string} - Normalized tier name or original name if no match
 */
function normalizePackageNameForDisplay(packageName) {
  if (!packageName) return packageName;
  
  const packageLower = packageName.trim().toLowerCase();
  
  // THIQA packages - keep as is (not in DAMAN tier system)
  if (packageLower.includes('thiqa') || packageLower.includes('tc')) {
    return packageName;
  }
  
  // If config not loaded, fall back to checking for explicit "daman" in package name
  if (!eligibilityMatchingConfig || !eligibilityMatchingConfig.matchingRules) {
    if (packageLower.includes('daman')) {
      return packageName; // Keep original if we can't determine tier
    }
    return packageName;
  }
  
  // Check if package name explicitly contains "daman" with a tier indicator
  if (packageLower.includes('daman')) {
    // Check for specific tier indicators in the package name
    if (packageLower.includes('basic')) {
      return 'Daman Basic';
    }
    if (packageLower.includes('enhanced')) {
      return 'Daman Enhanced';
    }
    if (packageLower.includes('high-end')) {
      return 'Daman Enhanced';
    }
    if (packageLower.includes('low-end')) {
      return 'Daman Enhanced';
    }
    if (packageLower.includes('mid')) {
      return 'Daman Enhanced';
    }
    if (packageLower.includes('key')) {
      return 'Daman Key';
    }
  }
  
  // For eligibility packages without explicit "daman", use JSON rules to determine tier
  // Iterate through rules to find matches and derive tier from the claim package pattern
  for (const rule of eligibilityMatchingConfig.matchingRules) {
    // Skip non-DAMAN rules
    if (!rule.id.startsWith('daman-')) {
      continue;
    }
    
    // Extract tier name from rule using displayTier field in JSON
    const tierName = extractTierNameFromRuleId(rule.id, rule);
    if (!tierName) continue;
    
    // Check if this package matches the eligibility patterns in this rule
    if (packageMatchesEligibilityPattern(packageLower, rule)) {
      return tierName;
    }
  }
  
  // If no specific tier identified, return original name
  return packageName;
}

/**
 * Extract display tier name from rule ID
 * @param {string} ruleId - Rule identifier (e.g., "daman-basic", "daman-enhanced-forward")
 * @param {Object} rule - The rule object from configuration
 * @returns {string|null} - Tier name (e.g., "Daman Basic") from rule.displayTier, or null if not available
 */
function extractTierNameFromRuleId(ruleId, rule) {
  // Use displayTier from JSON if available
  if (rule && rule.displayTier) {
    return rule.displayTier;
  }
  
  // Fallback: No tier name available (e.g., for generic or non-display rules)
  return null;
}

/**
 * Check if a package name matches the eligibility patterns in a rule
 * @param {string} packageLower - Lowercase package name
 * @param {Object} rule - Matching rule from configuration
 * @returns {boolean} - True if package matches eligibility pattern
 */
function packageMatchesEligibilityPattern(packageLower, rule) {
  // Check eligibilityContainsAny
  if (rule.eligibilityContainsAny) {
    if (rule.eligibilityContainsAny.some(term => packageLower.includes(term))) {
      return true;
    }
  }
  
  // Check eligibilityContainsAll
  if (rule.eligibilityContainsAll) {
    if (rule.eligibilityContainsAll.every(term => packageLower.includes(term))) {
      return true;
    }
  }
  
  // Check eligibilityContains (single check)
  if (rule.eligibilityContains) {
    if (rule.eligibilityContains.some(term => packageLower.includes(term))) {
      return true;
    }
  }
  
  // Check classification match
  if (rule.eligibilityContainsClassification) {
    if (containsClassification(packageLower, rule.eligibilityContainsClassification)) {
      return true;
    }
  }
  
  // Also check claimContainsAny for reverse rules (where claim patterns are in the package)
  if (rule.claimContainsAny && rule.eligibilityContainsAll) {
    // This handles reverse mappings where the package might have claim-side terms
    if (rule.claimContainsAny.some(term => packageLower.includes(term))) {
      return true;
    }
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

  formatForExport: function(date) {
    if (!(date instanceof Date) || isNaN(date)) return '';
    const day = date.getUTCDate();
    const month = date.getUTCMonth() + 1; // Month is 0-indexed, so add 1
    const year = date.getUTCFullYear();
    return `${month}/${day}/${year}`;
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
    'pri. patient insurance card no', 'institution', 'facility id', 'mr no.', 'pri. claim id',
    'clinician license', 'encounter date', 'visit id', 'total amount', 'source file',
    'emirates id no', 'emirates id no.'
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
    console.log(`📋 Total columns: ${headers.length}`);
    console.log(`📋 Column headers (first 15):`, headers.slice(0, 15).join(', '));
    if (headers.length > 15) {
      console.log(`📋 Column headers (columns 16-30):`, headers.slice(15, 30).join(', '));
    }
    if (headers.length > 30) {
      console.log(`📋 ALL column headers:`, headers.join(', '));
    }
    
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
    
    // Log which ID columns are available in the file
    const idCandidates = [
      'Card Number / DHA Member ID', 'Card Number', 'MemberID', 'Member ID',
      'Patient Insurance Card No', 'Policy1', 'Policy 1', 'PatientCardID'
    ];
    const availableIdColumns = idCandidates.filter(col => headers.includes(col));
    console.log(`\n🔍 ID Column Analysis:`);
    console.log(`   Candidates searched: ${idCandidates.join(', ')}`);
    console.log(`   Available in file: ${availableIdColumns.length > 0 ? availableIdColumns.join(', ') : 'NONE'}`);
    if (availableIdColumns.length === 0) {
      console.log(`   ⚠️ WARNING: No standard ID columns found! Will try first available from candidates.`);
      console.log(`   All columns in file: ${headers.join(', ')}`);
    }
    console.log(``);

    for (let i = headerRowIndex + 1; i < rawSheetArray.length; i++) {
      const row = rawSheetArray[i];
      if (!Array.isArray(row)) continue;

      const record = {};
      // For duplicate headers (e.g. "Package Name" repeating across policy blocks),
      // the first non-empty occurrence takes precedence over later ones.
      headers.forEach((h, idx) => {
        const val = row[idx] !== undefined ? row[idx] : '';
        if (!Object.prototype.hasOwnProperty.call(record, h) || record[h] === '' || record[h] === null || record[h] === undefined) {
          record[h] = val;
        }
      });

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
 * Find matching eligibility records for a claim
 * @param {Map} eligMap - Map of member IDs to eligibility records
 * @param {Date} claimDate - Claim date
 * @param {string} memberID - Member ID
 * @param {Array} claimClinicians - Array of clinician names/IDs
 * @param {string} provider - Provider/insurance name (used to filter diagnostic logging to Daman/Thiqa only)
 * @param {boolean} forceLog - Force diagnostic logging regardless of provider
 * @param {number} claimIndex - Index of the claim (for logging first 3)
 * @returns {Array} - Array of matching eligibility records (empty array if none found)
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
    return [];
  }
  
  if (shouldLog) {
    const clinInfo = claimClinicians.length > 0 ? `, Clinicians: ${claimClinicians.join(', ')}` : '';
    console.log(`⚠️ CLAIM #${claimIndex}: Found ${eligList.length} eligibilities for member ${memberID}, Date: ${DateHandler.format(claimDate)}${clinInfo}`);
  }
  
  // Collect ALL matching eligibilities instead of returning just the first one
  const matchingEligs = [];
  let eligIndex = 0;
  
  for (const elig of eligList) {
    eligIndex++;
    const eligDate = DateHandler.parse(elig["Answered On"] || elig["Ordered On"], { preferMDY: false });
    
    if (shouldLog) {
      const eligNum = elig["Eligibility Request Number"] || "(unknown)";
      const eligDateStr = elig["Answered On"] || elig["Ordered On"] || '(empty)';
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
    
    const eligStatus = (elig.Status || '').toLowerCase();
    const isBlockedOrCancelled = eligStatus === 'blocked' || eligStatus === 'cancelled';

    const eligClinician = (elig.Clinician || '').trim();
    const hasClaimClinician = claimClinicians.some(c => c && c.trim());
    const clinicianMismatch = eligClinician && hasClaimClinician && !claimClinicians.includes(eligClinician);

    // Eligible eligibilities require a clinician match.
    // Blocked/Cancelled eligibilities are included even when the clinician doesn't match —
    // they take priority over Eligible eligibilities when both exist for the same date
    // (the scenario where the wrong clinician was entered on the claim but the
    // Blocked/Cancelled eligibility is the correct one for that visit).
    if (clinicianMismatch && !isBlockedOrCancelled) {
      if (shouldLog) {
        console.log(`      ❌ Clinician mismatch: has "${eligClinician}" but need: ${claimClinicians.join(', ')}`);
      }
      continue;
    }

    if (clinicianMismatch && shouldLog) {
      console.log(`      ⚠️ Clinician mismatch for Blocked/Cancelled eligibility (including anyway — takes priority): has "${eligClinician}" but need: ${claimClinicians.join(', ')}`);
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
    
    // Include Eligible, Blocked, and Cancelled statuses; skip all others
    if (eligStatus !== 'eligible' && !isBlockedOrCancelled) {
      if (shouldLog) {
        console.log(`      ❌ Status "${elig.Status}" (must be eligible, blocked, or cancelled)`);
      }
      continue;
    }
    
    if (shouldLog) {
      console.log(`      ✅ MATCH FOUND - Adding to results (match ${matchingEligs.length + 1})`);
    }
    
    // Check if already used, but don't mark as used yet - will mark the selected one later
    const isUsed = usedEligibilities.has(elig['Eligibility Request Number']);
    if (shouldLog && isUsed) {
      console.log(`      ⚠️ Note: Already used for another claim`);
    }
    
    // Add this eligibility to matches with usage status and clinician-match flags
    // _clinicianMatch: true  = elig has a non-empty clinician AND it positively matched the claim's clinician
    // _clinicianMatch: false = not a positive clinician match
    //                         NOTE: when _clinicianMismatch is true, _clinicianMatch is always false
    const clinicianPositivelyMatched = eligClinician && hasClaimClinician && claimClinicians.includes(eligClinician);
    matchingEligs.push({...elig, _isUsed: isUsed, _clinicianMismatch: clinicianMismatch, _clinicianMatch: clinicianPositivelyMatched});
  }
  
  if (shouldLog) {
    if (matchingEligs.length === 0) {
      console.log(`   ❌ No match found after checking all ${eligList.length} eligibilities (see rejection reasons above)`);
    } else {
      console.log(`   ✅ Total matches found: ${matchingEligs.length}`);
    }
  }
  
  return matchingEligs;
}

function checkClinicianMatch(claimClinicians, eligClinician) {
  if (!eligClinician || !claimClinicians?.length) return true;
  const normElig = normalizeClinician(eligClinician);
  return claimClinicians.some(c => normalizeClinician(c) === normElig);
}

/**
 * Select the best matching eligibility from multiple options
 * Ranks by:
 * 1. Eligible with a positive clinician match (elig clinician non-empty and equals claim
 *    clinician) — always preferred over Blocked/Cancelled.
 * 2. Blocked/Cancelled — preferred over Eligible when no Eligible has a positive clinician
 *    match.  This handles the scenario where the correct eligibility was blocked/cancelled
 *    and an unrelated Eligible (with a blank clinician field) was otherwise being selected.
 * 3. Valid (passes all validation checks) over invalid
 * 4. Not already used (prefer unused over used)
 * 5. Service category/consultation status match specificity
 * 6. Date (most recent first)
 * 7. Request number (highest/most recent first)
 *
 * @param {Array} eligibilities - Array of matching eligibilities with _isUsed flag
 * @param {string} claimDepartment - The claim's department/service category
 * @param {string} claimPackage - The claim's package name for validation
 * @returns {Object|null} - The best matching eligibility
 */
function selectBestEligibility(eligibilities, claimDepartment = '', claimPackage = '') {
  if (!eligibilities || eligibilities.length === 0) return null;
  
  // PRE-FILTER: Remove eligibilities that fail validation
  // This ensures we always pick a valid eligibility if one exists
  const validEligibilities = eligibilities.filter(elig => {
    const statusLower = (elig.Status || '').toLowerCase();
    const isBlockedOrCancelled = statusLower === 'blocked' || statusLower === 'cancelled';

    // Must have 'eligible', 'blocked', or 'cancelled' status
    if (statusLower !== 'eligible' && !isBlockedOrCancelled) return false;
    
    // Check service category validity
    const categoryCheck = isServiceCategoryValid(
      elig['Service Category'], 
      elig['Consultation Status'], 
      claimDepartment
    );
    if (!categoryCheck.valid) return false;
    
    // Check package name match if both exist
    if (claimPackage && elig['Package Name']) {
      if (!packageNamesMatch(claimPackage, elig['Package Name'])) return false;
    }
    
    return true;
  });
  
  // If no valid eligibilities found, fall back to all eligibilities
  // This allows the system to show WHY it failed (package mismatch, etc.)
  const eligsToConsider = validEligibilities.length > 0 ? validEligibilities : eligibilities;
  
  if (eligsToConsider.length === 1) return eligsToConsider[0];
  
  const dept = (claimDepartment || '').toLowerCase().trim();

  // STATUS PRIORITY: Eligible eligibilities with a positive clinician match always win.
  // B/C eligibilities are preferred over Eligible eligibilities that passed only through
  // a neutral clinician check (blank eligibility clinician or blank claim clinician) —
  // that is the scenario where the "correct" eligibility was blocked/cancelled and an
  // unrelated Eligible (with no clinician on it) was otherwise being selected.
  // If a genuine Eligible+clinician-match exists alongside a B/C, use the Eligible.
  const eligiblesWithStrongClinician = eligsToConsider.filter(e =>
    (e.Status || '').toLowerCase() === 'eligible' && e._clinicianMatch === true
  );
  const blockedCancelledEligs = eligsToConsider.filter(e => {
    const s = (e.Status || '').toLowerCase();
    return s === 'blocked' || s === 'cancelled';
  });
  const statusPrioritizedEligs =
    eligiblesWithStrongClinician.length > 0 ? eligiblesWithStrongClinician :
    blockedCancelledEligs.length > 0 ? blockedCancelledEligs :
    eligsToConsider;

  // PRIORITY 2: Always prefer unused eligibilities over used ones
  // Separate unused and used eligibilities
  const unusedEligibilities = statusPrioritizedEligs.filter(e => !e._isUsed);
  const usedEligibilities = statusPrioritizedEligs.filter(e => e._isUsed);
  
  // If there are ANY unused eligibilities, only consider those
  // Otherwise, fall back to used eligibilities
  const eligibilitiesToScore = unusedEligibilities.length > 0 ? unusedEligibilities : usedEligibilities;
  
  console.log(`   🎯 Best eligibility selection (${eligibilities.length} total: ${validEligibilities.length} valid, ${eligsToConsider.length - validEligibilities.length} invalid, ${eligiblesWithStrongClinician.length} eligible+clinician-match, ${blockedCancelledEligs.length} blocked/cancelled, ${unusedEligibilities.length} unused, ${usedEligibilities.length} used):`);
  
  // Score each eligibility
  const scored = eligibilitiesToScore.map(elig => {
    let score = 0;
    
    const serviceCategory = (elig['Service Category'] || '').toLowerCase().trim();
    const consultationStatus = (elig['Consultation Status'] || '').toLowerCase().trim();
    
    // Exact service category match (weight: 100)
    if (dept && serviceCategory && dept.includes(serviceCategory)) {
      score += 100;
    } else if (dept && serviceCategory && serviceCategory.includes(dept)) {
      score += 100;
    }
    
    // Specific service category (weight: 50) - prefer specific over generic
    if (serviceCategory === 'consultation') {
      score += 50;
      // Consultation with specific status (weight: +25)
      if (consultationStatus === 'elective' || consultationStatus === 'emergency') {
        score += 25;
      }
    } else if (serviceCategory && serviceCategory !== 'other op services' && serviceCategory !== 'other') {
      score += 40; // Other specific categories
    }
    
    // Generic categories get lower score (weight: 10)
    if (serviceCategory === 'other op services' || serviceCategory === 'other') {
      score += 10;
    }
    
    return { elig, score };
  });
  
  // Sort by score first (highest first), then by date (most recent first), then by request number (highest first)
  scored.sort((a, b) => {
    // Primary: Sort by score (descending)
    if (b.score !== a.score) {
      return b.score - a.score;
    }
    
    // Secondary: Sort by "Answered On" date (most recent first)
    const dateA = DateHandler.parse(a.elig['Answered On']);
    const dateB = DateHandler.parse(b.elig['Answered On']);
    
    // If both dates are valid, compare them
    if (dateA && dateB) {
      const dateDiff = dateB.getTime() - dateA.getTime();
      if (dateDiff !== 0) return dateDiff;  // Dates are different, use date order
      // Dates are equal, continue to next tiebreaker
    }
    
    // Fallback: If one date is invalid, prefer the valid one
    if (dateA && !dateB) return -1;  // A has valid date, prefer A
    if (!dateA && dateB) return 1;   // B has valid date, prefer B
    
    // Tertiary: Sort by Eligibility Request Number (higher/more recent first)
    const reqNumA = extractRequestNumber(a.elig['Eligibility Request Number']);
    const reqNumB = extractRequestNumber(b.elig['Eligibility Request Number']);
    if (reqNumA !== null && reqNumB !== null) {
      return reqNumB - reqNumA;  // Higher number (more recent) first
    }
    
    return 0;  // Keep original order only as absolute last resort
  });
  
  // Log all options with their scores
  scored.forEach((item, idx) => {
    const status = item.elig._isUsed ? '(USED)' : '(available)';
    const svc = item.elig['Service Category'] || 'N/A';
    const consult = item.elig['Consultation Status'] ? ` - ${item.elig['Consultation Status']}` : '';
    console.log(`      ${idx + 1}. Score ${item.score}: ${svc}${consult} ${status}`);
  });
  
  // If we had to use a used eligibility, log a warning
  if (scored[0].elig._isUsed) {
    console.log(`   ⚠️ WARNING: Selected a USED eligibility (no unused options available)`);
  }
  
  console.log(`   ✅ Selected: #${scored[0].elig['Eligibility Request Number']} (score: ${scored[0].score})`);
  
  return scored[0].elig;
}

// Helper function to extract numeric portion from Eligibility Request Number
function extractRequestNumber(requestNumber) {
  if (!requestNumber) return null;
  // Extract numeric portion from format like "eELIG - 229624246"
  const match = requestNumber.match(/(\d+)/);
  return match ? parseInt(match[1], 10) : null;
}

function isServiceCategoryValid(serviceCategory, consultationStatus, rawPackage) {
  if (!serviceCategory) return { valid: true };
  const category = serviceCategory.trim().toLowerCase();
  const pkgRaw = rawPackage || '';
  const pkg = pkgRaw.toLowerCase();
  
  const rules = SERVICE_PACKAGE_RULES[serviceCategory];
  if (!rules) return { valid: true };  // No rules defined for this category
  
  // Check status-specific rules first if consultation status is provided
  if (consultationStatus && rules.statusRules) {
    const status = consultationStatus.toLowerCase();
    const statusRule = rules.statusRules[status];
    
    if (statusRule) {
      // Check word boundary terms (e.g., 'ent' as whole word only)
      if (statusRule.wordBoundaryTerms && statusRule.wordBoundaryTerms.length > 0) {
        for (const term of statusRule.wordBoundaryTerms) {
          // Escape special regex characters to prevent regex injection
          const escapedTerm = escapeRegex(term);
          const regex = new RegExp(`\\b${escapedTerm}\\b`, 'i');
          if (regex.test(pkg)) {
            return { valid: true };
          }
        }
      }
      
      // Check allowed departments (substring matching)
      if (statusRule.allowedDepartments && statusRule.allowedDepartments.length > 0) {
        const isAllowedDept = statusRule.allowedDepartments.some(dept => pkg.includes(dept));
        if (isAllowedDept) {
          return { valid: true };
        }
      }
      
      // If no status-specific match found, continue to general keyword validation below
    }
  }
  
  // Check general keyword requirements
  const keywords = rules.keywords || [];
  if (keywords.length > 0) {
    if (pkg && !keywords.some(keyword => pkg.includes(keyword))) {
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

/**
 * When no eligibility match was found for a claim, diagnose the specific reasons
 * to provide a more targeted error message.
 * Iterates over all date-matching eligible records and collects every failure
 * reason encountered (clinician mismatch and/or service category mismatch).
 * Returns a combined string such as 'Wrong Service Category', or null when no specific
 * diagnosis can be made.
 *
 * @param {Map} eligMap - Map of normalised member IDs to eligibility records
 * @param {Date} claimDate - Parsed claim date
 * @param {string} normalizedMemberID - Already-normalized member ID
 * @param {string} claimClinician - Clinician value from the claim row
 * @param {string} claimDepartment - Department/service category from the claim row
 * @returns {string|null}
 */
function diagnoseEligibilityFailure(eligMap, claimDate, normalizedMemberID, claimClinician, claimDepartment) {
  const eligList = eligMap.get(normalizedMemberID) || [];
  if (!eligList.length) return null;

  // Collect all eligibilities that match the date, have 'eligible' status,
  // and are not already consumed by a different claim.
  const dateMatches = eligList.filter(e => {
    const eligDate = DateHandler.parse(e['Answered On'] || e['Ordered On'], { preferMDY: false });
    const reqNo = e['Eligibility Request Number'];
    return DateHandler.isSameDay(claimDate, eligDate) &&
      (e.Status || '').toLowerCase() === 'eligible' &&
      !usedEligibilities.has(reqNo);
  });

  if (!dateMatches.length) return null;

  // Only diagnose against eligibilities that are clinician-relevant to this claim.
  // This avoids reporting "Wrong Service Category" based solely on unrelated
  // eligibilities that belong to other clinicians.
  const normalizedClaimClinician = normalizeClinician(claimClinician || '');
  const clinicianRelevantMatches = dateMatches.filter(elig => {
    const eligClinician = (elig.Clinician || '').trim();
    if (!normalizedClaimClinician) return true;
    if (!eligClinician) return false;
    return normalizeClinician(eligClinician) === normalizedClaimClinician;
  });

  if (!clinicianRelevantMatches.length) return null;

  // Gather every failure reason across all date-matching records
  const reasons = new Set();
  const dept = (claimDepartment || '').toLowerCase();

  for (const elig of clinicianRelevantMatches) {
    const categoryCheck = isServiceCategoryValid(elig['Service Category'], elig['Consultation Status'], dept);
    if (!categoryCheck.valid) {
      reasons.add('Wrong Service Category');
    }
  }

  if (!reasons.size) return null;
  return Array.from(reasons).join(', ');
}

/**
 * Scan all eligibility records for one whose EID matches rawEID.
 * Used as a fallback when member ID lookup yields no results.
 * Dashes are stripped from both values before comparison.
 *
 * @param {Map} eligMap - Map of normalised member IDs to eligibility records
 * @param {string} rawEID - Emirates ID from the claim row (may include dashes)
 * @param {Date} claimDate - Parsed claim date
 * @param {string[]} claimClinicians - Clinician values from the claim row
 * @param {string} insurance - Insurance/provider name
 * @param {string} claimDepartment - Department from the claim row
 * @param {string} claimPackage - Package name from the claim row
 * @returns {{ eligibility: Object|null, matchedMemberID: string }|null}
 */
function findEligibilityByEID(eligMap, rawEID, claimDate, claimClinicians, insurance, claimDepartment, claimPackage) {
  if (!rawEID) return null;
  const strippedEID = String(rawEID).replace(/-/g, '').trim();
  if (!strippedEID) return null;

  // Collect all map keys whose records contain a matching EID
  const matchedMemberIDs = [];
  for (const [mappedMemberID, records] of eligMap) {
    for (const record of records) {
      const recordEID = String(record['EID'] || '').replace(/-/g, '').trim();
      if (recordEID && recordEID === strippedEID) {
        matchedMemberIDs.push(mappedMemberID);
        break;
      }
    }
  }

  if (matchedMemberIDs.length === 0) return null;

  const firstMatchedMemberID = matchedMemberIDs[0];

  // Try to find a usable eligibility via the EID-matched member ID(s)
  for (const matchedMemberID of matchedMemberIDs) {
    const matches = findEligibilityForClaim(eligMap, claimDate, matchedMemberID, claimClinicians, insurance, false, 0);
    const best = selectBestEligibility(matches, claimDepartment, claimPackage);
    if (best) {
      return { eligibility: best, matchedMemberID };
    }
  }

  // EID matched a record but no valid eligibility for this claim date/clinician
  return { eligibility: null, matchedMemberID: firstMatchedMemberID };
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
      // Check for Combined report first (has Pri. Claim No + Total Amount; 'Visit Id' alone is not enough since Insta reports also have it)
      if (sample.hasOwnProperty('Pri. Claim No') && sample.hasOwnProperty('Total Amount')) return 'Combined';
      if (sample.hasOwnProperty('Pri. Claim No')) return 'Insta';
      if (sample.hasOwnProperty('Pri. Claim ID')) return 'Odoo';
    }
    return 'Generic';
  }
  
  // If it's already an array of objects
  if (Array.isArray(rawData) && rawData.length > 0 && typeof rawData[0] === 'object' && !Array.isArray(rawData[0])) {
    const sample = rawData[0];
    // Check for Combined report first (has Pri. Claim No + Total Amount; 'Visit Id' alone is not enough since Insta reports also have it)
    if (sample.hasOwnProperty('Pri. Claim No') && sample.hasOwnProperty('Total Amount')) return 'Combined';
    if (sample.hasOwnProperty('Pri. Claim No')) return 'Insta';
    if (sample.hasOwnProperty('Pri. Claim ID')) return 'Odoo';
    return 'Generic';
  }
  
  // If it has a {headers, rows} shape
  if (rawData.rows && Array.isArray(rawData.rows) && rawData.rows.length > 0) {
    const sample = rawData.rows[0];
    // Check for Combined report first (has Pri. Claim No + Total Amount; 'Visit Id' alone is not enough since Insta reports also have it)
    if (sample.hasOwnProperty('Pri. Claim No') && sample.hasOwnProperty('Total Amount')) return 'Combined';
    if (sample.hasOwnProperty('Pri. Claim No')) return 'Insta';
    if (sample.hasOwnProperty('Pri. Claim ID')) return 'Odoo';
    return 'Generic';
  }
  
  return 'Generic';
}

/**
 * Read a clinician license value from a raw spreadsheet cell.
 * Preserves Excel boolean FALSE (and the text "FALSE") as the sentinel
 * string 'FALSE' so validation can detect it and mark the claim as unknown.
 * Without this, boolean false is silently dropped by the || '' fallback.
 * @param {*} val - Raw cell value
 * @returns {string}
 */
function readClinicianLicense(val) {
  if (val === false || (typeof val === 'string' && val.trim().toUpperCase() === 'FALSE')) return 'FALSE';
  if (val == null || val === '') return '';
  return String(val);
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
    const isCombined = sample.hasOwnProperty('Pri. Claim No') && sample.hasOwnProperty('Total Amount');
    const isInsta = sample.hasOwnProperty('Pri. Claim No') && !isCombined;
    const isOdoo = sample.hasOwnProperty('Pri. Claim ID');
    return rawData.map(row => {
      if (isCombined) {
        return {
          claimID: row['Pri. Claim No'] || '',
          visitID: row['Visit Id'] || '',
          memberID: row['Pri. Patient Insurance Card No'] || '',
          claimDate: row['Encounter Date'] || '',
          clinician: readClinicianLicense(row['Clinician License']),
          clinicianName: row['Clinician Name'] || '',
          department: row['Department'] || '',
          packageName: row['Pri. Plan Type'] || '',
          insuranceCompany: row['Pri. Plan Type'] || '',
          claimStatus: '',
          fileNo: row['Patient Code'] || '',
          admittingDoctor: '',
          openedBy: row['Opened by'] || '',
          price: row['Total Amount'] ?? '',
          facilityID: row['Facility ID'] || '',
          sourceFile: row['Source File'] || '',
          emiratesID: getField(row, ['Emirates ID No', 'Emirates ID No.']) || ''
        };
      } else if (isInsta) {
        return {
          claimID: row['Pri. Claim No'] || '',
          visitID: row['Visit Id'] || '',
          memberID: row['Pri. Patient Insurance Card No'] || '',
          claimDate: row['Encounter Date'] || '',
          clinician: readClinicianLicense(row['Clinician License']),
          clinicianName: row['Clinician Name'] || '',
          department: row['Department'] || '',
          packageName: row['Pri. Plan Type'] || row['Pri. Plan Name'] || row['Pri. Payer Name'] || '',
          insuranceCompany: row['Pri. Plan Type'] || row['Pri. Plan Name'] || row['Pri. Payer Name'] || '',
          claimStatus: row['Codification Status'] || '',
          fileNo: row['Patient Code'] || '',
          admittingDoctor: '',  // Insta reports don't have a separate Admitting Doctor column
          openedBy: row['Opened by'] || '',
          price: row['Net Amount'] ?? row['Gross Amount'] ?? '',
          facilityID: row['Facility ID'] || '',
          emiratesID: getField(row, ['Emirates ID No', 'Emirates ID No.']) || ''
        };
      } else if (isOdoo) {
        return {
          claimID: row['Pri. Claim ID'] || '',
          visitID: row['Visit Id'] || '',
          memberID: row['Pri. Member ID'] || '',
          claimDate: row['Adm/Reg. Date'] || '',
          clinician: readClinicianLicense(row['Admitting License']),
          insuranceCompany: row['Pri. Plan Type'] || row['Pri. Sponsor'] || '',
          packageName: row['Pri. Plan Type'] || row['Pri. Sponsor'] || '',
          claimStatus: row['Codification Status'] || '',
          fileNo: row['MR No.'] || row['MR No'] || '',
          admittingDoctor: row['Admitting Doctor'] || '',
          openedBy: row['Opened by'] || '',
          price: row['Total Sponsor Amt'] ?? '',
          facilityID: row['Center Name'] || '',
          emiratesID: getField(row, ['Emirates ID No', 'Emirates ID No.']) || ''
        };
      } else {
        return {
          claimID: row['ClaimID'] || '',
          visitID: row['Visit Id'] || '',
          memberID: row['PatientCardID'] || '',
          claimDate: row['ClaimDate'] || '',
          clinician: readClinicianLicense(row['Clinician License']),
          insuranceCompany: row['Insurance Company'] || '',
          department: row['Clinic'] || '',
          claimStatus: row['VisitStatus'] || '',
          fileNo: row['MR No'] || row['Patient Code'] || row['File No'] || row['FileNo'] || '',
          admittingDoctor: row['Admitting Doctor'] || row['Doctor'] || row['Physician'] || '',
          openedBy: row['Opened by'] || '',
          price: row['Total Amount'] ?? '',
          facilityID: row['Facility ID'] || '',
          emiratesID: getField(row, ['Emirates ID No', 'Emirates ID No.']) || ''
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
    const isCombined = Object.prototype.hasOwnProperty.call(r, 'Pri. Claim No') && Object.prototype.hasOwnProperty.call(r, 'Total Amount');
    const isInsta = Object.prototype.hasOwnProperty.call(r, 'Pri. Claim No') && !isCombined;
    const isOdoo = Object.prototype.hasOwnProperty.call(r, 'Pri. Claim ID');

    if (isCombined) {
      return {
        claimID: r['Pri. Claim No'] || '',
        visitID: r['Visit Id'] || '',
        memberID: r['Pri. Patient Insurance Card No'] || '',
        claimDate: r['Encounter Date'] || '',
        clinician: readClinicianLicense(r['Clinician License']),
        clinicianName: r['Clinician Name'] || '',
        department: r['Department'] || '',
        packageName: r['Pri. Plan Type'] || '',
        insuranceCompany: r['Pri. Plan Type'] || '',
        claimStatus: '',
        fileNo: r['Patient Code'] || '',
        admittingDoctor: '',
        openedBy: r['Opened by'] || '',
        price: getField(r, ['Total Amount']),  // getField preserves 0 values, unlike || ''
        facilityID: r['Facility ID'] || '',
        sourceFile: r['Source File'] || '',
        emiratesID: getField(r, ['Emirates ID No', 'Emirates ID No.']) || ''
      };
    } else if (isInsta) {
      return {
        claimID: r['Pri. Claim No'] || '',
        visitID: r['Visit Id'] || '',
        memberID: r['Pri. Patient Insurance Card No'] || '',
        claimDate: r['Encounter Date'] || '',
        clinician: readClinicianLicense(r['Clinician License']),
        clinicianName: r['Clinician Name'] || '',
        department: r['Department'] || '',
        packageName: r['Pri. Plan Type'] || r['Pri. Plan Name'] || r['Pri. Payer Name'] || '',
        insuranceCompany: r['Pri. Plan Type'] || r['Pri. Plan Name'] || r['Pri. Payer Name'] || '',
        claimStatus: r['Codification Status'] || '',
        fileNo: r['Patient Code'] || '',
        admittingDoctor: '',  // Insta reports don't have a separate Admitting Doctor column
        openedBy: r['Opened by'] || '',
        price: getField(r, ['Net Amount', 'Gross Amount']),
        facilityID: r['Facility ID'] || '',
        emiratesID: getField(r, ['Emirates ID No', 'Emirates ID No.']) || ''
      };
    } else if (isOdoo) {
      return {
        claimID: r['Pri. Claim ID'] || '',
        visitID: r['Visit Id'] || '',
        memberID: r['Pri. Member ID'] || '',
        claimDate: r['Adm/Reg. Date'] || '',
        clinician: readClinicianLicense(r['Admitting License']),
        insuranceCompany: r['Pri. Plan Type'] || r['Pri. Sponsor'] || '',
        packageName: r['Pri. Plan Type'] || r['Pri. Sponsor'] || '',
        claimStatus: r['Codification Status'] || '',
        fileNo: r['MR No.'] || r['MR No'] || '',
        admittingDoctor: r['Admitting Doctor'] || '',
        openedBy: r['Opened by'] || '',
        price: getField(r, ['Total Sponsor Amt']),
        facilityID: r['Center Name'] || '',
        emiratesID: getField(r, ['Emirates ID No', 'Emirates ID No.']) || ''
      };
    } else {
      const out = {
        claimID: r['ClaimID'] || r['Pri. Claim No'] || r['Pri. Claim ID'] || getField(r, ['ClaimID','Pri. Claim No','Pri. Claim ID','Claim ID','Pri. Claim ID']) || '',
        visitID: r['Visit Id'] || '',
        memberID: r['Pri. Member ID'] || r['Pri. Patient Insurance Card No'] || r['PatientCardID'] || getField(r, ['PatientCardID','Patient Insurance Card No','Card Number / DHA Member ID']) || '',
        claimDate: r['Encounter Date'] || r['Adm/Reg. Date'] || r['ClaimDate'] || getField(r, ['Encounter Date','ClaimDate','Adm/Reg. Date','Date']) || '',
        clinician: readClinicianLicense(
          r['Clinician License'] != null && r['Clinician License'] !== ''
            ? r['Clinician License']
            : r['Admitting License'] || r['OrderDoctor'] || getField(r, ['Clinician','Admitting License','OrderDoctor'])),
        clinicianName: r['Clinician Name'] || '',
        department: r['Department'] || r['Clinic'] || r['Admitting Department'] || getField(r, ['Department','Clinic','Admitting Department']) || '',
        packageName: r['Pri. Plan Type'] || r['Pri. Plan Name'] || r['Pri. Payer Name'] || r['Insurance Company'] || r['Pri. Sponsor'] || getField(r, ['Pri. Plan Type','Pri. Plan Name','Pri. Payer Name','Insurance Company','Package','Pri. Sponsor']) || '',
        insuranceCompany: r['Pri. Plan Type'] || r['Pri. Payer Name'] || r['Insurance Company'] || r['Pri. Sponsor'] || getField(r, ['Pri. Plan Type','Payer Name','Insurance Company','Pri. Payer Name','Pri. Sponsor']) || '',
        claimStatus: r['Codification Status'] || r['VisitStatus'] || r['Status'] || getField(r, ['Codification Status','VisitStatus','Status','Claim Status']) || '',
        fileNo: r['MR No.'] || r['MR No'] || r['Patient Code'] || getField(r, ['MR No.','MR No','Patient Code','File No','FileNo']) || '',
        admittingDoctor: r['Admitting Doctor'] || getField(r, ['Admitting Doctor','Doctor','Physician']) || '',
        openedBy: r['Opened by'] || '',
        price: getField(r, ['Net Amount', 'Gross Amount', 'Total Sponsor Amt', 'Total Amount']),
        facilityID: r['Facility ID'] || r['Center Name'] || '',
        emiratesID: getField(r, ['Emirates ID No', 'Emirates ID No.']) || ''
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
        claimID, 
        visitID: row.visitID || '',
        memberID,
        rawMemberID,
        encounterStart: formattedDate,
        encounterDate: claimDate, // Store original Date object to avoid re-parsing issues
        status: 'VVIP', finalStatus: 'valid', 
        remarks: ['VVIP member, eligibility check bypassed'], 
        fullEligibilityRecord: null,
        fileNo: row.fileNo || '',
        admittingDoctor: row.admittingDoctor || '',
        clinicianName: row.clinicianName || '',
        department: row.department || '',
        clinician: row.clinician || '',
        packageName: row.packageName || '',
        provider: row.insuranceCompany || '',
        openedBy: row.openedBy || '',
        price: row.price || '',
        facilityID: row.facilityID || '',
        sourceFile: row.sourceFile || ''
      });
      continue;
    }

    // Detect garbage / multi-field paste: real member IDs never contain whitespace
    const hasWhitespaceMemberID = /\s/.test(rawMemberID);

    // Check for leading zeroes in original memberID
    // Use /^0/ to detect any ID starting with zero (including all-zeroes like "0000")
    const hasLeadingZero = /^0/.test(rawMemberID);
    
    if (claimIndex <= 3) {
      console.log(`  Claim #${claimIndex} (${claimID}): Processing - Member ${memberID}, Date ${formattedDate}`);
    }
    
    // Pass claimIndex to enable logging for first 3 claims
    const matchingEligibilities = findEligibilityForClaim(eligMap, claimDate, memberID, [row.clinician], insurance, false, claimIndex);
    
    // Select the BEST match from all available eligibilities
    // Pass package name and department for validation filtering
    const selectedEligibility = selectBestEligibility(matchingEligibilities, row.department, row.packageName);
    
    // Mark the selected eligibility as used (only mark the one we actually use)
    if (selectedEligibility && selectedEligibility['Eligibility Request Number']) {
      if (!usedEligibilities.has(selectedEligibility['Eligibility Request Number'])) {
        usedEligibilities.add(selectedEligibility['Eligibility Request Number']);
        console.log(`   📌 Marked eligibility #${selectedEligibility['Eligibility Request Number']} as USED`);
      }
    }
    
    // Use selected eligibility for validation
    let eligibility = selectedEligibility;
    let eidMatchedMemberID = null;
    let finalStatus = 'invalid', remarks = [];

    // EID fallback — only when member ID was not found at all and Emirates ID is available
    if (!eligibility && row.emiratesID && !eligMap.has(memberID)) {
      const eidResult = findEligibilityByEID(
        eligMap, row.emiratesID, claimDate, [row.clinician], insurance, row.department, row.packageName
      );
      if (eidResult) {
        eligibility = eidResult.eligibility;
        eidMatchedMemberID = eidResult.matchedMemberID;
        // Repopulate matchingEligibilities from the EID-resolved member ID so the
        // details modal has records to display (the original lookup returned empty).
        if (eidMatchedMemberID) {
          const eidMatches = findEligibilityForClaim(eligMap, claimDate, eidMatchedMemberID, [row.clinician], insurance, false, claimIndex);
          if (eidMatches && eidMatches.length > 0) {
            matchingEligibilities.length = 0;
            eidMatches.forEach(r => matchingEligibilities.push(r));
          } else {
            // Filtered search returned nothing (date/status mismatch, etc.) —
            // fall back to ALL raw records for this member so the details button
            // is always shown whenever "Wrong Member ID" appears in remarks.
            const allEidRecords = eligMap.get(eidMatchedMemberID) || [];
            if (allEidRecords.length > 0) {
              matchingEligibilities.length = 0;
              allEidRecords.forEach(r => matchingEligibilities.push({...r, _isUsed: usedEligibilities.has(r['Eligibility Request Number'])}));
            }
          }
        }
      }
    }
    const packageLower = (row.packageName || '').toLowerCase();

    // Clinician license is 'FALSE' — not found in our database
    if (hasWhitespaceMemberID) {
      finalStatus = 'unknown';
      remarks.push('Member ID appears to contain extra data; please recheck this claim.');
    } else if (row.clinician === 'FALSE') {
      finalStatus = 'unknown';
      remarks.push("This clinician hasn't been added to our database yet.");
    // Daman High-End claims cannot be determined as valid or invalid; mark as unknown
    } else if (packageLower.includes('daman') && packageLower.includes('high-end')) {
      finalStatus = 'unknown';
      remarks.push('Daman High-End claims are marked as unknown.');
    // Leading zero + matched eligibility → unknown so the record can still be used
    } else if (hasLeadingZero && eligibility) {
      finalStatus = 'unknown';
      remarks.push('Member ID has a leading zero; marked as unknown.');
    } else if (!eligibility) {
      if (eidMatchedMemberID && eidMatchedMemberID !== memberID) {
        remarks.push('Wrong Member ID');
      }
      const lookupMemberID = eidMatchedMemberID || memberID;
      const rawEligList = eligMap.get(lookupMemberID) || [];
      if (rawEligList.length > 0) {
        // Member found in map but no eligibility passed all filters —
        // attempt to surface the specific single-field failure reason.
        // When the root cause is already "Wrong Member ID", suppress clinician
        // diagnosis (clinician mismatch is irrelevant on the wrong member's records).
        const effectiveDiagnoseClinician = (eidMatchedMemberID && eidMatchedMemberID !== memberID) ? '' : row.clinician;
        const failureReason = diagnoseEligibilityFailure(eligMap, claimDate, lookupMemberID, effectiveDiagnoseClinician, row.department);
        remarks.push(failureReason || 'Eligibility Not Taken');
      } else {
        remarks.push('Eligibility Not Taken');
      }
    } else if (eligibility.Status?.toLowerCase() === 'eligible') {
      const categoryCheck = isServiceCategoryValid(eligibility['Service Category'], eligibility['Consultation Status'], (row.department || '').toLowerCase());
      if (categoryCheck.valid) {
        // Validate package name match if both claim and eligibility have package names
        if (row.packageName && eligibility['Package Name']) {
          // Use special matching logic that handles Thiqa/TC packages
          if (!packageNamesMatch(row.packageName, eligibility['Package Name'])) {
            finalStatus = 'invalid';
            // Normalize package names for display to show tier levels (e.g., "Daman Enhanced") instead of specific plan names (e.g., "Sahtak")
            const normalizedClaimPackage = normalizePackageNameForDisplay(row.packageName);
            const normalizedEligPackage = normalizePackageNameForDisplay(eligibility['Package Name']);
            remarks.push(`Registered under ${normalizedClaimPackage}, ${normalizedEligPackage} as per Eligibility`);
          } else {
            finalStatus = 'valid';
          }
        } else {
          // No package name to validate, so it's valid based on other checks
          finalStatus = 'valid';
        }
      } else {
        remarks.push('Wrong Service Category');
      }
    } else {
      if (eligibility.Status?.toLowerCase() === 'cancelled') {
        remarks.push('Eligibility Not Taken');
      } else {
        remarks.push(`Eligibility status: ${eligibility.Status}`);
      }
    }

    // If EID fallback resolved a different member ID, the claim had the wrong member ID.
    // Mark it invalid and surface "Wrong Member ID" regardless of which branch above ran
    // (e.g. the leading-zero branch sets finalStatus = 'invalid' without adding this remark).
    if (eidMatchedMemberID && eidMatchedMemberID !== memberID) {
      if (!remarks.includes('Wrong Member ID')) {
        remarks.push('Wrong Member ID');
      }
      finalStatus = 'invalid';
    }

    results.push({
      claimID, 
      visitID: row.visitID || '',
      memberID,
      rawMemberID,
      encounterStart: formattedDate,
      encounterDate: claimDate, // Store original Date object to avoid re-parsing issues
      packageName: eligibility?.['Package Name'] || row.packageName || '',
      provider: insurance,
      clinician: eligibility?.Clinician || row.clinician || '',
      clinicianName: row.clinicianName || '',
      department: row.department || '',
      serviceCategory: eligibility?.['Service Category'] || '',
      consultationStatus: eligibility?.['Consultation Status'] || '',
      status: eligibility?.Status || '',
      claimStatus: row.claimStatus || '',
      remarks, finalStatus,
      wrongMemberID: !!(eidMatchedMemberID && eidMatchedMemberID !== memberID),
      fullEligibilityRecord: eligibility,
      allEligibilityRecords: matchingEligibilities, // Store ALL matches
      fileNo: row.fileNo || '',
      admittingDoctor: row.admittingDoctor || '',
      openedBy: row.openedBy || '',
      price: row.price ?? '',
      facilityID: row.facilityID || '',
      sourceFile: row.sourceFile || ''
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
      <th>Department</th>
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
    if (result.allEligibilityRecords && result.allEligibilityRecords.length > 0) {
      // Show how many eligibilities were matched
      const count = result.allEligibilityRecords.length;
      const label = count === 1 ? 'View details' : `View details (${count} matches)`;
      detailsCellHtml = `<button class="btn btn-sm btn-outline-primary eligibility-details" data-index="${index}" data-claimdate="${escapeHtml(result.encounterStart)}">${label}</button>`;
    } else if (eligMap && typeof eligMap.get === 'function' && (eligMap.get(normalizeMemberID(result.memberID)) || []).length) {
      // Otherwise, if there are eligibilities in the map for this member, offer a secondary button to view all eligibilities for the member
      detailsCellHtml = `<button class="btn btn-sm btn-outline-secondary show-all-eligibilities" 
        data-index="${index}"
        data-member="${escapeHtml(result.memberID)}" 
        data-claimdate="${escapeHtml(result.encounterStart)}"
        data-claimclinician="${escapeHtml(result.clinician || '')}"
        data-claimpackage="${escapeHtml(result.packageName || '')}" aria-label="View Eligibilities">View Eligs</button>`;
      // Conditionally add diagnostics button based on toggle state
      if (showDiagnosticsButtons) {
        detailsCellHtml += ` <button class="btn btn-sm btn-outline-info show-diagnostics" data-index="${index}" title="Show diagnostic logging for this claim">
          <i class="bi bi-terminal"></i> Diagnostics
        </button>`;
      }
    } else {
      // Even if no eligibility found, conditionally add diagnostics button to help debug why
      if (showDiagnosticsButtons) {
        detailsCellHtml = `<button class="btn btn-sm btn-outline-info show-diagnostics" data-index="${index}" title="Show diagnostic logging for this claim">
          <i class="bi bi-terminal"></i> Diagnostics
        </button>`;
      }
    }

    row.innerHTML = `
      <td>${escapeHtml(result.claimID)}</td>
      <td>${escapeHtml(result.rawMemberID || result.memberID)}</td>
      <td>${escapeHtml(result.encounterStart)}</td>
      <td class="description-col">${escapeHtml(result.packageName)}</td>
      <td class="description-col">${escapeHtml(result.provider)}</td>
      <td class="description-col">${escapeHtml(result.clinician)}</td>
      <td class="description-col">${escapeHtml(result.department || '')}</td>
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
              <div id="modalTable" class="p-0" style="overflow:auto; max-height:70vh;"></div>
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
      if (!result?.allEligibilityRecords || result.allEligibilityRecords.length === 0) return;
      
      // Use stored Date object instead of re-parsing to avoid date shift issues
      const claimDate = result.encounterDate || null;
      const claimInfo = {
        claimClinician: result.clinician || '',
        claimPackage: result.packageName || '',
        wrongMemberID: result.wrongMemberID || false,
        claimMemberID: result.memberID || ''
      };
      
      const modalTable = document.getElementById("modalTable");
      const debugBtn = document.getElementById('modalDebugBtn');

      // Determine which record is selected/matched
      const selectedEligNumber = result.fullEligibilityRecord
        ? result.fullEligibilityRecord['Eligibility Request Number']
        : null;

      // Find the default tab (the selected/matched elig, or first)
      let defaultTabIdx = 0;
      if (selectedEligNumber && result.allEligibilityRecords.length > 1) {
        const found = result.allEligibilityRecords.findIndex(
          r => r['Eligibility Request Number'] === selectedEligNumber
        );
        if (found >= 0) defaultTabIdx = found;
      }

      window.__elig_current_debug = {
        mode: result.allEligibilityRecords.length === 1 ? 'single' : 'list',
        member: result.memberID,
        claimDate: result.encounterStart || '',
        record: result.allEligibilityRecords[0],
        listSnapshot: result.allEligibilityRecords,
        resultIndex: index
      };
      if (debugBtn) debugBtn.style.display = '';

      // Build Chrome-style tabbed interface for all eligibility records
      let tabsHtml = `<div class="elig-modal-tabs-wrapper">`;

      // Tab bar
      tabsHtml += `<ul class="nav elig-modal-tab-bar" role="tablist">`;
      result.allEligibilityRecords.forEach((rec, idx) => {
        const isSelected = selectedEligNumber &&
          rec['Eligibility Request Number'] === selectedEligNumber;
        const isActive = idx === defaultTabIdx;
        const reqNum = rec['Eligibility Request Number'] || '';
        const tabLabel = reqNum ? escapeHtml(reqNum) : `Elig #${idx + 1}`;
        const selectedBadge = isSelected
          ? ` <span class="badge bg-success ms-1 elig-tab-selected-badge">✓</span>`
          : (rec._isUsed ? ` <span class="badge bg-warning text-dark ms-1 elig-tab-used-badge">used</span>` : '');
        tabsHtml += `<li class="nav-item" role="presentation">
          <button class="nav-link elig-tab-btn${isActive ? ' active' : ''}" id="elig-tab-${idx}" data-elig-tab="${idx}" type="button" role="tab" aria-selected="${isActive ? 'true' : 'false'}" aria-controls="elig-tabpanel-${idx}">
            ${tabLabel}${selectedBadge}
          </button>
        </li>`;
      });
      tabsHtml += `</ul>`;

      // Tab content panels
      tabsHtml += `<div class="elig-modal-tab-content">`;
      result.allEligibilityRecords.forEach((rec, idx) => {
        const isActive = idx === defaultTabIdx;
        tabsHtml += `<div class="elig-tab-pane${isActive ? ' active' : ''}" id="elig-tabpanel-${idx}" role="tabpanel" aria-labelledby="elig-tab-${idx}">`;
        tabsHtml += formatEligibilityDetails(rec, result.memberID, claimDate, claimInfo);
        tabsHtml += `</div>`;
      });
      tabsHtml += `</div></div>`;

      modalTable.innerHTML = tabsHtml;

      // Wire up tab switching
      modalTable.querySelectorAll('.elig-tab-btn').forEach(tabBtn => {
        tabBtn.addEventListener('click', function () {
          const tabIdx = parseInt(this.dataset.eligTab, 10);
          modalTable.querySelectorAll('.elig-tab-btn').forEach(btn => {
            btn.classList.remove('active');
            btn.setAttribute('aria-selected', 'false');
          });
          modalTable.querySelectorAll('.elig-tab-pane').forEach(pane => {
            pane.classList.remove('active');
          });
          this.classList.add('active');
          this.setAttribute('aria-selected', 'true');
          const pane = modalTable.querySelector(`#elig-tabpanel-${tabIdx}`);
          if (pane) pane.classList.add('active');
        });
      });

      showModal();
    });
  });

  document.querySelectorAll(".show-all-eligibilities").forEach(btn => {
    btn.onclick = null;
    btn.addEventListener('click', function () {
      const index = parseInt(this.dataset.index, 10);
      const result = results[index];
      const member = this.dataset.member;
      const claimClinician = this.dataset.claimclinician || '';
      const claimPackage = this.dataset.claimpackage || '';
      // Use stored Date object instead of re-parsing to avoid timezone/parsing issues
      const claimDate = result && result.encounterDate ? result.encounterDate : null;
      const list = (typeof eligMap.get === 'function') ? (eligMap.get(normalizeMemberID(member)) || []) : [];
      const modalTable = document.getElementById("modalTable");
      window.__elig_current_debug = { mode: 'list', member, claimDate: result ? result.encounterStart : '', listSnapshot: list.slice(0,200) };
      const debugBtn = document.getElementById('modalDebugBtn'); if (debugBtn) debugBtn.style.display = '';

      if (!list.length) {
        modalTable.innerHTML = `<div class="p-3">No eligibilities found for <strong>${escapeHtml(member)}</strong></div>`;
        showModal();
        return;
      }

      const claimInfo = { claimClinician, claimPackage };

      // Pre-compute match status for each record to determine tab badge and default tab
      const recordMeta = list.map(rec => {
        const answeredOnRaw = rec['Answered On'] || rec['Ordered On'] || '';
        const eligDate = DateHandler.parse(answeredOnRaw);
        const reasons = [];
        if (claimDate && eligDate && !DateHandler.isSameDay(claimDate, eligDate)) reasons.push('date');
        const eligClinician = (rec['Clinician'] || '').trim();
        if (eligClinician && claimClinician && eligClinician !== claimClinician) reasons.push('clinician');
        const eligPackage = (rec['Package Name'] || '').trim();
        if (eligPackage && claimPackage && !packageNamesMatch(claimPackage, eligPackage)) reasons.push('package');
        const status = rec['Status'] || '';
        if (status.toLowerCase() !== 'eligible') reasons.push('status');
        return { isMatch: reasons.length === 0 };
      });

      // Default tab: first record that fully matches, else first
      let defaultTabIdx = recordMeta.findIndex(m => m.isMatch);
      if (defaultTabIdx < 0) defaultTabIdx = 0;

      // Build Chrome-style tabbed interface
      let tabsHtml = `<div class="elig-modal-tabs-wrapper">`;

      tabsHtml += `<ul class="nav elig-modal-tab-bar" role="tablist">`;
      list.forEach((rec, idx) => {
        const isActive = idx === defaultTabIdx;
        const reqNum = rec['Eligibility Request Number'] || '';
        const tabLabel = reqNum ? escapeHtml(reqNum) : `Elig #${idx + 1}`;
        const matchBadge = recordMeta[idx].isMatch
          ? ` <span class="badge bg-success ms-1 elig-tab-match-badge">✓</span>`
          : ` <span class="badge bg-danger ms-1 elig-tab-mismatch-badge">✗</span>`;
        tabsHtml += `<li class="nav-item" role="presentation">
          <button class="nav-link elig-tab-btn${isActive ? ' active' : ''}" id="elig-tab-${idx}" data-elig-tab="${idx}" type="button" role="tab" aria-selected="${isActive ? 'true' : 'false'}" aria-controls="elig-tabpanel-${idx}">
            ${tabLabel}${matchBadge}
          </button>
        </li>`;
      });
      tabsHtml += `</ul>`;

      tabsHtml += `<div class="elig-modal-tab-content">`;
      list.forEach((rec, idx) => {
        const isActive = idx === defaultTabIdx;
        tabsHtml += `<div class="elig-tab-pane${isActive ? ' active' : ''}" id="elig-tabpanel-${idx}" role="tabpanel" aria-labelledby="elig-tab-${idx}">`;
        tabsHtml += formatEligibilityDetails(rec, member, claimDate, claimInfo);
        tabsHtml += `</div>`;
      });
      tabsHtml += `</div></div>`;

      modalTable.innerHTML = tabsHtml;

      // Wire up tab switching
      modalTable.querySelectorAll('.elig-tab-btn').forEach(tabBtn => {
        tabBtn.addEventListener('click', function () {
          const tabIdx = parseInt(this.dataset.eligTab, 10);
          modalTable.querySelectorAll('.elig-tab-btn').forEach(btn => {
            btn.classList.remove('active');
            btn.setAttribute('aria-selected', 'false');
          });
          modalTable.querySelectorAll('.elig-tab-pane').forEach(pane => {
            pane.classList.remove('active');
          });
          this.classList.add('active');
          this.setAttribute('aria-selected', 'true');
          const pane = modalTable.querySelector(`#elig-tabpanel-${tabIdx}`);
          if (pane) pane.classList.add('active');
        });
      });

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
      // Use stored Date object to avoid re-parsing issues
      const claimDate = result.encounterDate || DateHandler.parse(result.encounterStart);
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
   Optional claimDate param can be used to colour date rows.
   claimInfo param can include: claimClinician, claimPackage to show mismatch reasons */
function formatEligibilityDetails(record, memberID, claimDate, claimInfo = {}) {
  if (!record) return '<div>No details</div>';

  const status = (record.Status || '').toString();
  const statusClass = status.toLowerCase() === 'eligible' ? 'status-badge eligible' : 'status-badge ineligible';
  const wrongMemberID = claimInfo.wrongMemberID || false;
  const claimMemberID = claimInfo.claimMemberID || memberID;
  let memberDisplay = escapeHtml(claimMemberID);
  if (wrongMemberID) {
    const correctID = record['Card Number / DHA Member ID'] || '';
    memberDisplay = `<span style="color:#dc3545;font-weight:bold;">${escapeHtml(claimMemberID)}</span>`
      + (correctID ? ` <span style="color:#6c757d;font-size:0.875em;">→ correct ID: <strong>${escapeHtml(correctID)}</strong></span>` : '');
  }
  let html = `<div class="mb-2"><strong>Member:</strong> ${memberDisplay} <span class="${statusClass}" style="margin-left:8px;">${escapeHtml(status)}</span></div>`;

  // Generate mismatch reasons
  const mismatches = [];
  
  // Check date mismatch
  const eligDate = DateHandler.parse(record["Answered On"]);
  if (claimDate && eligDate && !DateHandler.isSameDay(claimDate, eligDate)) {
    mismatches.push(`Date mismatch: Claim date is ${escapeHtml(DateHandler.format(claimDate))}, but eligibility date is ${escapeHtml(DateHandler.format(eligDate))}`);
  }
  
  // Check clinician mismatch — skip when Wrong Member ID is the root cause
  const eligClinician = (record.Clinician || '').trim();
  const claimClinician = claimInfo.claimClinician || '';
  if (eligClinician && claimClinician && eligClinician !== claimClinician && !wrongMemberID) {
    mismatches.push(`Different clinician: Claim has "${escapeHtml(claimClinician)}", but eligibility is for "${escapeHtml(eligClinician)}"`);
  }
  
  // Check package mismatch
  const eligPackage = (record['Package Name'] || '').trim();
  const claimPackage = claimInfo.claimPackage || '';
  if (eligPackage && claimPackage && !packageNamesMatch(claimPackage, eligPackage)) {
    // Normalize package names for display to show tier levels (e.g., "Daman Enhanced") instead of specific plan names (e.g., "Sahtak")
    const normalizedClaimPackage = normalizePackageNameForDisplay(claimPackage);
    const normalizedEligPackage = normalizePackageNameForDisplay(eligPackage);
    mismatches.push(`Registered under ${escapeHtml(normalizedClaimPackage)}, ${escapeHtml(normalizedEligPackage)} as per Eligibility`);
  }
  
  // Check status mismatch
  if (status.toLowerCase() !== 'eligible') {
    mismatches.push(`Status is not "Eligible": Currently "${escapeHtml(status)}"`);
  }
  
  // Display mismatch reasons if any
  if (mismatches.length > 0) {
    html += '<div class="alert alert-warning mb-3" style="padding: 10px; background-color: #fff3cd; border: 1px solid #ffc107; border-radius: 4px;">';
    html += '<strong>⚠️ Mismatch Reasons:</strong><ul style="margin: 8px 0 0 0; padding-left: 20px;">';
    mismatches.forEach(reason => {
      html += `<li>${reason}</li>`;
    });
    html += '</ul></div>';
  }

  const isAccepted = mismatches.length === 0;

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
          if (isAccepted) {
            if (DateHandler.isSameDay(claimDate, parsed)) rowClass = 'table-success';
          } else {
            if (!DateHandler.isSameDay(claimDate, parsed)) rowClass = 'table-danger';
          }
        }
      } else if (key === 'Card Number / DHA Member ID') {
        // Always mark red when the claim had a wrong member ID (EID fallback resolved it)
        if (wrongMemberID) rowClass = 'table-danger';
      } else if (key === 'Status') {
        if (isAccepted) rowClass = 'table-success';
        else if (status.toLowerCase() !== 'eligible') rowClass = 'table-danger';
      } else if (key === 'Clinician') {
        if (isAccepted && eligClinician && claimClinician) rowClass = 'table-success';
        else if (!isAccepted && eligClinician && claimClinician && eligClinician !== claimClinician && !wrongMemberID) rowClass = 'table-danger';
      } else if (key === 'Package Name') {
        if (isAccepted && eligPackage && claimPackage) rowClass = 'table-success';
        else if (!isAccepted && eligPackage && claimPackage && !packageNamesMatch(claimPackage, eligPackage)) rowClass = 'table-danger';
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
        const keyLow = key.toLowerCase();
        const isRangeStart = keyLow.startsWith('effective');
        const isRangeEnd = keyLow.startsWith('expir');
        if (isRangeStart) {
          // Claim date must be on or after the effective date
          if (isAccepted) {
            if (claimDate.getTime() >= parsed.getTime()) rowClass = 'table-success';
          } else {
            if (claimDate.getTime() < parsed.getTime()) rowClass = 'table-danger';
          }
        } else if (isRangeEnd) {
          // Claim date must be on or before the expiry date
          if (isAccepted) {
            if (claimDate.getTime() <= parsed.getTime()) rowClass = 'table-success';
          } else {
            if (claimDate.getTime() > parsed.getTime()) rowClass = 'table-danger';
          }
        } else {
          // Regular date field: must match exactly
          if (isAccepted) {
            if (DateHandler.isSameDay(claimDate, parsed)) rowClass = 'table-success';
          } else {
            if (!DateHandler.isSameDay(claimDate, parsed)) rowClass = 'table-danger';
          }
        }
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
  const exportData = invalidEntries.map((entry, index) => {
    // Convert date to Date object for Excel to recognize it as a date
    let exportDate = entry.encounterStart || '';
    if (exportDate) {
      const parsedDate = DateHandler.parse(exportDate);
      if (parsedDate) {
        exportDate = parsedDate;  // Keep as Date object for proper Excel formatting
      }
    }
    
    return {
      'File No': entry.fileNo || '',
      'Claim ID': entry.claimID || '',
      'Visit ID': entry.visitID || '',
      'Phy Lic': entry.clinician || '',
      'Date': exportDate,
      'Member ID': entry.rawMemberID || entry.memberID || '',
      'Clinician Name': entry.clinicianName || '',
      'Verdict': (entry.remarks || []).join('; '),
      'Opened By': entry.openedBy || '',
      'Price': entry.price ?? '',
      'Admitting DEPT': entry.department || '',
      'Pri. Plan Type': entry.packageName || entry.provider || '',
      'Facility ID': entry.facilityID || '',
      'File Name': entry.sourceFile || sourceFileName || ''
    };
  });
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
      sourceFileName = file.name; // Store source filename for export
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
                       `  • Combined report must have: "Pri. Claim No" and "Total Amount"\n\n` +
                       `These columns were not found in your file.\n` +
                       `Please verify you're uploading a valid Insta, Odoo, or Combined export.`;
      console.error(errorMsg);
      updateStatus('Error: Unknown report type - see console for details');
      alert(`Cannot process report file.\n\n` +
            `This file does not appear to be a valid Insta, Odoo, or Combined report.\n\n` +
            `Expected columns in original file:\n` +
            `  • Insta reports must have: "Pri. Claim No"\n` +
            `  • Odoo reports must have: "Pri. Claim ID"\n` +
            `  • Combined reports must have: "Pri. Claim No" and "Total Amount"\n\n` +
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

function onDiagnosticsToggle() {
  if (!diagnosticsStatus) return;
  const on = diagnosticsCheckbox && diagnosticsCheckbox.checked;
  diagnosticsStatus.textContent = on ? 'ON' : 'OFF';
  diagnosticsStatus.classList.toggle('active', on);
  
  // Update the global state
  showDiagnosticsButtons = on;
  
  // Re-render the results to show/hide diagnostics buttons
  if (window.lastValidationResults && window.lastValidationResults.length > 0) {
    const displayed = getDisplayedResultsFromStored(window.lastValidationResults);
    const eligMap = lastEligMap || new Map();
    renderResults(displayed, eligMap, window.lastValidationResults);
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
  diagnosticsCheckbox = document.getElementById('showDiagnosticsButtons');
  diagnosticsStatus = document.getElementById('diagnosticsStatus');
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
  if (diagnosticsCheckbox) {
    // Default to false (unchecked) to hide diagnostics buttons by default
    diagnosticsCheckbox.checked = false;
    diagnosticsCheckbox.addEventListener('change', onDiagnosticsToggle);
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
  if (diagnosticsStatus) onDiagnosticsToggle();
}

document.addEventListener('DOMContentLoaded', async () => {
  // Load eligibility matching configuration first
  await loadEligibilityMatchingConfig();
  
  initializeEventListeners();
  updateStatus('Ready to process files');
});
