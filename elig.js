/*******************************
 * GLOBAL VARIABLES & CONSTANTS *
 *******************************/
const SERVICE_PACKAGE_RULES = {
  'Dental Services': ['dental', 'orthodontic'],
  'Physiotherapy': ['physio'],
  'Other OP Services': ['physio', 'diet', 'occupational', 'speech'],
  'Consultation': []  // Special handling below
};
const DATE_KEYS = ['Date', 'On'];
const MONTHS = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"];

// Create a normalized version of the rules for case-insensitive lookup
const NORMALIZED_SERVICE_PACKAGE_RULES = {};
Object.keys(SERVICE_PACKAGE_RULES).forEach(k => {
  NORMALIZED_SERVICE_PACKAGE_RULES[k.trim().toLowerCase()] = SERVICE_PACKAGE_RULES[k];
});

// Application state
let xlsData = null;
let eligData = null;
const usedEligibilities = new Set();

// DOM Elements
const reportInput = document.getElementById("reportFileInput");
const eligInput = document.getElementById("eligibilityFileInput");
const processBtn = document.getElementById("processBtn");
const exportInvalidBtn = document.getElementById("exportInvalidBtn");
const status = document.getElementById("uploadStatus");
const resultsContainer = document.getElementById("results");
const reportGroup = document.getElementById("reportInputGroup");
const xlsRadio = document.querySelector('input[name="reportSource"][value="xls"]');

/*************************
 * RADIO BUTTON HANDLING *
 *************************/
function initializeRadioButtons() {
  if (xlsRadio) {
    xlsRadio.addEventListener('change', handleReportSourceChange);
    handleReportSourceChange();
  }
}

function handleReportSourceChange() {
  // This function existed in original code context; keep it defensive
  if (!reportGroup) return;
  const useXls = xlsRadio && xlsRadio.checked;
  reportGroup.style.display = useXls ? 'block' : 'block';
}

/*************************
 * DATE HANDLING UTILITIES *
 *************************/
let lastReportWasCSV = false;
const DateHandler = {
  parse: function(input, options = {}) {
    const preferMDY = !!options.preferMDY;
    if (!input && input !== 0) return null;
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
    // Excel serial -> UTC midnight
    const utcDays = Math.floor(serial) - 25569;
    const ms = utcDays * 86400 * 1000;
    const date = new Date(ms);
    return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
  },

  // Always parse string dates as UTC and handle two-digit years
  _parseStringDate: function(dateStr, preferMDY = false) {
    if (!dateStr || typeof dateStr !== 'string') return null;

    if (dateStr.includes(' ')) {
      dateStr = dateStr.split(' ')[0];
    }

    // Matches DD/MM/YYYY or MM/DD/YYYY (ambiguous)
    const dmyMdyMatch = dateStr.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
    if (dmyMdyMatch) {
      const part1 = parseInt(dmyMdyMatch[1], 10);
      const part2 = parseInt(dmyMdyMatch[2], 10);
      let year = parseInt(dmyMdyMatch[3], 10);
      if (year < 100) year += 2000;

      if (part1 > 12 && part2 <= 12) {
        // dmy
        return new Date(Date.UTC(year, part2 - 1, part1));
      } else if (part2 > 12 && part1 <= 12) {
        // mdy
        return new Date(Date.UTC(year, part1 - 1, part2));
      } else {
        if (preferMDY) {
          return new Date(Date.UTC(year, part1 - 1, part2));
        } else {
          return new Date(Date.UTC(year, part2 - 1, part1));
        }
      }
    }

    // Matches 30-Jun-2025 or 30 Jun 2025 (or 30-Jun-25)
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
    if (isoMatch) {
      const y = parseInt(isoMatch[1], 10);
      const mo = parseInt(isoMatch[2], 10);
      const d = parseInt(isoMatch[3], 10);
      return new Date(Date.UTC(y, mo - 1, d));
    }
    return null;
  }
};

/*****************************
 * DATA NORMALIZATION FUNCTIONS *
 *****************************/
function normalizeMemberID(id) {
  if (id === null || id === undefined) return '';
  return String(id).trim().replace(/^0+/, '');
}

function normalizeClinician(name) {
  if (!name) return '';
  return name.trim().toLowerCase().replace(/\s+/g, ' ');
}

/*******************************
 * ELIGIBILITY MATCHING FUNCTIONS *
 *******************************/
function prepareEligibilityMap(eligData) {
  const eligMap = new Map();

  eligData.forEach(e => {
    const rawID =
      e['Card Number / DHA Member ID'] ||
      e['Card Number'] ||
      e['_5'] ||
      e['MemberID'] ||
      e['Member ID'] ||
      e['Patient Insurance Card No'] ||
      e['PatientCardID'];

    if (!rawID && rawID !== 0) return;

    const memberID = normalizeMemberID(rawID); // normalized key

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
      'Package Name': e['Package Name'],
      'Department': e['Department'] || e['Clinic'] || ''
    };

    eligMap.get(memberID).push(eligRecord);
  });

  return eligMap;
}

function findEligibilityForClaim(eligMap, claimDate, memberID, claimClinicians = []) {
  const normalizedID = normalizeMemberID(memberID);
  const eligList = eligMap.get(normalizedID) || [];

  if (!eligList.length) return null;

  console.log(`[Diagnostics] Searching eligibilities for member "${memberID}" (normalized: "${normalizedID}")`);
  console.log(`[Diagnostics] Claim date: ${claimDate} (${DateHandler.format(claimDate)}), Claim clinicians: ${JSON.stringify(claimClinicians)}`);

  // Ensure claimClinicians is an array of non-empty strings
  const claimCliniciansFiltered = Array.isArray(claimClinicians)
    ? claimClinicians.filter(Boolean).map(c => (typeof c === 'string' ? c.trim() : String(c).trim()))
    : [];

  for (const elig of eligList) {
    console.log(`[Diagnostics] Checking eligibility ${elig["Eligibility Request Number"] || "(unknown)"}:`);

    const eligDate = DateHandler.parse(elig["Answered On"]);
    if (!DateHandler.isSameDay(claimDate, eligDate)) {
      console.log(`  ❌ Date mismatch: claim ${DateHandler.format(claimDate)} vs elig ${DateHandler.format(eligDate)}`);
      continue;
    }

    const eligClinician = (elig.Clinician || '').trim();
    if (eligClinician && claimCliniciansFiltered.length && !checkClinicianMatch(claimCliniciansFiltered, eligClinician)) {
      console.log(`  ❌ Clinician mismatch: claim clinicians ${JSON.stringify(claimCliniciansFiltered)} vs elig clinician "${eligClinician}"`);
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
    // Mark as used
    if (elig["Eligibility Request Number"]) usedEligibilities.add(elig["Eligibility Request Number"]);
    return elig;
  }

  console.log(`[Diagnostics] No matching eligibility passed all checks for member "${memberID}"`);
  return null;
}

function checkClinicianMatch(claimClinicians, eligClinician) {
  if (!eligClinician || !claimClinicians?.length) return true;
  const normElig = normalizeClinician(eligClinician);
  return claimClinicians.some(c => normalizeClinician(c) === normElig);
}

/************************
 * VALIDATION FUNCTIONS *
 ************************/
function isServiceCategoryValid(serviceCategory, consultationStatus, rawPackage) {
  if (!serviceCategory) return { valid: true };

  const categoryLower = serviceCategory.trim().toLowerCase();
  const pkgRaw = rawPackage || '';
  const pkg = pkgRaw.toLowerCase();

  // Consultation rule: allow anything EXCEPT the restricted types when elective
  if (categoryLower === 'consultation' && consultationStatus?.toLowerCase() === 'elective') {
    const disallowed = ['dental', 'physio', 'diet', 'occupational', 'speech'];
    if (disallowed.some(term => pkg.includes(term))) {
      return {
        valid: false,
        reason: `Consultation (Elective) cannot include restricted service types. Found: "${pkgRaw}"`
      };
    }
    return { valid: true };
  }

  // Use normalized rule map
  const allowedKeywords = NORMALIZED_SERVICE_PACKAGE_RULES[categoryLower];
  if (allowedKeywords && allowedKeywords.length > 0) {
    if (pkg && !allowedKeywords.some(keyword => pkg.includes(keyword))) {
      return {
        valid: false,
        reason: `${serviceCategory} category requires related package. Found: "${pkgRaw}"`
      };
    }
  }

  return { valid: true };
}

function validateReportClaims(reportData, eligMap) {
  console.log(`Validating ${reportData.length} report rows`);

  const results = reportData.map(row => {
    if (!row.claimID || String(row.claimID).trim() === '') return null;

    const memberID = String(row.memberID || '').trim();
    const claimDateRaw = row.claimDate;
    const claimDate = DateHandler.parse(claimDateRaw, { preferMDY: lastReportWasCSV });
    const formattedDate = DateHandler.format(claimDate);

    // VVIP IDs: mark as valid with a special remark
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
        fullEligibilityRecord: null
      };
    }

    // Check for leading zero in original memberID
    const hasLeadingZero = memberID.match(/^0+\d+$/);

    // Build clinicians array for lookup
    const claimClinicians = [];
    if (row.clinician) claimClinicians.push(row.clinician);

    // Proceed with normal eligibility lookup
    const eligibility = findEligibilityForClaim(eligMap, claimDate, memberID, claimClinicians);
    let status = 'invalid';
    const remarks = [];
    const department = (row.department || row.clinic || '').toLowerCase();

    // If leading zero, mark invalid and add remark
    if (hasLeadingZero) {
      remarks.push('Member ID has a leading zero; claim marked as invalid.');
    }

    if (!eligibility) {
      remarks.push(`No matching eligibility found for ${memberID} on ${formattedDate}`);
      logNoEligibilityMatch(
        'REPORT',
        {
          claimID: row.claimID,
          memberID,
          claimDateRaw,
          department: row.department || row.clinic,
          clinician: row.clinician,
          packageName: row.packageName
        },
        memberID,
        claimDate,
        claimClinicians,
        eligMap
      );
    } else if ((eligibility.Status || '').toLowerCase() !== 'eligible') {
      remarks.push(`Eligibility status: ${eligibility.Status}`);
    } else {
      const serviceCategory = eligibility['Service Category']?.trim() || '';
      const consultationStatus = eligibility['Consultation Status']?.trim()?.toLowerCase() || '';
      const matchesCategory = isServiceCategoryValid(serviceCategory, consultationStatus, department).valid;

      if (!matchesCategory) {
        remarks.push(`Invalid for category: ${serviceCategory}, department: ${row.department || row.clinic}`);
      } else if (!hasLeadingZero) {
        // Only mark as valid if there is no leading zero
        status = 'valid';
      }
      // If hasLeadingZero, status remains 'invalid'
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
  });

  return results.filter(r => r);
}

// --- Helper logger ---
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
 * FILE PARSING FUNCTIONS *
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
          const currentRow = (allRows[headerRow] || []).map(c => String(c).trim());

          // Skip likely title rows
          if (isLikelyTitleRow(currentRow)) {
            headerRow++;
            continue;
          }

          // Check for known headers
          if (currentRow.some(cell => cell.toLowerCase().includes('pri. claim no')) ||
              currentRow.some(cell => cell.toLowerCase().includes('pri. claim id')) ||
              currentRow.some(cell => cell.toLowerCase().includes('card number / dha member id')) ||
              currentRow.some(cell => cell.toLowerCase().includes('claimid')) ) {
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
        const headers = (allRows[headerRow] || []).map(h => String(h).trim());
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

async function parseCsvFile(file) {
  console.log(`Parsing CSV file: ${file.name}`);

  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = function(e) {
      try {
        const text = e.target.result;
        const workbook = XLSX.read(text, { type: 'string' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const allRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

        // Dynamically detect header row by scanning first 10 rows
        let headerRowIndex = -1;
        for (let i = 0; i < Math.min(10, allRows.length); i++) {
          const row = allRows[i];
          if (!row) continue;
          const joined = row.join(',').toLowerCase();
          if (joined.includes('pri. claim no') || joined.includes('claimid') || joined.includes('claim id') || joined.includes('pri. claim id')) {
            headerRowIndex = i;
            break;
          }
        }

        // If not found, fallback to first row that has >=3 non-empty cells
        if (headerRowIndex === -1) {
          for (let i = 0; i < Math.min(10, allRows.length); i++) {
            const row = allRows[i] || [];
            const nonEmpty = row.filter(c => String(c).trim() !== '').length;
            if (nonEmpty >= 3) {
              headerRowIndex = i;
              break;
            }
          }
        }

        if (headerRowIndex === -1) throw new Error("Could not detect header row in CSV");

        const headers = allRows[headerRowIndex];
        const dataRows = allRows.slice(headerRowIndex + 1);

        console.log(`Detected header at row ${headerRowIndex + 1}:`, headers);

        const rawParsed = dataRows.map(row => {
          const obj = {};
          headers.forEach((header, index) => {
            obj[header] = row[index] || '';
          });
          return obj;
        });

        // Deduplicate based on claim ID
        const seen = new Set();
        const uniqueRows = [];

        const claimIdHeader = headers.find(h =>
          h && h.toString().toLowerCase().replace(/\s+/g, '') === 'claimid'
        ) || headers.find(h => h && h.toString().toLowerCase().includes('claim'));

        if (!claimIdHeader) throw new Error("Could not find a Claim ID column");

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

function normalizeReportData(rawData) {
  if (!Array.isArray(rawData)) return [];

  // Check if data is from InstaHMS (has 'Pri. Claim No' header)
  const isInsta = rawData[0]?.hasOwnProperty('Pri. Claim No');
  const isOdoo = rawData[0]?.hasOwnProperty('Pri. Claim ID');

  return rawData.map(row => {
    if (isInsta) {
      // InstaHMS report format
      return {
        claimID: row['Pri. Claim No'] || '',
        memberID: row['Pri. Patient Insurance Card No'] || '',
        claimDate: row['Encounter Date'] || '',
        clinician: row['Clinician License'] || '',
        department: row['Department'] || '',
        packageName: row['Pri. Payer Name'] || '', // shown in table as "Package"
        insuranceCompany: row['Pri. Payer Name'] || '',
        claimStatus: row['Codification Status'] || ''
      };
    } else if (isOdoo) {
      // Odoo report format
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
      // ClinicPro / generic report format
      return {
        claimID: row['ClaimID'] || row['Pri. Claim No'] || '',
        memberID: row['PatientCardID'] || row['Patient Insurance Card No'] || '',
        claimDate: row['ClaimDate'] || row['Encounter Date'] || '',
        clinician: row['Clinician License'] || row['Clinician'] || '',
        packageName: row['Insurance Company'] || '', // shown in table as "Package"
        insuranceCompany: row['Insurance Company'] || '',
        department: row['Clinic'] || row['Department'] || '',
        claimStatus: row['VisitStatus'] || row['Codification Status'] || ''
      };
    }
  });
}

/********************
 * UI RENDERING FUNCTIONS *
 ********************/
function renderResults(results, eligMap) {
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
    // Skip rows where Member ID is missing/empty
    if (!result.memberID || result.memberID.trim() === '') return;

    // Ignore claims whose status is "Not Seen"
    const statusToCheck = (result.claimStatus || result.status || result.fullEligibilityRecord?.Status || '')
      .toString()
      .trim()
      .toLowerCase();

    if (statusToCheck === 'not seen') return;

    // Count statuses safely
    if (result.finalStatus && statusCounts.hasOwnProperty(result.finalStatus)) {
      statusCounts[result.finalStatus]++;
    }

    const row = document.createElement('tr');
    row.className = result.finalStatus;

    const statusBadge = result.status 
      ? `<span class="status-badge ${result.status.toLowerCase() === 'eligible' ? 'eligible' : 'ineligible'}">${result.status}</span>`
      : '';

    const remarksHTML = result.remarks && result.remarks.length > 0
      ? result.remarks.map(r => `<div>${r}</div>`).join('')
      : '<div class="source-note">No remarks</div>';

    // Details cell: either a single eligibility button or a 'View All' button
    let detailsCell = '<div class="source-note">N/A</div>';
    if (result.fullEligibilityRecord?.['Eligibility Request Number']) {
      detailsCell = `<button class="details-btn eligibility-details" data-index="${index}" aria-label="Eligibility details">${escapeHtml(result.fullEligibilityRecord['Eligibility Request Number'])}</button>`;
    } else {
      // If there are eligibilities in map for this member, show View All
      const normId = normalizeMemberID(result.memberID);
      if (eligMap && typeof eligMap.has === 'function' && eligMap.has(normId)) {
        detailsCell = `<button class="details-btn show-all-eligibilities" data-member="${escapeAttr(result.memberID)}">View All</button>`;
      }
    }

    row.innerHTML = `
      <td>${escapeHtml(result.claimID)}</td>
      <td>${escapeHtml(result.memberID)}</td>
      <td>${escapeHtml(result.encounterStart)}</td>
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

// Simple HTML escape helpers for safety in injected strings
function escapeHtml(s) {
  if (s === null || s === undefined) return '';
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#x27;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}
function escapeAttr(s) {
  return escapeHtml(s).replace(/"/g, '&quot;');
}

function initEligibilityModal(results, eligMap) {
  // Ensure modal exists
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
    document.getElementById("modalOverlay").onclick = function(e) {
      if (e.target.id === "modalOverlay") hideModal();
    };
  }

  // Event delegation for details buttons
  resultsContainer.querySelectorAll('.details-btn').forEach(btn => {
    btn.onclick = function() {
      // single eligibility button
      if (this.classList.contains('eligibility-details')) {
        const index = parseInt(this.dataset.index, 10);
        const result = results[index];
        if (!result?.fullEligibilityRecord) return;
        const html = formatEligibilityDetails(result.fullEligibilityRecord, result.memberID);
        document.getElementById("modalTable").innerHTML = html;
        document.getElementById("modalOverlay").style.display = "block";
      } else if (this.classList.contains('show-all-eligibilities')) {
        const member = this.dataset.member;
        const normalizedID = normalizeMemberID(member);
        const eligList = eligMap.get(normalizedID) || [];
        if (!eligList.length) {
          document.getElementById("modalTable").innerHTML = `<div>No eligibilities found for ${escapeHtml(member)}</div>`;
          document.getElementById("modalOverlay").style.display = "block";
          return;
        }
        // Build a table with all eligibilities
        let html = `<h3>All Eligibilities for ${escapeHtml(member)}</h3>`;
        html += '<div style="overflow-x:auto;"><table style="width:100%;border-collapse:collapse;">';
        html += '<thead><tr><th style="text-align:left;padding:6px;border-bottom:1px solid #ccc;">#</th><th style="text-align:left;padding:6px;border-bottom:1px solid #ccc;">Eligibility Request Number</th><th style="text-align:left;padding:6px;border-bottom:1px solid #ccc;">Answered On</th><th style="text-align:left;padding:6px;border-bottom:1px solid #ccc;">Status</th><th style="text-align:left;padding:6px;border-bottom:1px solid #ccc;">Clinician</th><th style="text-align:left;padding:6px;border-bottom:1px solid #ccc;">Service Category</th><th style="text-align:left;padding:6px;border-bottom:1px solid #ccc;">Package Name</th></tr></thead><tbody>';
        eligList.forEach((rec, idx) => {
          html += '<tr>';
          html += `<td style="padding:6px;border-bottom:1px solid #eee;">${idx+1}</td>`;
          html += `<td style="padding:6px;border-bottom:1px solid #eee;">${escapeHtml(rec['Eligibility Request Number'] || '')}</td>`;
          html += `<td style="padding:6px;border-bottom:1px solid #eee;">${escapeHtml(rec['Answered On'] || rec['Ordered On'] || '')}</td>`;
          html += `<td style="padding:6px;border-bottom:1px solid #eee;">${escapeHtml(rec['Status'] || '')}</td>`;
          html += `<td style="padding:6px;border-bottom:1px solid #eee;">${escapeHtml(rec['Clinician'] || '')}</td>`;
          html += `<td style="padding:6px;border-bottom:1px solid #eee;">${escapeHtml(rec['Service Category'] || '')}</td>`;
          html += `<td style="padding:6px;border-bottom:1px solid #eee;">${escapeHtml(rec['Package Name'] || '')}</td>`;
          html += '</tr>';
        });
        html += '</tbody></table></div>';
        document.getElementById("modalTable").innerHTML = html;
        document.getElementById("modalOverlay").style.display = "block";
      }
    };
  });
}

function hideModal() {
  const overlay = document.getElementById("modalOverlay");
  if (overlay) overlay.style.display = "none";
}

function formatEligibilityDetails(record, memberID) {
  if (!record) return '<div>No details available</div>';

  let html = `
    <div class="form-row">
      <strong>Member:</strong> ${escapeHtml(memberID)}
      <span class="status-badge ${((record.Status||'').toLowerCase() === 'eligible') ? 'eligible' : 'ineligible'}" style="margin-left:10px;">
        ${escapeHtml(record.Status || '')}
      </span>
    </div>
    <table class="eligibility-details" style="width:100%;border-collapse:collapse;margin-top:12px;">
      <tbody>
  `;

  Object.entries(record).forEach(([key, value]) => {
    if ((value === null || value === undefined || value === '') && value !== 0) return;

    let displayed = value;
    // Format dates
    if (DATE_KEYS.some(k => key.includes(k)) || key.toLowerCase().includes('answered') || key.toLowerCase().includes('ordered')) {
      const parsed = DateHandler.parse(value);
      displayed = parsed ? DateHandler.format(parsed) : value;
      displayed = escapeHtml(displayed);
    } else {
      displayed = escapeHtml(displayed);
    }

    html += `
      <tr>
        <th style="text-align:left;padding:6px;border-bottom:1px solid #eee;width:30%">${escapeHtml(key)}</th>
        <td style="padding:6px;border-bottom:1px solid #eee;">${displayed}</td>
      </tr>
    `;
  });

  html += `
      </tbody>
    </table>
  `;

  return html;
}

function updateStatus(message) {
  if (status) status.textContent = message || 'Ready';
}

function updateProcessButtonState() {
  const hasEligibility = !!eligData;
  const hasReportData = !!xlsData;
  if (processBtn) processBtn.disabled = !hasEligibility || !hasReportData;
  if (exportInvalidBtn) exportInvalidBtn.disabled = !hasEligibility || !hasReportData;
}

/************************
 * EXPORT FUNCTIONALITY *
 ************************/
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
    'Remarks': (entry.remarks || []).join('; ')
  }));

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(exportData);

  XLSX.utils.book_append_sheet(wb, ws, 'Invalid Claims');

  XLSX.writeFile(wb, `invalid_claims_${new Date().toISOString().slice(0,10)}.xlsx`);
}

/********************
 * EVENT HANDLERS *
 ********************/
async function handleFileUpload(event, type) {
  const file = event.target.files[0];
  if (!file) return;

  try {
    updateStatus(`Loading ${type} file...`);

    if (type === 'eligibility') {
      eligData = await parseExcelFile(file);
      updateStatus(`Loaded ${eligData.length} eligibility records`);
      lastReportWasCSV = false;
    }
    else {
      lastReportWasCSV = file.name.toLowerCase().endsWith('.csv');

      const rawData = lastReportWasCSV
        ? await parseCsvFile(file)
        : (file.name.toLowerCase().endsWith('.csv') ? await parseCsvFile(file) : await parseExcelFile(file));

      xlsData = normalizeReportData(rawData).filter(r => {
        return r.claimID !== null && r.claimID !== undefined && String(r.claimID).trim() !== '';
      });
      console.log(xlsData);
      updateStatus(`Loaded ${xlsData.length} report rows`);
    }

    updateProcessButtonState();
  } catch (error) {
    console.error(`${type} file error:`, error);
    updateStatus(`Error loading ${type} file`);
  }
}

async function handleProcessClick() {
  if (!eligData) {
    updateStatus('Error: Missing eligibility file');
    alert('Please upload eligibility file first');
    return;
  }

  if (!xlsData || xlsData.length === 0) {
    updateStatus('Error: Missing report data');
    alert('Please upload a report file first');
    return;
  }

  try {
    updateStatus('Processing...');
    usedEligibilities.clear();
    const eligMap = prepareEligibilityMap(eligData);
    const filteredResults = validateReportClaims(xlsData, eligMap);
    window.lastValidationResults = filteredResults;
    renderResults(filteredResults, eligMap);

    updateStatus(`Processed ${filteredResults.length} claims successfully`);
  } catch (error) {
    console.error('Processing error:', error);
    updateStatus('Processing failed');
    if (resultsContainer) resultsContainer.innerHTML = `<div class="error">${escapeHtml(error.message || String(error))}</div>`;
  }
}

function handleExportInvalidClick() {
  if (!window.lastValidationResults) {
    alert('Please run the validation first.');
    return;
  }
  exportInvalidEntries(window.lastValidationResults);
}

/********************
 * INITIALIZATION *
 ********************/
function initializeEventListeners() {
  if (eligInput) eligInput.addEventListener('change', (e) => handleFileUpload(e, 'eligibility'));
  if (reportInput) reportInput.addEventListener('change', (e) => handleFileUpload(e, 'report'));
  if (processBtn) processBtn.addEventListener('click', handleProcessClick);
  if (exportInvalidBtn) exportInvalidBtn.addEventListener('click', handleExportInvalidClick);
  initializeRadioButtons();
}

document.addEventListener('DOMContentLoaded', () => {
  initializeEventListeners();
  updateStatus('Ready to process files');
});
