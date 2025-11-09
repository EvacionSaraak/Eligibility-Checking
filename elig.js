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

/************************************
 * UTILITY / LOGGING HELPERS
 ************************************/
function updateStatus(message) {
  const status = document.getElementById("uploadStatus");
  if (status) status.textContent = message || 'Ready';
}

function logGroupCollapsed(title, fn) {
  try {
    console.groupCollapsed(title);
    fn();
    console.groupEnd();
  } catch (err) {
    console.error('Logging helper error:', err);
  }
}

function info(...args) { console.info('[elig] ', ...args); }
function warn(...args) { console.warn('[elig] ', ...args); }
function error(...args) { console.error('[elig] ', ...args); }

function updateProcessButtonState() {
  const processBtn = document.getElementById("processBtn");
  const exportInvalidBtn = document.getElementById("exportInvalidBtn");
  const hasEligibility = !!eligData;
  const hasReportData = !!reportData;
  if (processBtn) processBtn.disabled = !(hasEligibility && hasReportData);
  if (exportInvalidBtn) exportInvalidBtn.disabled = !(hasEligibility && hasReportData);
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
    const d = date.getDate().toString().padStart(2, '0'); // LOCAL date
    const m = (date.getMonth() + 1).toString().padStart(2, '0'); // LOCAL month
    const y = date.getFullYear();
    return `${d}/${m}/${y}`;
  },
  isSameDay: function(date1, date2) { // LOCAL comparison
    if (!date1 || !date2) return false;
    return date1.getFullYear() === date2.getFullYear() &&
           date1.getMonth() === date2.getMonth() &&
           date1.getDate() === date2.getDate();
  },
  _parseExcelDate: function(serial) {
    const utcDays = Math.floor(serial) - 25569;
    const ms = utcDays * 86400 * 1000;
    const date = new Date(ms);
    return new Date(date.getFullYear(), date.getMonth(), date.getDate()); // LOCAL
  },
  _parseStringDate: function(dateStr, preferMDY = false) {
    if (dateStr.includes(' ')) dateStr = dateStr.split(' ')[0];
    const dmyMdyMatch = dateStr.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
    if (dmyMdyMatch) {
      const part1 = parseInt(dmyMdyMatch[1], 10);
      const part2 = parseInt(dmyMdyMatch[2], 10);
      const year = parseInt(dmyMdyMatch[3], 10);

      if (part1 > 12 && part2 <= 12) {
        return new Date(year, part2 - 1, part1);
      } else if (part2 > 12 && part1 <= 12) {
        return new Date(year, part1 - 1, part2);
      } else {
        if (preferMDY) {
          return new Date(year, part1 - 1, part2);
        } else {
          return new Date(year, part2 - 1, part1);
        }
      }
    }
    const textMatch = dateStr.match(/^(\d{1,2})[\/\- ]([a-z]{3,})[\/\- ](\d{2,4})$/i);
    if (textMatch) {
      const monthIndex = MONTHS.indexOf(textMatch[2].toLowerCase().substr(0, 3));
      if (monthIndex >= 0) return new Date(textMatch[3], monthIndex, textMatch[1]);
    }
    const isoMatch = dateStr.match(/^(\d{4})[\/\-](\d{2})[\/\-](\d{2})$/);
    if (isoMatch) return new Date(isoMatch[1], isoMatch[2] - 1, isoMatch[3]);
    return null;
  }
};

/************************************
 * NORMALIZERS
 ************************************/
function normalizeMemberID(id) {
  if (!id) return '';
  return String(id).replace(/\s+/g, '').replace(/^0+/, '').toLowerCase();
}

function normalizeClinician(name) {
  if (!name) return '';
  return String(name).trim().toLowerCase().replace(/\s+/g, ' ');
}

function checkClinicianMatch(claimClinicians, eligClinician) {
  // if eligibility doesn't specify clinician or claim has none, consider it a match
  if (!eligClinician || !claimClinicians?.length) return true;
  const normElig = normalizeClinician(eligClinician);
  return claimClinicians.some(c => normalizeClinician(c) === normElig);
}

/************************************
 * HEADER DETECTION / DATA NORMALIZATION
 ************************************/
function normalizeReportData(rawData) {
  const isInsta = rawData[0]?.hasOwnProperty('Pri. Claim No');
  const isOdoo = rawData[0]?.hasOwnProperty('Pri. Claim ID');
  return rawData.map(row => {
    const rawMemberID =
      isInsta ? row['Pri. Patient Insurance Card No'] :
      isOdoo ? row['Pri. Member ID'] :
      row['PatientCardID'] || '';
    const memberID = normalizeMemberID(rawMemberID);
    return {
      claimID: isInsta ? row['Pri. Claim No'] || '' :
               isOdoo ? row['Pri. Claim ID'] || '' :
               row['ClaimID'] || '',
      rawMemberID, // <-- save original
      memberID,
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
 * Prepare Eligibility Map
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
      rawMemberID: rawID, // Added for leading-zero checks
      'Answered On': e['Answered On'],
      'Ordered On': e['Ordered On'],
      'Status': e['Status'],
      'Clinician': e['Clinician'],
      'Payer Name': e['Payer Name'],
      'Service Category': e['Service Category'],
      'Package Name': e['Package Name'],
      'Consultation Status': e['Consultation Status'] || '',
      'Department': e['Department'] || e['Clinic'] || ''
    };
    eligMap.get(memberID).push(eligRecord);
  });
  return eligMap;
}

/************************************
 * Find Eligibility for a Claim
 ************************************/
function findEligibilityForClaim(eligMap, claimDate, memberID, claimClinicians = []) {
  const normalizedID = normalizeMemberID(memberID);
  const eligList = eligMap.get(normalizedID) || [];

  if (!eligList.length) {
    logGroupCollapsed(`[Diagnostics] No eligibility records for ${memberID} (normalized: ${normalizedID})`, () => {
      warn('No eligibility records found in eligMap for this member ID.');
    });
    return null;
  }

  logGroupCollapsed(`[Diagnostics] Searching eligibilities for ${memberID} (normalized: ${normalizedID})`, () => {
    info('Claim date object:', claimDate, 'Formatted:', DateHandler.format(claimDate));
    info('Claim clinicians:', claimClinicians);
    eligList.forEach((elig, idx) => {
      info(`#${idx+1}`, {
        'Eligibility Request Number': elig['Eligibility Request Number'],
        'Answered On': elig['Answered On'],
        'Status': elig['Status'],
        'Clinician': elig['Clinician'],
        'Service Category': elig['Service Category'],
        'Package Name': elig['Package Name'],
        'Department': elig['Department']
      });
    });
  });

  for (const elig of eligList) {
    const eligDate = DateHandler.parse(elig["Answered On"]);
    if (!DateHandler.isSameDay(claimDate, eligDate)) {
      info('Skipping eligibility (date mismatch):', elig['Eligibility Request Number'], elig["Answered On"]);
      continue;
    }

    // Clinician matching
    const eligClinician = (elig.Clinician || '').trim();
    if (eligClinician && claimClinicians.length && !checkClinicianMatch(claimClinicians, eligClinician)) {
      info('Skipping eligibility (clinician mismatch):', elig['Eligibility Request Number'], 'elig clinician:', eligClinician);
      continue;
    }

    const serviceCategory = (elig['Service Category'] || '').trim();
    const consultationStatus = (elig['Consultation Status'] || '').trim().toLowerCase();
    const department = (elig.Department || elig.Clinic || '').toLowerCase();
    const categoryCheck = isServiceCategoryValid(serviceCategory, consultationStatus, department);
    if (!categoryCheck.valid) {
      info('Skipping eligibility (category/package mismatch):', elig['Eligibility Request Number'], categoryCheck.reason || '');
      continue;
    }

    if ((elig.Status || '').trim().toLowerCase() !== 'eligible') {
      info('Skipping eligibility (status not eligible):', elig['Eligibility Request Number'], elig.Status);
      continue;
    }

    info('Eligibility matched:', elig['Eligibility Request Number']);
    return elig;
  }

  warn(`No matching eligibility passed all checks for member ${memberID} on ${DateHandler.format(claimDate)}`);
  return null;
}

/************************************
 * Service Category Validation
 ************************************/
function isServiceCategoryValid(serviceCategory, consultationStatus, rawPackage) {
  if (!serviceCategory) return { valid: true };

  const category = serviceCategory.trim().toLowerCase();
  const pkgRaw = rawPackage || '';
  const pkg = pkgRaw.toLowerCase();

  if (category === 'consultation' && consultationStatus === 'elective') {
    const disallowed = ['dental', 'physio', 'diet', 'occupational', 'speech'];
    if (disallowed.some(term => pkg.includes(term))) {
      return {
        valid: false,
        reason: `Consultation (Elective) cannot include restricted types. Found: "${pkgRaw}"`
      };
    }
    return { valid: true };
  }

  const allowedKeywords = SERVICE_PACKAGE_RULES[category];
  if (allowedKeywords && allowedKeywords.length > 0) {
    if (pkg && !allowedKeywords.some(keyword => pkg.includes(keyword.toLowerCase()))) {
      return {
        valid: false,
        reason: `${serviceCategory} requires related package. Found: "${pkgRaw}"`
      };
    }
  }

  return { valid: true };
}

/************************************
 * Validate Report Claims
 ************************************/
function validateReportClaims(reportData, eligMap) {
  if (!reportData) return [];

  return reportData.map(row => {
    if (!row.claimID) return null;

    const memberID = row.memberID;
    const claimDateRaw = row.claimDate;
    const claimDate = DateHandler.parse(claimDateRaw, { preferMDY: lastReportWasCSV });
    const formattedDate = DateHandler.format(claimDate);

    // VVIP bypass
    const isVVIP = String(memberID || '').toLowerCase().startsWith('(vvip)');
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

    // Leading zero check using rawMemberID
    const hasLeadingZero = String(row.rawMemberID || '').match(/^0+\d+$/);
    const eligibility = findEligibilityForClaim(eligMap, claimDate, memberID, [row.clinician].filter(Boolean));
    let status = 'invalid';
    const remarks = [];

    if (!eligibility) {
      remarks.push(`No matching eligibility found for ${memberID} on ${formattedDate}`);
      warn('Eligibility not found:', { memberID, claimDateRaw, formattedDate, clinician: row.clinician });
    } else if ((eligibility.Status || '').trim().toLowerCase() !== 'eligible') {
      remarks.push(`Eligibility status: ${eligibility.Status}`);
      info('Eligibility status not eligible:', eligibility.Status);
    } else {
      const serviceCategory = eligibility['Service Category']?.trim() || '';
      const consultationStatus = eligibility['Consultation Status']?.trim()?.toLowerCase() || '';
      const matchesCategory = isServiceCategoryValid(serviceCategory, consultationStatus, row.department || row.clinic).valid;
      if (!matchesCategory) {
        remarks.push(`Invalid for category: ${serviceCategory}, department: ${row.department || row.clinic}`);
        info('Service category mismatch for claim:', serviceCategory, row.department || row.clinic);
      } else if (!hasLeadingZero) {
        status = 'valid';
      }
    }

    if (hasLeadingZero) remarks.push('Member ID has a leading zero; claim marked as invalid.');

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
 * PARSING XLSX/CSV AND EXPORT
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

function exportInvalidEntries(results) {
  // When the results passed are already filtered (by Daman/Thiqa toggle during processing),
  // the export will respect that filtered set.
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
 * RENDERING RESULTS
 ************************************/
function renderResults(results, eligMap) {
  const resultsContainer = document.getElementById("results");
  if (!resultsContainer) {
    warn('#results container not found');
    return;
  }
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

  // This script is report-only, so xmlRadio is not used here. Keep column rendering consistent.
  const thead = document.createElement('thead');
  thead.innerHTML = `
    <tr>
      <th>Claim ID</th>
      <th>Member ID</th>
      <th>Encounter Date</th>
      <th>Package</th><th>Provider</th>
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
    if (!result.memberID || String(result.memberID).trim() === '') return;

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

    const statusBadge = result.status 
      ? `<span class="status-badge ${result.status.toLowerCase() === 'eligible' ? 'eligible' : 'ineligible'}">${result.status}</span>`
      : '';

    const remarksHTML = result.remarks && result.remarks.length > 0
      ? result.remarks.map(r => `<div>${r}</div>`).join('')
      : '<div class="source-note">No remarks</div>';

    let detailsCell = '<div class="source-note">N/A</div>';
    if (result.fullEligibilityRecord?.['Eligibility Request Number']) {
      detailsCell = `<div class="source-note">${result.fullEligibilityRecord['Eligibility Request Number']}</div>`;
    } else if (eligMap && eligMap.has && eligMap.has(result.memberID)) {
      detailsCell = `<div class="source-note">Elig records: ${eligMap.get(result.memberID).length}</div>`;
    }

    row.innerHTML = `
      <td>${result.claimID}</td>
      <td>${result.memberID}</td>
      <td>${result.encounterStart}</td>
      <td class="description-col">${result.packageName}</td><td class="description-col">${result.provider}</td>
      <td class="description-col">${result.clinician}</td>
      <td class="description-col">${result.serviceCategory}</td>
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
}

/************************************
 * DOM READY/EVENTS
 ************************************/
document.addEventListener("DOMContentLoaded", () => {
  // ===== Elements =====
  const checkbox = document.getElementById("filterDamanThiqa");
  const status = document.getElementById("filterStatus");
  const reportInput = document.getElementById("reportFileInput");
  const eligInput = document.getElementById("eligibilityFileInput");
  const processBtn = document.getElementById("processBtn");
  const exportInvalidBtn = document.getElementById("exportInvalidBtn");
  const resultsContainer = document.getElementById("results");

  // ===== Checkbox status display =====
  if (checkbox && status) {
    const updateStatusToggle = () => {
      if (checkbox.checked) {
        status.textContent = "ON";
        status.classList.add("active");
      } else {
        status.textContent = "OFF";
        status.classList.remove("active");
      }
    };
    checkbox.addEventListener("change", updateStatusToggle);
    updateStatusToggle();
  }

  // ===== File input listeners =====
  if (reportInput) {
    reportInput.addEventListener("change", async (e) => {
      const file = e.target.files[0];
      if (!file) return;
      try {
        updateStatus("Loading report file...");
        lastReportWasCSV = file.name.toLowerCase().endsWith(".csv");
        const rawData = lastReportWasCSV
          ? await parseCsvFile(file)
          : await parseExcelFile(file);
        reportData = normalizeReportData(rawData).filter(
          (r) => r.claimID && String(r.claimID).trim() !== ""
        );
        info(
          "--- Report File Uploaded ---",
          file.name,
          `rows: ${reportData.length}`
        );
        updateStatus(`Loaded ${reportData.length} report rows`);
        updateProcessButtonState();
      } catch (err) {
        updateStatus("Error loading report file");
        error("Report file error:", err);
        if (processBtn) processBtn.disabled = true;
        if (exportInvalidBtn) exportInvalidBtn.disabled = true;
      }
    });
  }

  if (eligInput) {
    eligInput.addEventListener("change", async (e) => {
      const file = e.target.files[0];
      if (!file) return;
      try {
        updateStatus("Loading eligibility file...");
        eligData = await parseExcelFile(file);
        info(
          "--- Eligibility File Uploaded ---",
          file.name,
          `rows: ${eligData.length}`
        );
        updateStatus(`Loaded ${eligData.length} eligibility records`);
        updateProcessButtonState();
      } catch (err) {
        updateStatus("Error loading eligibility file");
        error("Eligibility file error:", err);
        if (processBtn) processBtn.disabled = true;
        if (exportInvalidBtn) exportInvalidBtn.disabled = true;
      }
    });
  }

  // ===== Process button =====
  if (processBtn) {
    processBtn.addEventListener("click", async () => {
      if (!eligData) {
        updateStatus("Missing eligibility file");
        alert("Please upload eligibility file first");
        return;
      }
      if (!reportData) {
        updateStatus("Missing report file");
        alert("Please upload patient report file");
        return;
      }

      try {
        updateStatus("Processing...");
        usedEligibilities.clear();
        const eligMap = prepareEligibilityMap(eligData);
        info("--- Eligibility Map Keys ---", Array.from(eligMap.keys()));

        let results = validateReportClaims(reportData, eligMap);
        info("--- Raw Validation Results ---", `count: ${results.length}`);

        // ===== Safe Daman/Thiqa filter =====
        const filterOn = checkbox?.checked ?? false;

        if (filterOn) {
          logGroupCollapsed("[Filter] Applying Daman/Thiqa filter", () => {
            info(
              'Filter ON — only claims with provider/package/payer that includes "daman" or "thiqa" will be kept.'
            );
          });

          results = results.filter((r) => {
            const provider = (
              r.provider ||
              r.insuranceCompany ||
              r.packageName ||
              r["Payer Name"] ||
              r["Insurance Company"] ||
              ""
            )
              .toString()
              .toLowerCase();
            return provider.includes("daman") || provider.includes("thiqa");
          });

          info(
            `[Filter] ${results.length} claims remain after Daman/Thiqa filter.`
          );
        } else {
          info(
            "[Filter] Daman/Thiqa filter is OFF — all validated claims will be shown."
          );
        }

        // ===== Save results for export =====
        window.lastValidationResults = results;
        renderResults(results, eligMap);
        updateStatus(`Processed ${results.length} claims`);

        if (results.length === 0) {
          warn(
            "No claims processed! Check input files, normalization, and provider fields."
          );
        }
      } catch (err) {
        updateStatus("Processing failed");
        if (resultsContainer)
          resultsContainer.innerHTML = `<div class="error">${err.message}</div>`;
        error("Processing error:", err);
      }
    });
  }

  // ===== Export button =====
  if (exportInvalidBtn) {
    exportInvalidBtn.addEventListener("click", () => {
      if (!window.lastValidationResults) {
        alert("Please run the validation first.");
        return;
      }
      exportInvalidEntries(window.lastValidationResults);
    });
  }

  // ===== Initialize =====
  updateProcessButtonState();
  updateStatus("Ready to process files");
  info("--- Eligibility Checker Initialized (report-only) ---");
});
