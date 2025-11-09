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
  return String(id).trim().replace(/^0+/, '');
}

function normalizeClinician(name) {
  if (!name) return '';
  return name.trim().toLowerCase().replace(/\s+/g, ' ');
}

/************************************
 * HEADER DETECTION / DATA NORMALIZATION
 ************************************/
function normalizeReportData(rawData) {
  // Check if data is from InstaHMS (has 'Pri. Claim No' header)
  const isInsta = rawData[0]?.hasOwnProperty('Pri. Claim No');
  const isOdoo = rawData[0]?.hasOwnProperty('Pri. Claim ID');

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
  const normalizedID = String(memberID || '').trim();
  const eligList = eligMap.get(normalizedID) || [];
  if (!eligList.length) return null;

  for (const elig of eligList) {
    const eligDate = DateHandler.parse(elig["Answered On"]);
    if (!DateHandler.isSameDay(claimDate, eligDate)) continue;
    const eligClinician = (elig.Clinician || '').trim();
    if (eligClinician && claimClinicians.length && !claimClinicians.includes(eligClinician)) continue;

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
        // Find header row
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

        // Deduplicate by Claim ID
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
        fullEligibilityRecord: null
      };
    }

    const hasLeadingZero = memberID.match(/^0+\d+$/);
    const eligibility = findEligibilityForClaim(eligMap, claimDate, memberID, [row.clinician]);
    let status = 'invalid';
    const remarks = [];

    if (hasLeadingZero) remarks.push('Member ID has a leading zero; marked as invalid.');

    if (!eligibility) {
      remarks.push(`No matching eligibility found for ${memberID} on ${formattedDate}`);
    } else if (eligibility.Status?.toLowerCase() !== 'eligible') {
      remarks.push(`Eligibility status: ${eligibility.Status}`);
    } else {
      const serviceCategory = eligibility['Service Category']?.trim() || '';
      const consultationStatus = eligibility['Consultation Status']?.trim()?.toLowerCase() || '';
      const matchesCategory = isServiceCategoryValid(serviceCategory, consultationStatus, row.department || row.clinic).valid;

      if (!matchesCategory) {
        remarks.push(`Invalid for category: ${serviceCategory}, department: ${row.department || row.clinic}`);
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
 * UI RENDERING & STATUS
 ************************************/
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
      .toString().trim().toLowerCase();
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
      detailsCell = `<button class="details-btn eligibility-details" data-index="${index}">${result.fullEligibilityRecord['Eligibility Request Number']}</button>`;
    } else if (eligMap && eligMap.has && eligMap.has(result.memberID)) {
      detailsCell = `<button class="details-btn show-all-eligibilities" data-member="${result.memberID}" data-clinicians="${result.clinician || ''}">View All</button>`;
    }

    row.innerHTML = `
      <td>${result.claimID}</td>
      <td>${result.memberID}</td>
      <td>${result.encounterStart}</td>
      <td class="description-col">${result.packageName}</td>
      <td class="description-col">${result.provider}</td>
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

  initEligibilityModal(results, eligMap);
}

function initEligibilityModal(results) {
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

  document.querySelectorAll(".eligibility-details").forEach(btn => {
    btn.onclick = function() {
      const index = parseInt(this.dataset.index, 10);
      const result = results[index];
      if (!result?.fullEligibilityRecord) return;
      const record = result.fullEligibilityRecord;
      const tableHtml = `
        <h3>Eligibility Details</h3>
        <div style="overflow-x:auto;">
          <table style="width:100%;border-collapse:collapse;">
            <tr><th style="text-align:left;padding:6px;border-bottom:1px solid #ccc;">Eligibility Request Number</th><td style="padding:6px;border-bottom:1px solid #ccc;">${record["Eligibility Request Number"] || ''}</td></tr>
            <tr><th style="text-align:left;padding:6px;border-bottom:1px solid #ccc;">Card Number / DHA Member ID</th><td style="padding:6px;border-bottom:1px solid #ccc;">${record["Card Number / DHA Member ID"] || ''}</td></tr>
            <tr><th style="text-align:left;padding:6px;border-bottom:1px solid #ccc;">Answered On</th><td style="padding:6px;border-bottom:1px solid #ccc;">${record["Answered On"] || ''}</td></tr>
            <tr><th style="text-align:left;padding:6px;border-bottom:1px solid #ccc;">Ordered On</th><td style="padding:6px;border-bottom:1px solid #ccc;">${record["Ordered On"] || ''}</td></tr>
            <tr><th style="text-align:left;padding:6px;border-bottom:1px solid #ccc;">Status</th><td style="padding:6px;border-bottom:1px solid #ccc;">${record["Status"] || ''}</td></tr>
            <tr><th style="text-align:left;padding:6px;border-bottom:1px solid #ccc;">Clinician</th><td style="padding:6px;border-bottom:1px solid #ccc;">${record["Clinician"] || ''}</td></tr>
            <tr><th style="text-align:left;padding:6px;border-bottom:1px solid #ccc;">Payer Name</th><td style="padding:6px;border-bottom:1px solid #ccc;">${record["Payer Name"] || ''}</td></tr>
            <tr><th style="text-align:left;padding:6px;border-bottom:1px solid #ccc;">Service Category</th><td style="padding:6px;border-bottom:1px solid #ccc;">${record["Service Category"] || ''}</td></tr>
            <tr><th style="text-align:left;padding:6px;">Package Name</th><td style="padding:6px;">${record["Package Name"] || ''}</td></tr>
          </table>
        </div>
      `;

      document.getElementById("modalTable").innerHTML = tableHtml;
      document.getElementById("modalOverlay").style.display = "block";
    };
  });
}

function hideModal() {
  const overlay = document.getElementById("modalOverlay");
  if (overlay) overlay.style.display = "none";
}

function updateStatus(message) {
  status.textContent = message || 'Ready';
}

function updateProcessButtonState() {
  const hasEligibility = !!eligData;
  const hasReportData = !!reportData;
  processBtn.disabled = !(hasEligibility && hasReportData);
  exportInvalidBtn.disabled = !(hasEligibility && hasReportData);
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

    // Troubleshooting Print
    console.log("--- Report File Uploaded ---");
    console.log("Raw file:", file.name);
    console.log("Parsed reportData:", reportData);
    if (reportData.length === 0) {
      console.warn("No rows parsed from report file.");
      updateStatus("No rows found in report file.");
    }
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

    // Troubleshooting Print
    console.log("--- Eligibility File Uploaded ---");
    console.log("Raw file:", file.name);
    console.log("Parsed eligData:", eligData);
    if (eligData.length === 0) {
      console.warn("No rows parsed from eligibility file.");
      updateStatus("No rows found in eligibility file.");
    }
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

    // Troubleshooting Print - show eligibility map keys
    console.log("--- Eligibility Map ---");
    console.log("Member IDs:", Array.from(eligMap.keys()));

    let results = validateReportClaims(reportData, eligMap);

    // Troubleshooting Print - raw results before any filter
    console.log("--- Pre-filtered Validation Results ---");
    results.forEach((r, idx) => {
      console.log(`[${idx}]`, r);
    });

    // REMOVE any overly aggressive filtering for troubleshooting!
    // results = results.filter(r => {
    //   const provider = (r.provider || r.insuranceCompany || r.packageName || '').toString().toLowerCase();
    //   return provider.includes('daman') || provider.includes('thiqa');
    // });

    // Troubleshooting Print - after filter (if any)
    // console.log("--- Post-filtered Validation Results ---");
    // results.forEach((r, idx) => {
    //   console.log(`[${idx}]`, r);
    // });

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

/************************************
 * INITIALIZATION
 ************************************/
document.addEventListener('DOMContentLoaded', () => {
  updateProcessButtonState();
  updateStatus('Ready to process files');
  // Troubleshooting Print
  console.log("--- Eligibility Checker Initialized ---");
});
