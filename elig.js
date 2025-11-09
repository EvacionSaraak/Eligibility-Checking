/*******************************
 * Eligibility Checker - elig.js
 * Adapted to work with elig.html (reportFileInput, eligibilityFileInput, filterDamanThiqa, etc.)
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

// DOM
const reportInput = document.getElementById('reportFileInput');
const eligInput = document.getElementById('eligibilityFileInput');
const processBtn = document.getElementById('processBtn');
const exportInvalidBtn = document.getElementById('exportInvalidBtn');
const statusEl = document.getElementById('uploadStatus');
const resultsContainer = document.getElementById('results');
const filterCheckbox = document.getElementById('filterDamanThiqa');
const filterStatus = document.getElementById('filterStatus');

/***********************
 * Utility / DateFuncs *
 ***********************/
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
    if (!dateStr) return null;
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

    const textMatch = dateStr.match(/^(\d{1,2})[\/\- ]([a-z]{3,})[\/\- ](\d{2,4})$/i);
    if (textMatch) {
      const day = parseInt(textMatch[1], 10);
      let year = parseInt(textMatch[3], 10);
      if (year < 100) year += 2000;
      const mon = MONTHS.indexOf(textMatch[2].toLowerCase().substr(0,3));
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

function normalizeMemberID(id) {
  if (id === null || id === undefined) return '';
  return String(id).trim().replace(/^0+/, '');
}
function normalizeClinician(name) {
  if (!name) return '';
  return String(name).trim().toLowerCase().replace(/\s+/g, ' ');
}
function escapeHtml(s) {
  if (s === null || s === undefined) return '';
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#x27;');
}

/*******************************
 * Eligibility map / matching  *
 *******************************/
function prepareEligibilityMap(eligArray) {
  const map = new Map();
  eligArray.forEach(e => {
    const rawID = e['Card Number / DHA Member ID'] || e['Card Number'] || e['PatientCardID'] || e['Patient Insurance Card No'] || e['Member ID'] || e['MemberID'] || e['_5'];
    if (rawID === undefined || rawID === null || rawID === '') return;
    const key = normalizeMemberID(rawID);
    if (!map.has(key)) map.set(key, []);
    const rec = {
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
    map.get(key).push(rec);
  });
  return map;
}

function checkClinicianMatch(claimClinicians, eligClinician) {
  if (!eligClinician || !claimClinicians?.length) return true;
  const normElig = normalizeClinician(eligClinician);
  return claimClinicians.some(c => normalizeClinician(c) === normElig);
}

function isServiceCategoryValid(serviceCategory, consultationStatus, rawPackage) {
  if (!serviceCategory) return { valid: true };
  const cat = String(serviceCategory).trim().toLowerCase();
  const pkgRaw = rawPackage || '';
  const pkg = String(pkgRaw).toLowerCase();

  if (cat === 'consultation' && consultationStatus?.toLowerCase() === 'elective') {
    const disallowed = ['dental','physio','diet','occupational','speech'];
    if (disallowed.some(t => pkg.includes(t))) {
      return { valid: false, reason: `Consultation (Elective) cannot include restricted service types. Found: "${pkgRaw}"` };
    }
    return { valid: true };
  }

  const allowed = NORMALIZED_SERVICE_PACKAGE_RULES[cat];
  if (allowed && allowed.length > 0) {
    if (pkg && !allowed.some(k => pkg.includes(k))) {
      return { valid: false, reason: `${serviceCategory} requires related package. Found: "${pkgRaw}"` };
    }
  }
  return { valid: true };
}

function findEligibilityForClaim(eligMap, claimDate, memberID, claimClinicians = []) {
  const key = normalizeMemberID(memberID);
  const list = eligMap.get(key) || [];
  if (!list.length) return null;

  const clinics = Array.isArray(claimClinicians) ? claimClinicians.filter(Boolean) : [];
  for (const elig of list) {
    const eligDate = DateHandler.parse(elig['Answered On']);
    if (!DateHandler.isSameDay(claimDate, eligDate)) continue;

    if (elig.Clinician && clinics.length && !checkClinicianMatch(clinics, elig.Clinician)) continue;

    const serviceCategory = elig['Service Category'] || '';
    const consultationStatus = elig['Consultation Status'] || '';
    const dept = (elig.Department || elig.Clinic || '').toLowerCase();
    if (!isServiceCategoryValid(serviceCategory, consultationStatus, dept).valid) continue;

    if ((elig.Status || '').toLowerCase() !== 'eligible') continue;

    if (elig['Eligibility Request Number']) usedEligibilities.add(elig['Eligibility Request Number']);
    return elig;
  }
  return null;
}

/*****************************
 * Parsing: Excel & CSV
 *****************************/
function parseExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: 'array' });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const allRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

        // find header row
        let headerRow = 0;
        let found = false;
        for (let i = 0; i < Math.min(10, allRows.length); i++) {
          const row = (allRows[i] || []).map(c => String(c).trim().toLowerCase());
          const nonEmpty = row.filter(c => c !== '').length;
          if (nonEmpty >= 3) { headerRow = i; found = true; break; }
        }
        if (!found) headerRow = 0;

        const headers = (allRows[headerRow] || []).map(h => String(h).trim());
        const dataRows = allRows.slice(headerRow + 1);
        const json = dataRows.map(r => {
          const obj = {};
          headers.forEach((h, idx) => obj[h] = r[idx] !== undefined && r[idx] !== null ? r[idx] : '');
          return obj;
        });
        resolve(json);
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
        const text = e.target.result;
        const wb = XLSX.read(text, { type: 'string' });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const allRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

        // header detection: first 10 rows
        let headerRow = -1;
        for (let i = 0; i < Math.min(10, allRows.length); i++) {
          const joined = (allRows[i] || []).join(',').toLowerCase();
          if (joined.includes('pri. claim no') || joined.includes('claimid') || joined.includes('member')) {
            headerRow = i; break;
          }
        }
        if (headerRow === -1) {
          for (let i = 0; i < Math.min(10, allRows.length); i++) {
            const row = allRows[i] || [];
            if (row.filter(c => String(c).trim() !== '').length >= 3) { headerRow = i; break; }
          }
        }
        if (headerRow === -1) throw new Error('Could not detect header row in CSV');

        const headers = allRows[headerRow];
        const dataRows = allRows.slice(headerRow + 1);
        const parsed = dataRows.map(r => {
          const obj = {};
          headers.forEach((h, idx) => obj[h] = r[idx] !== undefined && r[idx] !== null ? r[idx] : '');
          return obj;
        });

        // dedupe by claim id if present
        const claimIdHeader = headers.find(h => h && String(h).toLowerCase().replace(/\s+/g,'') === 'claimid') || headers.find(h => h && String(h).toLowerCase().includes('claim'));
        if (!claimIdHeader) return resolve(parsed);
        const seen = new Set();
        const unique = [];
        parsed.forEach(row => {
          const cid = row[claimIdHeader];
          if (!cid) return;
          if (!seen.has(cid)) { seen.add(cid); unique.push(row); }
        });
        resolve(unique);
      } catch (err) { reject(err); }
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsText(file);
  });
}

/*****************************
 * Normalization of reports
 *****************************/
function normalizeReportData(rawData) {
  if (!Array.isArray(rawData)) return [];
  const first = rawData[0] || {};
  const isInsta = Object.prototype.hasOwnProperty.call(first, 'Pri. Claim No');
  const isOdoo = Object.prototype.hasOwnProperty.call(first, 'Pri. Claim ID');

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
        packageName: row['Pri. Plan Type'] || '',
        insuranceCompany: row['Pri. Plan Type'] || '',
        claimStatus: row['Codification Status'] || ''
      };
    } else {
      return {
        claimID: row['ClaimID'] || row['Pri. Claim No'] || '',
        memberID: row['PatientCardID'] || row['Patient Insurance Card No'] || '',
        claimDate: row['ClaimDate'] || row['Encounter Date'] || '',
        clinician: row['Clinician License'] || row['Clinician'] || '',
        packageName: row['Insurance Company'] || '',
        insuranceCompany: row['Insurance Company'] || row['Pri. Payer Name'] || '',
        department: row['Clinic'] || row['Department'] || '',
        claimStatus: row['VisitStatus'] || row['Codification Status'] || ''
      };
    }
  });
}

/*****************************
 * Rendering / Modal UI
 *****************************/
function renderResults(results, eligMap) {
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

    // apply Daman/Thiqa filter if enabled
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
  // create modal if missing
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

  // wire buttons
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
 * Processing / Validation
 *****************************/
function validateReportClaims(reportData, eligMap) {
  const results = reportData.map(row => {
    if (!row || !row.claimID || String(row.claimID).trim() === '') return null;
    const memberID = String(row.memberID || '').trim();
    const claimDateRaw = row.claimDate;
    const claimDate = DateHandler.parse(claimDateRaw, { preferMDY: lastReportWasCSV });
    const encounterStart = DateHandler.format(claimDate);
    const isVVIP = memberID.startsWith('(VVIP)');
    if (isVVIP) {
      return { claimID: row.claimID, memberID, encounterStart, packageName: row.packageName||'', provider: row.provider||'', clinician: row.clinician||'', serviceCategory:'', consultationStatus:'', status:'VVIP', claimStatus: row.claimStatus||'', remarks:['VVIP member, eligibility check bypassed'], finalStatus:'valid', fullEligibilityRecord: null, insuranceCompany: row.insuranceCompany||'' };
    }

    const hasLeadingZero = /^0+\d+$/.test(memberID);
    const claimClinicians = row.clinician ? [row.clinician] : [];
    const eligibility = findEligibilityForClaim(eligMap, claimDate, memberID, claimClinicians);

    const remarks = [];
    let finalStatus = 'invalid';

    if (hasLeadingZero) remarks.push('Member ID has a leading zero; claim marked as invalid.');

    if (!eligibility) {
      remarks.push(`No matching eligibility found for ${memberID} on ${encounterStart}`);
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
      encounterStart,
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
 * Exporting
 *****************************/
function exportInvalidEntries(results) {
  const invalid = (results || []).filter(r => r && r.finalStatus === 'invalid');
  if (!invalid.length) { alert('No invalid entries to export.'); return; }

  const exportData = invalid.map(e => ({
    'Claim ID': e.claimID,
    'Member ID': e.memberID,
    'Encounter Date': e.encounterStart,
    'Package Name': e.packageName || '',
    'Provider': e.provider || '',
    'Clinician': e.clinician || '',
    'Service Category': e.serviceCategory || '',
    'Consultation Status': e.consultationStatus || '',
    'Eligibility Status': e.status || '',
    'Final Status': e.finalStatus,
    'Remarks': (e.remarks || []).join('; ')
  }));
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(exportData);
  XLSX.utils.book_append_sheet(wb, ws, 'Invalid Claims');
  XLSX.writeFile(wb, `invalid_claims_${new Date().toISOString().slice(0,10)}.xlsx`);
}

/*****************************
 * Event handlers
 *****************************/
async function handleFileUpload(e, type) {
  const file = e.target.files && e.target.files[0];
  if (!file) return;
  try {
    updateStatus(`Loading ${type} file...`);
    if (type === 'eligibility') {
      eligData = await parseExcelFile(file);
      updateStatus(`Loaded ${eligData.length} eligibility records`);
      lastReportWasCSV = false;
    } else if (type === 'report') {
      lastReportWasCSV = file.name.toLowerCase().endsWith('.csv');
      const raw = lastReportWasCSV ? await parseCsvFile(file) : await parseExcelFile(file);
      xlsData = normalizeReportData(raw).filter(r => r && r.claimID && String(r.claimID).trim() !== '');
      updateStatus(`Loaded ${xlsData.length} report rows`);
    }
    updateProcessButtonState();
  } catch (err) {
    console.error('File load error:', err);
    updateStatus(`Error loading ${type} file`);
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

/*****************************
 * UI helpers / init
 *****************************/
function updateStatus(msg){ if (statusEl) statusEl.textContent = msg || 'Ready'; }

function onFilterToggle() {
  if (!filterStatus) return;
  const on = filterCheckbox && filterCheckbox.checked;
  filterStatus.textContent = on ? 'ON' : 'OFF';
  filterStatus.classList.toggle('active', on);
  // If results already rendered, re-render using last results
  if (window.lastValidationResults) {
    // Need eligMap to show "View All" buttons; rebuild from eligData
    const eligMap = eligData ? prepareEligibilityMap(eligData) : new Map();
    renderResults(window.lastValidationResults, eligMap);
  }
}

/*********************
 * Initialization
 *********************/
function initializeEventListeners() {
  if (eligInput) eligInput.addEventListener('change', (e) => handleFileUpload(e, 'eligibility'));
  if (reportInput) reportInput.addEventListener('change', (e) => handleFileUpload(e, 'report'));
  if (processBtn) processBtn.addEventListener('click', handleProcessClick);
  if (exportInvalidBtn) exportInvalidBtn.addEventListener('click', () => exportInvalidEntries(window.lastValidationResults || []));
  if (filterCheckbox) filterCheckbox.addEventListener('change', onFilterToggle);
  if (filterStatus) onFilterToggle();
}

document.addEventListener('DOMContentLoaded', () => {
  initializeEventListeners();
  updateStatus('Ready to process files');
});
