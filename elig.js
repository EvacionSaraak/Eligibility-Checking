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
    const d = date.getDate().toString().padStart(2, '0');
    const m = (date.getMonth() + 1).toString().padStart(2, '0');
    const y = date.getFullYear();
    return `${d}/${m}/${y}`;
  },
  isSameDay: function(date1, date2) {
    if (!date1 || !date2) return false;
    return date1.getFullYear() === date2.getFullYear() &&
           date1.getMonth() === date2.getMonth() &&
           date1.getDate() === date2.getDate();
  },
  _parseExcelDate: function(serial) {
    const utcDays = Math.floor(serial) - 25569;
    const ms = utcDays * 86400 * 1000;
    const date = new Date(ms);
    return new Date(date.getFullYear(), date.getMonth(), date.getDate());
  },
  _parseStringDate: function(dateStr, preferMDY = false) {
    if (dateStr.includes(' ')) dateStr = dateStr.split(' ')[0];
    const dmyMdyMatch = dateStr.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
    if (dmyMdyMatch) {
      const part1 = parseInt(dmyMdyMatch[1], 10);
      const part2 = parseInt(dmyMdyMatch[2], 10);
      const year = parseInt(dmyMdyMatch[3], 10);

      if (part1 > 12 && part2 <= 12) return new Date(year, part2 - 1, part1);
      if (part2 > 12 && part1 <= 12) return new Date(year, part1 - 1, part2);
      return preferMDY ? new Date(year, part1 - 1, part2) : new Date(year, part2 - 1, part1);
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
      rawMemberID,
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
      rawMemberID: rawID,
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

  for (const elig of eligList) {
    const eligDate = DateHandler.parse(elig["Answered On"]);
    if (!DateHandler.isSameDay(claimDate, eligDate)) continue;

    if (!checkClinicianMatch(claimClinicians, elig.Clinician)) continue;

    const serviceCategory = (elig['Service Category'] || '').trim();
    const consultationStatus = (elig['Consultation Status'] || '').trim().toLowerCase();
    const department = (elig.Department || elig.Clinic || '').toLowerCase();
    const categoryCheck = isServiceCategoryValid(serviceCategory, consultationStatus, department);
    if (!categoryCheck.valid) continue;

    if ((elig.Status || '').trim().toLowerCase() !== 'eligible') continue;

    return elig;
  }
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
      return { valid: false, reason: `${serviceCategory} requires related package. Found: "${pkgRaw}"` };
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
    const claimDate = DateHandler.parse(row.claimDate, { preferMDY: lastReportWasCSV });
    const formattedDate = DateHandler.format(claimDate);

    const isVVIP = String(memberID || '').toLowerCase().startsWith('(vvip)');
    if (isVVIP) return {
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

    const hasLeadingZero = String(row.rawMemberID || '').match(/^0+\d+$/);
    const eligibility = findEligibilityForClaim(eligMap, claimDate, memberID, [row.clinician].filter(Boolean));
    let status = 'invalid';
    const remarks = [];

    if (!eligibility) remarks.push(`No matching eligibility found for ${memberID} on ${formattedDate}`);
    else if (!hasLeadingZero) status = 'valid';
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
          if (row.filter(c => c !== '').length >= 3) foundHeaders = true;
          else headerRow++;
        }
        if (!foundHeaders) headerRow = 0;
        const headers = allRows[headerRow].map(h => String(h).trim());
        const dataRows = allRows.slice(headerRow + 1);
        resolve(dataRows.map(row => {
          const obj = {};
          headers.forEach((h,i) => obj[h] = row[i] || '');
          return obj;
        }));
      } catch (err) { reject(err); }
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
        const workbook = XLSX.read(e.target.result, { type: 'string' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const allRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
        let headerRowIndex = -1;
        for (let i=0;i<5;i++){
          if (allRows[i]?.join(',').toLowerCase().includes('claim') && allRows[i].join(',').toLowerCase().includes('id')){
            headerRowIndex = i; break;
          }
        }
        if(headerRowIndex === -1) throw new Error("Could not detect header row in CSV");
        const headers = allRows[headerRowIndex];
        const rawParsed = allRows.slice(headerRowIndex+1).map(row => {
          const obj = {};
          headers.forEach((h,i)=>obj[h]=row[i]||'');
          return obj;
        });
        const seen = new Set();
        const uniqueRows = [];
        const claimIdHeader = headers.find(h=>h.toLowerCase().includes('claim'));
        rawParsed.forEach(r=>{
          const id=r[claimIdHeader];
          if(id&&!seen.has(id)){seen.add(id);uniqueRows.push(r);}
        });
        resolve(uniqueRows);
      } catch(err){reject(err);}
    };
    reader.onerror = ()=>reject(reader.error);
    reader.readAsText(file);
  });
}

function exportInvalidEntries(results) {
  const invalidEntries = results.filter(r=>r&&r.finalStatus==='invalid');
  if(invalidEntries.length===0){alert('No invalid entries to export.'); return;}
  const exportData = invalidEntries.map(e=>({
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
    'Remarks': e.remarks.join('; ')
  }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(exportData),'Invalid Claims');
  XLSX.writeFile(wb,`invalid_claims_${new Date().toISOString().slice(0,10)}.xlsx`);
}

/************************************
 * RENDERING RESULTS
 ************************************/
function renderResults(results, eligMap) {
  const resultsContainer=document.getElementById("results");
  if(!resultsContainer){warn('#results container not found'); return;}
  resultsContainer.innerHTML='';
  if(!results||results.length===0){resultsContainer.innerHTML='<div class="no-results">No claims to display</div>';return;}

  const tableContainer=document.createElement('div');
  tableContainer.className='analysis-results';
  tableContainer.style.overflowX='auto';

  const table=document.createElement('table');
  table.className='shared-table';

  const thead=document.createElement('thead');
  const headerRow=document.createElement('tr');
  ['Claim ID','Member ID','Encounter','Package','Provider','Clinician','Service Category','Consultation Status','Eligibility Status','Final Status','Remarks'].forEach(h=>{
    const th=document.createElement('th');th.textContent=h;headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);
  table.appendChild(thead);

  const tbody=document.createElement('tbody');
  results.forEach(result=>{
    const tr=document.createElement('tr');
    ['claimID','memberID','encounterStart','packageName','provider','clinician','serviceCategory','consultationStatus','status','finalStatus'].forEach(key=>{
      const td=document.createElement('td');td.textContent=result[key]||'';tr.appendChild(td);
    });
    const tdRemark=document.createElement('td');tdRemark.textContent=result.remarks.join('; ');tr.appendChild(tdRemark);
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  tableContainer.appendChild(table);
  resultsContainer.appendChild(tableContainer);
  updateStatus(`Processed ${results.length} claims`);
}
