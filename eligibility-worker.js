/*******************************
 * eligibility-worker.js
 * 
 * Web Worker for processing claims in parallel.
 * Receives batches of claims and validates them against eligibility data.
 *******************************/

// Import constants and utilities (these need to be self-contained in the worker)
const SERVICE_PACKAGE_RULES = {
  'Dental Services': ['dental', 'orthodontic'],
  'Physiotherapy': ['physio'],
  'Other OP Services': ['physio', 'diet', 'occupational', 'speech'],
  'Consultation': []
};
const MONTHS = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"];

/* ===========================
   Utility functions (duplicated from main script for worker context)
   =========================== */
function normalizeMemberID(id) {
  if (!id) return "";
  return String(id).replace(/\D/g, "").trim();
}

function normalizeClinician(name) {
  if (!name) return '';
  return name.trim().toLowerCase().replace(/\s+/g, ' ');
}

/* ===========================
   Date handling
   =========================== */
const DateHandler = {
  parse: function(input, options = {}) {
    const preferMDY = !!options.preferMDY;
    if (!input) return null;
    if (input instanceof Date) return isNaN(input) ? null : input;
    if (typeof input === 'number') return this._parseExcelDate(input);

    const cleanStr = input.toString().trim().replace(/[,.]/g, '');
    const parsed = this._parseStringDate(cleanStr, preferMDY) || new Date(cleanStr);
    if (isNaN(parsed)) {
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
    const utcDays = Math.floor(serial) - 25569;
    const ms = utcDays * 86400 * 1000;
    const date = new Date(ms);
    return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
  },

  _parseStringDate: function(dateStr, preferMDY = false) {
    if (!dateStr) return null;
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
        if (preferMDY) return new Date(Date.UTC(year, part1 - 1, part2));
        return new Date(Date.UTC(year, part2 - 1, part1));
      }
    }

    const textMatch = dateStr.match(/^(\d{1,2})[\/\- ]([a-z]{3,})[\/\- ](\d{2,4})$/i);
    if (textMatch) {
      const monthIndex = MONTHS.indexOf(textMatch[2].toLowerCase().substr(0, 3));
      if (monthIndex >= 0) return new Date(Date.UTC(parseInt(textMatch[3], 10), monthIndex, parseInt(textMatch[1], 10)));
    }

    const isoMatch = dateStr.match(/^(\d{4})[\/\-](\d{2})[\/\-](\d{2})$/);
    if (isoMatch) return new Date(Date.UTC(parseInt(isoMatch[1], 10), parseInt(isoMatch[2], 10) - 1, parseInt(isoMatch[3], 10)));
    return null;
  }
};

/* ===========================
   Validation utilities
   =========================== */
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

function checkClinicianMatch(claimClinicians, eligClinician) {
  if (!eligClinician || !claimClinicians?.length) return true;
  const normElig = normalizeClinician(eligClinician);
  return claimClinicians.some(c => normalizeClinician(c) === normElig);
}

/**
 * Enhanced eligibility finder that returns reasons for rejection
 */
function findEligibilityForClaimWithReasons(eligList, claimDate, claimClinicians = []) {
  if (!eligList.length) return { eligibility: null, reasons: ['No eligibility records found for this member'] };
  
  const rejectionReasons = [];
  
  for (const elig of eligList) {
    const currentReasons = [];
    const eligReqNum = elig['Eligibility Request Number'] || 'Unknown';
    
    // Check date match
    const eligDate = DateHandler.parse(elig["Answered On"]);
    if (!DateHandler.isSameDay(claimDate, eligDate)) {
      currentReasons.push(`Date mismatch: eligibility ${eligReqNum} dated ${DateHandler.format(eligDate)}, claim dated ${DateHandler.format(claimDate)}`);
      rejectionReasons.push(...currentReasons);
      continue;
    }
    
    // Check clinician match (use normalized comparison for consistency)
    const eligClinician = (elig.Clinician || '').trim();
    if (eligClinician && claimClinicians.length && !checkClinicianMatch(claimClinicians, eligClinician)) {
      currentReasons.push(`Clinician mismatch: eligibility has "${eligClinician}", claim has "${claimClinicians.join(', ')}"`);
      rejectionReasons.push(...currentReasons);
      continue;
    }
    
    // Check service category
    const serviceCategory = (elig['Service Category'] || '').trim();
    const consultationStatus = (elig['Consultation Status'] || '').trim();
    const department = (elig.Department || elig.Clinic || '').toLowerCase();
    const categoryCheck = isServiceCategoryValid(serviceCategory, consultationStatus, department);
    if (!categoryCheck.valid) {
      currentReasons.push(categoryCheck.reason);
      rejectionReasons.push(...currentReasons);
      continue;
    }
    
    // Check eligibility status
    if ((elig.Status || '').toLowerCase() !== 'eligible') {
      currentReasons.push(`Status is "${elig.Status}" not "Eligible"`);
      rejectionReasons.push(...currentReasons);
      continue;
    }
    
    // If we get here, this eligibility matches
    return { eligibility: elig, reasons: [] };
  }
  
  // No match found, return all rejection reasons
  return { eligibility: null, reasons: rejectionReasons.length > 0 ? rejectionReasons : ['No matching eligibility found'] };
}

/* ===========================
   Main worker message handler
   =========================== */
self.onmessage = function(e) {
  const { type, data } = e.data;
  
  if (type === 'PROCESS_BATCH') {
    try {
      const { claims, eligibilityMap, preferMDY, workerId } = data;
      const results = [];
      
      // Convert eligibilityMap from plain object back to Map
      const eligMap = new Map(Object.entries(eligibilityMap));
      
      for (let i = 0; i < claims.length; i++) {
        const row = claims[i];
        const claimID = String(row.claimID || '').trim();
        if (!claimID) continue;

        const rawMemberID = String(row.memberID || '').trim();
        if (!rawMemberID) continue;
        
        // Handle VVIP members (check before normalization)
        if (rawMemberID.includes('VVIP') || rawMemberID.startsWith('(VVIP)')) {
          results.push({
            claimID,
            memberID: rawMemberID,
            encounterStart: DateHandler.format(DateHandler.parse(row.claimDate, { preferMDY })),
            status: 'VVIP',
            finalStatus: 'valid',
            remarks: ['VVIP member, eligibility check bypassed'],
            fullEligibilityRecord: null,
            ineligibilityReasons: []
          });
          continue;
        }
        
        const memberID = normalizeMemberID(rawMemberID);

        const insurance = (row.insuranceCompany || '').trim();
        const claimDate = DateHandler.parse(row.claimDate, { preferMDY });
        if (!claimDate) continue;
        const formattedDate = DateHandler.format(claimDate);

        // Get eligibility list for this member
        const eligList = eligMap.get(memberID) || [];
        const { eligibility, reasons } = findEligibilityForClaimWithReasons(
          eligList, 
          claimDate, 
          [row.clinician]
        );
        
        let finalStatus = 'invalid';
        let remarks = [];
        let ineligibilityReasons = [];
        
        if (!eligibility) {
          remarks.push(`No matching eligibility found for ${memberID} on ${formattedDate}`);
          ineligibilityReasons = reasons;
        } else if (eligibility.Status?.toLowerCase() === 'eligible') {
          const categoryCheck = isServiceCategoryValid(
            eligibility['Service Category'],
            eligibility['Consultation Status'],
            (row.department || '').toLowerCase()
          );
          if (categoryCheck.valid) {
            finalStatus = 'valid';
          } else {
            remarks.push(categoryCheck.reason || 'Service category mismatch');
            ineligibilityReasons = [categoryCheck.reason || 'Service category mismatch'];
          }
        } else {
          remarks.push(`Eligibility status: ${eligibility.Status}`);
          ineligibilityReasons = [`Eligibility status: ${eligibility.Status}`];
        }

        results.push({
          claimID,
          memberID,
          encounterStart: formattedDate,
          packageName: eligibility?.['Package Name'] || row.packageName || '',
          provider: insurance,
          clinician: eligibility?.Clinician || row.clinician || '',
          serviceCategory: eligibility?.['Service Category'] || '',
          consultationStatus: eligibility?.['Consultation Status'] || '',
          status: eligibility?.Status || '',
          claimStatus: row.claimStatus || '',
          remarks,
          finalStatus,
          fullEligibilityRecord: eligibility,
          ineligibilityReasons
        });
        
        // Report progress periodically
        if ((i + 1) % 100 === 0 || i === claims.length - 1) {
          self.postMessage({
            type: 'PROGRESS',
            data: {
              workerId,
              processed: i + 1,
              total: claims.length
            }
          });
        }
      }
      
      self.postMessage({
        type: 'BATCH_COMPLETE',
        data: {
          workerId,
          results
        }
      });
    } catch (error) {
      self.postMessage({
        type: 'ERROR',
        data: {
          workerId: data.workerId,
          error: error.message,
          stack: error.stack
        }
      });
    }
  }
};
