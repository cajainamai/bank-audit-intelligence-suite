import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { calculateTenure, getInstalmentsDue, calculateEMI, getAmortizationBreakdown, calculateIrregularity, getSMATag } from './loanCalculations';

// Helper to reliably parse numeric fields from strings like "1,234,567.89" or "Rs. 100"
const parseAmount = (val) => {
    if (val === null || val === undefined) return 0;
    if (typeof val === 'number') return val;
    const str = String(val).replace(/,/g, '').replace(/[^\d.-]/g, '');
    const parsed = parseFloat(str);
    return isNaN(parsed) ? 0 : parsed;
};

// Helper to convert 1-based index or 0-based index to Excel column letter
// For 0-based index: getColumnLetter(colIndex + 1)
const getColumnLetter = (colIdx) => {
    if (!colIdx || colIdx <= 0) return "A"; // Fallback to avoid empty strings
    let letter = '';
    let temp = colIdx;
    while (temp > 0) {
        let remainder = (temp - 1) % 26;
        letter = String.fromCharCode(65 + remainder) + letter;
        temp = Math.floor((temp - remainder) / 26);
    }
    return letter;
};

// Helper for dates. Excel dates can be string "DD/MM/YYYY" or numeric serial logs
const parseDateString = (val) => {
    if (!val) return null;
    // If it's a numeric excel date:
    if (typeof val === 'number') {
        const date = new Date(Math.round((val - 25569) * 86400 * 1000));
        return date;
    }

    if (val instanceof Date) return val;

    const str = String(val);
    // Attempt to parse standard formats (MM/DD/YYYY, YYYY-MM-DD)
    const dateObj = new Date(str);
    if (!isNaN(dateObj.getTime())) return dateObj;

    // Attempt DD/MM/YYYY strictly
    if (str.includes('/')) {
        const parts = str.split('/');
        if (parts.length === 3) {
            // Assuming DD/MM/YYYY
            return new Date(`${parts[2]}-${parts[1]}-${parts[0]}`);
        }
    } else if (str.includes('-')) {
        const parts = str.split('-');
        if (parts.length === 3) {
            // Assuming DD-MM-YY
            return new Date(`${parts[2]}-${parts[1]}-${parts[0]}`);
        }
    }
    return null;
};

const extractEntityType = (pan) => {
    // If PAN is missing or too short, return NA
    const panStr = pan ? String(pan).trim() : '';
    if (!panStr || panStr.length < 4) return 'NA';
    const char = panStr.toUpperCase().charAt(3);

    switch (char) {
        case 'P': return 'Individual';
        case 'C': return 'Company';
        case 'H': return 'HUF';
        case 'F': return 'Partnership Firm / LLP';
        case 'L': return 'LLP';
        case 'A': return 'AOP (Association of Persons)';
        case 'T': return 'Trust';
        case 'B': return 'BOI (Body of Individuals)';
        case 'G': return 'Government Agency';
        case 'J': return 'Artificial Juridical Person';
        default: return 'Unknown';
    }
};

export const processAuditData = (rawData, mappingRules, auditPeriod, targetFields, externalNpaData = null) => {

    const fromDateObj = new Date(auditPeriod.fromDate);
    const toDateObj = new Date(auditPeriod.toDate);

    // Convert standard data based on map
    const processedData = rawData.map(row => {
        const newRow = {};

        targetFields.forEach(field => {

            if (field === "Loan Type") {
                // This is mapped in Phase 2
                newRow[field] = row['Mapped Loan Type'] || '';
            }
            else if (field === "Overdrawn (Yes/No)") {
                const sourceNetOs = mappingRules["Net O/s"] ? row[mappingRules["Net O/s"]] : 0;
                const sourceSanction = mappingRules["Sanction Limit"] ? row[mappingRules["Sanction Limit"]] : 0;

                const netOs = parseAmount(sourceNetOs);
                const sanction = parseAmount(sourceSanction);

                newRow[field] = netOs > sanction ? 'Yes' : 'No';
            }
            else if (field === "Sanction Type") {
                const sourceSanctionDate = mappingRules["Sanction Date"] ? row[mappingRules["Sanction Date"]] : null;
                const sanctionDateObj = parseDateString(sourceSanctionDate);

                if (sanctionDateObj) {
                    if (sanctionDateObj >= fromDateObj && sanctionDateObj <= toDateObj) {
                        newRow[field] = 'New';
                    } else {
                        newRow[field] = 'Old';
                    }
                } else {
                    newRow[field] = '';
                }
            }
            else if (field === "PAN") {
                // Direct mapping but mark blank PAN as "NA"
                const sourceCol = mappingRules["PAN"];
                const rawPan = sourceCol ? String(row[sourceCol] || '').trim() : '';
                newRow[field] = rawPan || 'NA';
            }
            else if (field === "Entity Type") {
                const sourcePan = mappingRules["PAN"] ? String(row[mappingRules["PAN"]] || '').trim() : '';
                newRow[field] = extractEntityType(sourcePan);
            }
            else if (field === "Asset Classification") {
                const sourceAssetClass = mappingRules["Asset Class"] ? String(row[mappingRules["Asset Class"]] || '').trim().toUpperCase() : '';
                // Blank / unmapped asset class → leave empty; don't default to NPA
                if (!sourceAssetClass) {
                    newRow[field] = '';
                } else if (['STD', 'SMA0', 'SMA1', 'SMA2'].includes(sourceAssetClass)) {
                    newRow[field] = 'Standard';
                } else {
                    newRow[field] = 'NPA';
                }
            }
            else if (field === "NPA Provision") {
                if (externalNpaData) {
                    const acNoCol = mappingRules["A/c No."];
                    const acNo = acNoCol ? String(row[acNoCol] || '').trim() : '';
                    newRow[field] = parseAmount(externalNpaData[acNo]);
                } else {
                    const sourceCol = mappingRules["NPA Provision"];
                    newRow[field] = sourceCol ? parseAmount(row[sourceCol]) : 0;
                }
            }
            else if (field === "Outstanding EMIs") {
                const loanType = String(row['Mapped Loan Type'] || '').toUpperCase();
                const facilityCol = mappingRules["Facility Code"];
                const code = facilityCol ? String(row[facilityCol] || '').toUpperCase() : '';
                
                // Identify if it's a Term Loan / Instalment loan
                const isTL = code === 'TL' || code.includes('TERM LOAN') || code.includes('HL') || code.includes('VL') || code.includes('PL') || code.includes('EDU')
                    || loanType.includes('TERM LOAN') || loanType === 'TL'
                    || loanType.includes('HOUSING') || loanType.includes('VEHICLE') || loanType.includes('PERSONAL') || loanType.includes('MSME');

                if (isTL) {
                    const principal = parseAmount(mappingRules["Sanction Limit"] ? row[mappingRules["Sanction Limit"]] : 0);
                    const rawRoi = mappingRules["ROI"] ? row[mappingRules["ROI"]] : 0;
                    // Calculation utility expects percentage as 8.5 for 8.5%
                    const roiPercent = typeof rawRoi === 'number' ? rawRoi * 100 : parseAmount(rawRoi);
                    const sDate = parseDateString(mappingRules["Sanction Date"] ? row[mappingRules["Sanction Date"]] : null);
                    const eDate = parseDateString(mappingRules["Limit Expiry Date"] ? row[mappingRules["Limit Expiry Date"]] : null);
                    const netOs = parseAmount(mappingRules["Net O/s"] ? row[mappingRules["Net O/s"]] : 0);
                    const creditBal = parseAmount(mappingRules["Credit Balance"] ? row[mappingRules["Credit Balance"]] : 0);
                    const nettedOs = netOs - creditBal;

                    if (principal > 0 && sDate && eDate) {
                        const ten = calculateTenure(sDate, eDate);
                        const theoEmi = calculateEMI(principal, roiPercent, ten);
                        const instDue = getInstalmentsDue(sDate, toDateObj);
                        const bk = getAmortizationBreakdown(principal, roiPercent, ten, instDue);
                        const osDiff = Math.max(0, nettedOs - bk.theoBalance);
                        newRow[field] = (theoEmi > 0 && osDiff > 0) ? Math.floor(osDiff / theoEmi) : 0;
                    } else {
                        newRow[field] = 0;
                    }
                } else {
                    newRow[field] = 0;
                }
            }
            else {
                // Direct straight mapping or blank
                const sourceCol = mappingRules[field];
                let val = sourceCol ? row[sourceCol] : '';

                // Formatting Cleanup
                if (val !== null && val !== undefined && val !== '') {
                    if (['Sanction Limit', 'Net O/s', 'Credit Balance', 'Drawing Power', 'Primary Security', 'Collateral Security', 'Overdue Amount'].includes(field)) {
                        val = parseAmount(val);
                    } else if (['Sanction Date', 'Limit Expiry Date', 'A/c Open Date', 'NPA Date'].includes(field)) {
                        val = parseDateString(val) || val;
                    } else if (field === 'Aadhar No.') {
                        val = String(val);
                    } else if (field === 'ROI') {
                        const rawStr = String(val);
                        let num = parseAmount(val);
                        if (rawStr.includes('%')) {
                            val = num / 100;
                        } else if (num > 1) {
                            val = num / 100;
                        } else {
                            val = num;
                        }
                    }
                }
                newRow[field] = val;
            }

        });

        return newRow;
    });

    return processedData;
};
// ─────────────────────────────────────────────────────────────────────────────
// RISK METRICS ENGINE
// Computes a weighted Risk Score (0-100), a Risk Label, and a one-line Audit
// Remark for each account row on the processed dataset.
//
// Weight Design Philosophy: NPA accounts are already classified — the audit
// value is in identifying STANDARD accounts with hidden problems. All 7
// dimensions focus on observable financial health signals applicable to any
// account (Standard or NPA), so a stressed SMA-2 overdrawn Standard account
// can score higher than a well-secured, newly-classified NPA.
//
// Dimension Weights:
//   1. Overdrawn / Limit Breach      25  —  direct credit discipline failure
//   2. Limit Expired (still active)  20  —  regulatory / renewal risk
//   3. Overdue Amount                20  —  payment irregularity signal
//   4. SMA / NPA-age stress          15  —  pre-NPA stress or NPA severity
//   5. Security / Collateral         10  —  coverage adequacy
//   6. Drawing Power Breach           5  —  CC/OD norm violation
//   7. PAN / KYC Missing              5  —  compliance gap
// ─────────────────────────────────────────────────────────────────────────────
const calculateRiskMetrics = (row, auditToObj) => {
    const netOs          = parseAmount(row['Net O/s']);
    const sanction       = parseAmount(row['Sanction Limit']);
    const primarySec     = parseAmount(row['Primary Security']);
    const collateralSec  = parseAmount(row['Collateral Security']);
    const overdueAmt     = parseAmount(row['Overdue Amount']);
    const drawingPower   = parseAmount(row['Drawing Power']);
    const assetClass     = String(row['Asset Classification'] || '').trim();
    const assetClassRaw  = String(row['Asset Class'] || '').trim().toUpperCase(); // Raw CBS tag e.g. SMA0, SMA1, SMA2
    const overdrawn      = String(row['Overdrawn (Yes/No)'] || '').toUpperCase();
    const pan            = String(row['PAN'] || '').trim().toUpperCase();
    const expiryDate     = parseDateString(row['Limit Expiry Date']);
    const npaDate        = parseDateString(row['NPA Date']);

    const today = auditToObj instanceof Date && !isNaN(auditToObj) ? auditToObj : new Date();
    const isNpa = assetClass === 'NPA';
    const utilisationPct = sanction > 0 ? netOs / sanction : 0;

    // ── NPA Accounts: Risk scoring matrix (NPA Age + Security + Overdue) ────────────────
    if (isNpa) {
        const npaAgeDays = npaDate ? Math.floor((today - npaDate) / (1000 * 60 * 60 * 24)) : 0;
        const totalSecurity = primarySec + collateralSec;
        const secRatio = netOs > 0 ? totalSecurity / netOs : 1.0;
        const overdueRatio = netOs > 0 ? overdueAmt / netOs : 0;

        // 1. Asset Aging (Weight: 50)
        let ageScore = 10; // Sub-standard (< 1 yr)
        if      (npaAgeDays >= 1095) ageScore = 50; // Doubtful D3 / Loss (> 3 yrs)
        else if (npaAgeDays >= 730)  ageScore = 40; // Doubtful D2 (2-3 yrs)
        else if (npaAgeDays >= 365)  ageScore = 25; // Doubtful D1 (1-2 yrs)

        // 2. Security Coverage (Weight: 40)
        let securityScore = 0;
        if      (secRatio < 0.25) securityScore = 40; // Practically Unsecured
        else if (secRatio < 0.50) securityScore = 30; // Highly Under-secured
        else if (secRatio < 0.75) securityScore = 20; // Partially Secured
        else if (secRatio < 1.00) securityScore = 10; // Marginal Security gap

        // 3. Overdue Intensity (Weight: 10)
        let overdueScore = 0;
        if      (overdueRatio > 0.50) overdueScore = 10;
        else if (overdueRatio > 0.10) overdueScore = 5;

        const score = ageScore + securityScore + overdueScore;

        let riskLabel;
        if      (score >= 81) riskLabel = 'NPA - Critical';
        else if (score >= 61) riskLabel = 'NPA - High Risk';
        else if (score >= 31) riskLabel = 'NPA - Medium Risk';
        else                  riskLabel = 'NPA - Low Risk';

        const npaAgeDesc = npaAgeDays > 365 ? `${(npaAgeDays / 365).toFixed(1)} years` : `${npaAgeDays} days`;
        const remarkParts = [`NPA account (${npaAgeDesc} since NPA date)`];
        if (overdueAmt > 0) remarkParts.push(`overdue Rs.${Math.round(overdueAmt).toLocaleString('en-IN')}`);
        if (totalSecurity < netOs) {
            const secCovPct = netOs > 0 ? ((totalSecurity / netOs) * 100).toFixed(0) : '0';
            remarkParts.push(`security coverage ${secCovPct}% of outstanding`);
        } else {
            remarkParts.push('fully secured');
        }
        
        const remark = remarkParts.length > 0
            ? remarkParts.map((r, i) => i === 0 ? r.charAt(0).toUpperCase() + r.slice(1) : r).join('; ') + '.'
            : 'NPA account details.';

        return { riskScore: score, riskLabel, auditRemark: remark };
    }


    let overdrawnRaw = 0;
    if (overdrawn === 'YES' || utilisationPct >= 1.0) {
        if      (utilisationPct >= 1.30) overdrawnRaw = 5; // >130% — severe breach
        else if (utilisationPct >= 1.15) overdrawnRaw = 4; // >115%
        else if (utilisationPct >= 1.00) overdrawnRaw = 3; // >100%
    } else if (utilisationPct >= 0.90) {
        overdrawnRaw = 2; // 90–100% — borderline
    } else if (utilisationPct >= 0.75) {
        overdrawnRaw = 1; // 75–90% — moderately high
    }

    // ─── 2. Limit Expiry Risk  (weight: 20) ──────────────────────────────────
    let expiryRaw = 0;
    let daysToExpiry = null;
    if (expiryDate) {
        daysToExpiry = Math.floor((expiryDate - today) / (1000 * 60 * 60 * 24));
        if      (daysToExpiry < 0)    expiryRaw = 5; // Already expired
        else if (daysToExpiry <= 30)  expiryRaw = 4; // Expires within 30 days
        else if (daysToExpiry <= 90)  expiryRaw = 3; // Expires within 90 days
        else if (daysToExpiry <= 180) expiryRaw = 1; // Expiring within 6 months
    }

    // ─── 3. Overdue Amount  (weight: 20) ─────────────────────────────────────
    let overdueRaw = 0;
    if (overdueAmt > 0) {
        const overduePct = sanction > 0 ? overdueAmt / sanction : 0;
        if      (overduePct >= 0.50) overdueRaw = 5;
        else if (overduePct >= 0.25) overdueRaw = 4;
        else if (overduePct >= 0.10) overdueRaw = 3;
        else                         overdueRaw = 2;
    }

    // ─── 4. SMA / NPA-age Stress  (weight: 15) ───────────────────────────────
    let smaRaw = 0;
    if (isNpa) {
        if (npaDate) {
            const npaAgeDays = Math.floor((today - npaDate) / (1000 * 60 * 60 * 24));
            if      (npaAgeDays >= 730) smaRaw = 5;
            else if (npaAgeDays >= 365) smaRaw = 4;
            else if (npaAgeDays >= 90)  smaRaw = 3;
            else                        smaRaw = 2;
        } else {
            smaRaw = 3;
        }
    } else {
        if      (assetClassRaw === 'SMA2') smaRaw = 5;
        else if (assetClassRaw === 'SMA1') smaRaw = 4;
        else if (assetClassRaw === 'SMA0') smaRaw = 3;
    }

    // ─── 5. Security / Collateral Shortfall  (weight: 10) ────────────────────
    const totalSecurity = primarySec + collateralSec;
    let securityRaw = 0;
    if (sanction > 0) {
        const secRatio = totalSecurity / sanction;
        if      (secRatio < 0.50) securityRaw = 5;
        else if (secRatio < 0.75) securityRaw = 4;
        else if (secRatio < 1.00) securityRaw = 3;
        else if (secRatio < 1.25) securityRaw = 1;
    }

    // ─── 6. Drawing Power Breach  (weight: 5) ────────────────────────────────
    let dpBreachRaw = 0;
    if (drawingPower > 0 && netOs > drawingPower) {
        const dpExcess = netOs / drawingPower;
        if      (dpExcess >= 1.25) dpBreachRaw = 5;
        else if (dpExcess >= 1.10) dpBreachRaw = 4;
        else                       dpBreachRaw = 3;
    }

    // ─── 7. PAN / KYC Missing  (weight: 5) ───────────────────────────────────
    const isPanMissing = !pan || pan === 'NA' || pan.length < 10;
    const panRaw = isPanMissing ? 5 : 0;

    // ─── Weighted Score (0–100) ───────────────────────────────────────────────
    const score = Math.round(
        (overdrawnRaw / 5) * 25 +
        (expiryRaw    / 5) * 20 +
        (overdueRaw   / 5) * 20 +
        (smaRaw       / 5) * 15 +
        (securityRaw  / 5) * 10 +
        (dpBreachRaw  / 5) *  5 +
        (panRaw       / 5) *  5
    );

    // ── Risk Label ──
    let riskLabel;
    if      (score >= 71) riskLabel = 'Critical';
    else if (score >= 46) riskLabel = 'High';
    else if (score >= 21) riskLabel = 'Medium';
    else                  riskLabel = 'Low';

    // ── Audit Remark (Standard/SMA accounts only — NPA already returned above) ──
    const remarks = [];

    if (assetClassRaw === 'SMA2') {
        remarks.push('SMA-2 (61-90 days past due) — near-NPA stress');
    } else if (assetClassRaw === 'SMA1') {
        remarks.push('SMA-1 (31-60 days past due) — payment stress');
    } else if (assetClassRaw === 'SMA0') {
        remarks.push('SMA-0 (1-30 days past due) — early delinquency');
    }

    if (overdrawnRaw >= 3) {
        const excessPct = ((utilisationPct - 1) * 100).toFixed(1);
        remarks.push(utilisationPct >= 1.0
            ? `overdrawn by ${excessPct}% above sanction limit`
            : `utilisation at ${(utilisationPct * 100).toFixed(1)}%`);
    } else if (overdrawnRaw >= 1) {
        remarks.push(`utilisation at ${(utilisationPct * 100).toFixed(1)}%`);
    }

    if (expiryRaw > 0) {
        if (daysToExpiry < 0) remarks.push(`limit expired ${Math.abs(daysToExpiry)} days ago`);
        else                  remarks.push(`limit expiring in ${daysToExpiry} days`);
    }

    if (overdueAmt > 0) {
        remarks.push(`overdue Rs.${Math.round(overdueAmt).toLocaleString('en-IN')}`);
    }

    if (securityRaw >= 3) {
        const secCovPct = sanction > 0 ? ((totalSecurity / sanction) * 100).toFixed(0) : '0';
        remarks.push(`security coverage ${secCovPct}% of limit`);
    }

    if (dpBreachRaw > 0) {
        remarks.push(`outstanding exceeds drawing power by Rs.${Math.round(netOs - drawingPower).toLocaleString('en-IN')}`);
    }

    if (isPanMissing) remarks.push('PAN not available (KYC gap)');

    const remark = remarks.length > 0
        ? remarks.map((r, i) => i === 0 ? r.charAt(0).toUpperCase() + r.slice(1) : r).join('; ') + '.'
        : 'No adverse flags detected as on audit date.';

    return { riskScore: score, riskLabel, auditRemark: remark };
};

export const exportToExcel = async (processedData, targetFields, auditPeriod, npaConfig = null) => {
    const workbook = new ExcelJS.Workbook();

    // Parse audit period dates — these are used throughout the function
    const auditFromObj = new Date(auditPeriod.fromDate);
    const auditToObj = new Date(auditPeriod.toDate);

    // --- Generate Parameters Sheet ---
    const paramSheet = workbook.addWorksheet('Parameters');
    paramSheet.columns = [
        { header: 'Parameter', key: 'param', width: 35 },
        { header: 'Value', key: 'val', width: 25 }
    ];
    paramSheet.addRows([
        { param: 'Audit From Date', val: auditFromObj },
        { param: 'Audit To Date', val: auditToObj }
    ]);
    paramSheet.getRow(1).font = { bold: true };
    paramSheet.getCell('B2').numFmt = 'DD-MM-YY';
    paramSheet.getCell('B3').numFmt = 'DD-MM-YY';
    // ---------------------------------

    const worksheet = workbook.addWorksheet('Audit Report');

    // Sort master data by netted Net O/s descending, then augment with risk metrics
    const auditReportData = processedData.map(row => ({
        ...row,
        'Net O/s': parseAmount(row['Net O/s']) - parseAmount(row['Credit Balance'])
    })).sort((a, b) => b['Net O/s'] - a['Net O/s']);

    // Augment each row with risk metrics
    auditReportData.forEach(row => {
        const { riskScore, riskLabel, auditRemark } = calculateRiskMetrics(row, auditToObj);
        row['Risk Score'] = riskScore;
        row['Risk Label'] = riskLabel;
        row['Audit Remark'] = auditRemark;
    });

    // Define Columns (exclude Credit Balance since it's netted into Net O/s)
    // Exclude Credit Balance from display since it's netted into Net O/s
    // Exclude NPA Provision if not available
    const baseAuditReportFields = targetFields.filter(f =>
        f !== 'Credit Balance' &&
        (f !== 'NPA Provision' || npaConfig?.available)
    );
    // Append risk columns at the end
    const auditReportFields = [...baseAuditReportFields, 'Risk Score', 'Risk Label', 'Audit Remark'];

    worksheet.columns = auditReportFields.map(field => {
        let width = 20; // Default
        if (field === 'Borrower Name') width = 35;
        if (field === 'CIF' || field === 'A/c No.' || field === 'PAN' || field === 'Aadhar No.') width = 18;
        if (field.includes('Date')) width = 15;
        if (field === 'Audit Remark') width = 55;
        if (field === 'Risk Label') width = 14;
        if (field === 'Risk Score') width = 13;
        return { header: field, key: field, width };
    });

    // Add Data
    worksheet.addRows(auditReportData);

    // Apply specific formulas to Master Data
    const netOsCol = getColumnLetter(auditReportFields.indexOf('Net O/s') + 1);
    const sanctionCol = getColumnLetter(auditReportFields.indexOf('Sanction Limit') + 1);
    const overdrawnCol = auditReportFields.indexOf('Overdrawn (Yes/No)') + 1;

    const assetClassCol = getColumnLetter(auditReportFields.indexOf('Asset Class') + 1);
    const assetClassificationCol = auditReportFields.indexOf('Asset Classification') + 1;

    const sanctionTypeCol = auditReportFields.indexOf('Sanction Type') + 1;
    const sanctionTypeDateCol = auditReportFields.indexOf('Sanction Date') + 1;

    // Styling the Header Row
    const headerRow = worksheet.getRow(1);
    headerRow.height = 30;
    headerRow.eachCell((cell) => {
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFE6E6FA' } // Pastel Lavender
        };
        cell.font = {
            bold: true,
            color: { argb: 'FF483D8B' }, // Dark Slate Blue
            size: 12
        };
        cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        cell.border = {
            top: { style: 'thin', color: { argb: 'FFCCCCCC' } },
            left: { style: 'thin', color: { argb: 'FFCCCCCC' } },
            bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } },
            right: { style: 'thin', color: { argb: 'FFCCCCCC' } }
        };
    });

    // Styling the Data Rows and Applying Formulas
    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return; // Skip header

        // Apply Master Sheet Formulas
        if (netOsCol && sanctionCol && overdrawnCol > 0) {
            const netOsRef = `${netOsCol}${rowNumber}`;
            const sanctionRef = `${sanctionCol}${rowNumber}`;
            const fallbackResult = auditReportData[rowNumber - 2]['Overdrawn (Yes/No)'];
            row.getCell(overdrawnCol).value = { formula: `IF(${netOsRef}>${sanctionRef},"Yes","No")`, result: fallbackResult };
        }

        if (assetClassCol && assetClassificationCol > 0) {
            const acRef = `${assetClassCol}${rowNumber}`;
            const fallbackResult = auditReportData[rowNumber - 2]['Asset Classification'];
            row.getCell(assetClassificationCol).value = {
                formula: `IF(OR(${acRef}="STD", ${acRef}="SMA0", ${acRef}="SMA1", ${acRef}="SMA2"), "Standard", "NPA")`,
                result: fallbackResult
            };
        }

        if (sanctionTypeCol > 0 && sanctionTypeDateCol > 0) {
            const stDateRef = `${getColumnLetter(sanctionTypeDateCol)}${rowNumber}`;
            const fallbackResult = auditReportData[rowNumber - 2]['Sanction Type'];
            row.getCell(sanctionTypeCol).value = {
                formula: `IF(${stDateRef}="","",IF(AND(${stDateRef}>='Parameters'!$B$2, ${stDateRef}<='Parameters'!$B$3),"New","Old"))`,
                result: fallbackResult
            };
        }

        // Determine risk label for color-coding (pre-computed in augment step)
        const rowData = auditReportData[rowNumber - 2];
        const rowRiskLabel = rowData ? rowData['Risk Label'] : 'Low';

        row.eachCell((cell, colNumber) => {
            const fieldName = auditReportFields[colNumber - 1];

            cell.alignment = { vertical: 'middle', horizontal: 'left' };

            cell.border = {
                top: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                left: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                bottom: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                right: { style: 'thin', color: { argb: 'FFEEEEEE' } }
            };

            // Number Formatting Rules
            if (['Sanction Limit', 'Net O/s', 'Credit Balance', 'Drawing Power', 'Primary Security', 'Collateral Security', 'Overdue Amount'].includes(fieldName)) {
                cell.numFmt = '#,##0';
                cell.alignment = { vertical: 'middle', horizontal: 'right' };
            } else if (fieldName === 'ROI') {
                cell.numFmt = '0.00%';
                cell.alignment = { vertical: 'middle', horizontal: 'right' };
            } else if (['Sanction Date', 'Limit Expiry Date', 'A/c Open Date', 'NPA Date'].includes(fieldName)) {
                cell.numFmt = 'DD-MM-YY';
                cell.alignment = { vertical: 'middle', horizontal: 'center' };
            } else if (['CIF', 'A/c No.', 'PAN', 'Aadhar No.'].includes(fieldName)) {
                cell.numFmt = '@'; // Force text format
            } else if (['Overdrawn (Yes/No)', 'Sanction Type', 'PS Flag', 'Loan Type', 'Asset Classification'].includes(fieldName)) {
                cell.alignment = { vertical: 'middle', horizontal: 'center' };
            } else if (fieldName === 'NPA Provision') {
                cell.numFmt = '#,##0';
                cell.alignment = { vertical: 'middle', horizontal: 'right' };
            } else if (fieldName === 'Risk Score') {
                cell.alignment = { vertical: 'middle', horizontal: 'center' };
                const rs = { bg: 'FFFFFFFF', fg: 'FF000000' };
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rs.bg } };
                cell.font = { bold: true, size: 10, color: { argb: rs.fg } };
                cell.numFmt = '0';
            } else if (fieldName === 'Risk Label') {
                cell.alignment = { vertical: 'middle', horizontal: 'center' };
                const rs = { bg: 'FFFFFFFF', fg: 'FF000000' };
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rs.bg } };
                cell.font = { bold: true, color: { argb: rs.fg } };
            } else if (fieldName === 'Audit Remark') {
                cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
                cell.font = { italic: true, size: 10, color: { argb: 'FF374151' } };
            }
        });
    });

    // ─────────────────────────────────────────────────────────────────────────────
    // ADDITIONAL AUDIT SHEETS
    // ─────────────────────────────────────────────────────────────────────────────

    // --- Generate Audit Findings Summary Sheet ---
    // ─────────────────────────────────────────────────────────────────────────────
    {
        const summarySheet = workbook.addWorksheet('Audit Findings Summary');

        // ── Color palette (consistent with indigo theme) ──
        const COLOR = {
            INDIGO_700:  'FF000000',
            INDIGO_100:  'FFFFFFFF',
            WHITE:       'FFFFFFFF',
            GRAY_50:     'FFFFFFFF',
            GRAY_200:    'FFCCCCCC',
            RED_BG:      'FFFFFFFF',
            RED_FG:      'FF000000',
            ORANGE_BG:   'FFFFFFFF',
            ORANGE_FG:   'FF000000',
            AMBER_BG:    'FFFFFFFF',
            AMBER_FG:    'FF000000',
            GREEN_BG:    'FFFFFFFF',
            GREEN_FG:    'FF000000',
            DARK:        'FF000000'
        };

        const riskColors = {
            'Low':      { bg: 'FFFFFFFF',  fg: 'FF000000' },
            'Medium':   { bg: 'FFFFFFFF',  fg: 'FF000000' },
            'High':     { bg: 'FFFFFFFF', fg: 'FF000000' },
            'Critical': { bg: 'FFFFFFFF',    fg: 'FF000000' },
            'NPA - Low Risk':    { bg: 'FFFFFFFF', fg: 'FF000000' },
            'NPA - Medium Risk': { bg: 'FFFFFFFF', fg: 'FF000000' },
            'NPA - High Risk':   { bg: 'FFFFFFFF', fg: 'FF000000' },
            'NPA - Critical':    { bg: 'FFFFFFFF', fg: 'FF000000' },
            'NPA':               { bg: 'FFFFFFFF', fg: 'FF000000' },
        };

        // Helpers
        const sectionHeader = (rowIdx, colSpan, text, bgColor) => {
            summarySheet.mergeCells(rowIdx, 1, rowIdx, colSpan);
            const cell = summarySheet.getCell(rowIdx, 1);
            cell.value = text;
            cell.font = { bold: true, size: 12, color: { argb: COLOR.WHITE } };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
            cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
            summarySheet.getRow(rowIdx).height = 24;
        };

        const kpiRow = (rowIdx, label, value, numFmt, valueColor) => {
            const labelCell = summarySheet.getCell(rowIdx, 1);
            const valueCell = summarySheet.getCell(rowIdx, 2);
            summarySheet.mergeCells(rowIdx, 3, rowIdx, 5);
            labelCell.value = label;
            valueCell.value = value;
            labelCell.font = { size: 10, color: { argb: COLOR.DARK } };
            valueCell.numFmt = numFmt || '#,##0';
            valueCell.alignment = { horizontal: 'right' };
            if (valueColor) valueCell.font = { bold: true, color: { argb: valueColor } };
            [labelCell, valueCell].forEach(c => {
                c.border = { bottom: { style: 'hair', color: { argb: COLOR.GRAY_200 } } };
            });
            labelCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 2 };
            summarySheet.getRow(rowIdx).height = 18;
        };

        const tableHeaderRow = (rowIdx, headers) => {
            headers.forEach((h, i) => {
                const cell = summarySheet.getCell(rowIdx, i + 1);
                cell.value = h;
                cell.font = { bold: true, size: 10, color: { argb: COLOR.DARK } };
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR.INDIGO_100 } };
                cell.alignment = { horizontal: i === 0 ? 'left' : 'right', vertical: 'middle', indent: i === 0 ? 2 : 0, wrapText: true };
                cell.border = {
                    top: { style: 'thin', color: { argb: COLOR.INDIGO_700 } },
                    bottom: { style: 'medium', color: { argb: COLOR.INDIGO_700 } },
                    left:  { style: 'hair', color: { argb: COLOR.GRAY_200 } },
                    right: { style: 'hair', color: { argb: COLOR.GRAY_200 } },
                };
            });
            summarySheet.getRow(rowIdx).height = 22;
        };

        // Set column widths
        summarySheet.getColumn(1).width = 48;
        summarySheet.getColumn(2).width = 20;
        summarySheet.getColumn(3).width = 18;
        summarySheet.getColumn(4).width = 16;
        summarySheet.getColumn(5).width = 55;

        // ──────────────────────────────────────────────────────────
        // TITLE
        // ──────────────────────────────────────────────────────────
        summarySheet.mergeCells('A1:E1');
        const titleCell = summarySheet.getCell('A1');
        titleCell.value = 'AUDIT FINDINGS SUMMARY — BRANCH ADVANCE ACCOUNT REVIEW';
        titleCell.font = { bold: true, size: 14, color: { argb: COLOR.WHITE } };
        titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR.DARK } };
        titleCell.alignment = { vertical: 'middle', horizontal: 'center' };
        summarySheet.getRow(1).height = 36;

        // Audit period subtitle
        summarySheet.mergeCells('A2:E2');
        const subCell = summarySheet.getCell('A2');
        const fmtDate = (d) => d instanceof Date ? d.toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' }) : '';
        subCell.value = `Audit Period: ${fmtDate(auditFromObj)} to ${fmtDate(auditToObj)}  |  Report Generated: ${fmtDate(new Date())}`;
        subCell.font = { italic: true, size: 10, color: { argb: 'FF6B7280' } };
        subCell.alignment = { vertical: 'middle', horizontal: 'center' };
        summarySheet.getRow(2).height = 18;

        let currentRow = 4;

        // ──────────────────────────────────────────────────────────
        // SECTION 1 — PORTFOLIO OVERVIEW
        // ──────────────────────────────────────────────────────────
        sectionHeader(currentRow, 5, '  SECTION 1 — PORTFOLIO OVERVIEW', COLOR.INDIGO_700);
        currentRow++;

        const totalAccounts = auditReportData.length;
        const totalOs = auditReportData.reduce((s, r) => s + parseAmount(r['Net O/s']), 0);
        const npaAccountsSummary = auditReportData.filter(r => r['Asset Classification'] === 'NPA');
        const npaCount = npaAccountsSummary.length;
        const npaOs = npaAccountsSummary.reduce((s, r) => s + parseAmount(r['Net O/s']), 0);
        const standardCount = totalAccounts - npaCount;
        const standardOs = totalOs - npaOs;

        const overdrawCount = auditReportData.filter(r => String(r['Overdrawn (Yes/No)'] || '').toUpperCase() === 'YES').length;
        const expiredLimitCount = auditReportData.filter(r => {
            const exp = parseDateString(r['Limit Expiry Date']);
            return exp && exp < auditToObj;
        }).length;
        const missingPanCount = auditReportData.filter(r => {
            const p = String(r['PAN'] || '').trim().toUpperCase();
            return !p || p === 'NA' || p.length < 10;
        }).length;
        const overdueCount = auditReportData.filter(r => parseAmount(r['Overdue Amount']) > 0).length;
        const totalOverdue = auditReportData.reduce((s, r) => s + parseAmount(r['Overdue Amount']), 0);
        const avgRiskScore = totalAccounts > 0
            ? Math.round(auditReportData.reduce((s, r) => s + (r['Risk Score'] || 0), 0) / totalAccounts)
            : 0;
        const portfolioRiskLabel = avgRiskScore >= 81 ? 'Critical' : avgRiskScore >= 61 ? 'High' : avgRiskScore >= 31 ? 'Medium' : 'Low';

        kpiRow(currentRow++, 'Total Number of Accounts', totalAccounts, '#,##0', null);
        kpiRow(currentRow++, 'Total Outstanding (Net O/s)', totalOs, '#,##0', null);
        kpiRow(currentRow++, 'NPA Accounts', npaCount, '#,##0', npaCount > 0 ? COLOR.RED_FG : null);
        kpiRow(currentRow++, 'NPA Outstanding', npaOs, '#,##0', npaOs > 0 ? COLOR.RED_FG : null);
        kpiRow(currentRow++, 'Standard Accounts', standardCount, '#,##0', null);
        kpiRow(currentRow++, 'Standard Outstanding', standardOs, '#,##0', null);
        kpiRow(currentRow++, 'Accounts Overdrawn (Beyond Sanction)', overdrawCount, '#,##0', overdrawCount > 0 ? COLOR.ORANGE_FG : null);
        kpiRow(currentRow++, 'Accounts with Expired Limit', expiredLimitCount, '#,##0', expiredLimitCount > 0 ? COLOR.ORANGE_FG : null);
        kpiRow(currentRow++, 'Accounts with Overdue Amount', overdueCount, '#,##0', overdueCount > 0 ? COLOR.AMBER_FG : null);
        kpiRow(currentRow++, 'Total Overdue Amount', totalOverdue, '#,##0', null);
        kpiRow(currentRow++, 'Accounts with Missing PAN', missingPanCount, '#,##0', missingPanCount > 0 ? COLOR.AMBER_FG : null);
        kpiRow(currentRow++, `Portfolio Average Risk Score (${portfolioRiskLabel})`, avgRiskScore, '#,##0', (riskColors[portfolioRiskLabel] || {}).fg || null);

        currentRow += 1; // spacer

        // ──────────────────────────────────────────────────────────
        // SECTION 2 — RISK DISTRIBUTION
        // ──────────────────────────────────────────────────────────
        sectionHeader(currentRow, 5, '  SECTION 2 — RISK DISTRIBUTION', COLOR.INDIGO_700);
        currentRow++;
        tableHeaderRow(currentRow, ['Risk Category', 'No. of Accounts', 'Outstanding (Rs.)', '% of Portfolio', 'Average Score']);
        currentRow++;

        ['Low', 'Medium', 'High', 'Critical', 'NPA - Low Risk', 'NPA - Medium Risk', 'NPA - High Risk', 'NPA - Critical'].forEach(label => {
            const group = auditReportData.filter(r => r['Risk Label'] === label);
            const cnt = group.length;
            if (cnt === 0 && label.startsWith('NPA')) return; // Hide empty NPA categories
            const os = group.reduce((s, r) => s + parseAmount(r['Net O/s']), 0);
            const pct = totalAccounts > 0 ? cnt / totalAccounts : 0;
            const avgScore = cnt > 0 ? Math.round(group.reduce((s, r) => s + (r['Risk Score'] || 0), 0) / cnt) : 0;

            const { bg, fg } = riskColors[label] || { bg: COLOR.GRAY_50, fg: COLOR.DARK };
            const cols = [
                { col: 1, val: label,    fmt: '@',      align: 'left' },
                { col: 2, val: cnt,      fmt: '#,##0',  align: 'right' },
                { col: 3, val: os,       fmt: '#,##0',  align: 'right' },
                { col: 4, val: pct,      fmt: '0.0%',   align: 'right' },
                { col: 5, val: avgScore, fmt: '#,##0',  align: 'right' },
            ];
            cols.forEach(({ col, val, fmt, align }) => {
                const cell = summarySheet.getCell(currentRow, col);
                cell.value = val;
                cell.numFmt = fmt;
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
                cell.font = { bold: col === 1, color: { argb: fg } };
                cell.alignment = { horizontal: align, indent: col === 1 ? 2 : 0, vertical: 'middle' };
                cell.border = {
                    bottom: { style: 'hair', color: { argb: COLOR.GRAY_200 } },
                    left:   { style: 'hair', color: { argb: COLOR.GRAY_200 } },
                    right:  { style: 'hair', color: { argb: COLOR.GRAY_200 } },
                };
            });
            summarySheet.getRow(currentRow).height = 18;
            currentRow++;
        });

        // Total Row for Risk Distribution
        {
            const makeTotal = (col, val, fmt) => {
                const c = summarySheet.getCell(currentRow, col);
                c.value = val;
                c.numFmt = fmt || '#,##0';
                c.font = { bold: true, color: { argb: COLOR.DARK } };
                c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR.GRAY_50 } };
                c.alignment = { horizontal: col === 1 ? 'left' : 'right', indent: col === 1 ? 2 : 0, vertical: 'middle' };
                c.border = {
                    top: { style: 'thin', color: { argb: COLOR.INDIGO_700 } },
                    bottom: { style: 'thin', color: { argb: COLOR.INDIGO_700 } },
                };
            };
            makeTotal(1, 'TOTAL', '@');
            makeTotal(2, totalAccounts);
            makeTotal(3, totalOs);
            makeTotal(4, 1, '0.0%');
            makeTotal(5, avgRiskScore);
            summarySheet.getRow(currentRow).height = 20;
            currentRow++;
        }

        currentRow += 1; // spacer

        // ──────────────────────────────────────────────────────────
        // SECTION 3 — KEY AUDIT FINDINGS
        // ──────────────────────────────────────────────────────────
        sectionHeader(currentRow, 5, '  SECTION 3 — KEY AUDIT FINDINGS', COLOR.INDIGO_700);
        currentRow++;

        const lakh = (n) => `Rs.${(n / 100000).toFixed(2)} Lakh`;

        const keyFindings = [];

        if (npaCount > 0) {
            keyFindings.push({ text: `${npaCount} account(s) classified as NPA with outstanding of ${lakh(npaOs)} — requires immediate provisioning review.`, severity: 'Critical' });
        }
        if (overdrawCount > 0) {
            keyFindings.push({ text: `${overdrawCount} account(s) are overdrawn beyond sanction limit — credit discipline risk.`, severity: 'High' });
        }
        if (expiredLimitCount > 0) {
            keyFindings.push({ text: `${expiredLimitCount} account(s) have expired sanction limits as on audit date — renewal/review action required.`, severity: 'High' });
        }
        if (overdueCount > 0) {
            keyFindings.push({ text: `${overdueCount} account(s) have overdue amounts aggregating ${lakh(totalOverdue)} — recovery follow-up needed.`, severity: 'Medium' });
        }

        // Security shortfall finding
        const secShortfallAccounts = auditReportData.filter(r => {
            const sl = parseAmount(r['Sanction Limit']);
            const ps = parseAmount(r['Primary Security']);
            const cs = parseAmount(r['Collateral Security']);
            return sl > 0 && (ps + cs) < sl * 0.75;
        });
        if (secShortfallAccounts.length > 0) {
            keyFindings.push({ text: `${secShortfallAccounts.length} account(s) have combined security coverage below 75% of sanction limit — collateral adequacy concern.`, severity: 'Medium' });
        }

        if (missingPanCount > 0) {
            keyFindings.push({ text: `${missingPanCount} account(s) have missing/incomplete PAN — KYC compliance gap.`, severity: 'Medium' });
        }

        // High utilisation standard accounts (≥90%, not overdrawn)
        const highUtilAccounts = auditReportData.filter(r => {
            const sl = parseAmount(r['Sanction Limit']);
            const nos = parseAmount(r['Net O/s']);
            const util = sl > 0 ? nos / sl : 0;
            return sl > 0 && util >= 0.9 && util < 1.0 && r['Asset Classification'] === 'Standard';
        });
        if (highUtilAccounts.length > 0) {
            keyFindings.push({ text: `${highUtilAccounts.length} standard account(s) with utilisation >= 90% — monitor closely for potential slippage.`, severity: 'Low' });
        }

        // Combined critical/high finding
        const critHighAccounts = auditReportData.filter(r => r['Risk Label'] === 'Critical' || r['Risk Label'] === 'High');
        if (critHighAccounts.length > 0) {
            const critHighOs = critHighAccounts.reduce((s, r) => s + parseAmount(r['Net O/s']), 0);
            keyFindings.push({ text: `${critHighAccounts.length} account(s) rated High/Critical risk with combined outstanding of ${lakh(critHighOs)} — require priority audit attention.`, severity: 'High' });
        }

        keyFindings.push({ text: `Overall portfolio average risk score: ${avgRiskScore}/100 — ${portfolioRiskLabel} risk portfolio.`, severity: portfolioRiskLabel });

        if (keyFindings.length === 0) {
            keyFindings.push({ text: 'No major adverse findings observed in the advance portfolio as on audit date.', severity: 'Low' });
        }

        const severityPrefix = { Critical: '[CRITICAL]', High: '[HIGH]', Medium: '[MEDIUM]', Low: '[LOW]' };
        keyFindings.forEach((finding, idx) => {
            summarySheet.mergeCells(currentRow, 1, currentRow, 5);
            const cell = summarySheet.getCell(currentRow, 1);
            cell.value = `  ${idx + 1}.  ${finding.text}`;
            const { bg, fg } = riskColors[finding.severity] || { bg: COLOR.GRAY_50, fg: COLOR.DARK };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
            cell.font = { size: 10, color: { argb: fg } };
            cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true, indent: 1 };
            cell.border = { bottom: { style: 'hair', color: { argb: COLOR.GRAY_200 } } };
            summarySheet.getRow(currentRow).height = 22;
            currentRow++;
        });

        currentRow += 1; // spacer

        // ──────────────────────────────────────────────────────────
        // SECTION 4 — TOP HIGH-RISK ACCOUNTS
        // ──────────────────────────────────────────────────────────
        sectionHeader(currentRow, 5, '  SECTION 4 — TOP HIGH-RISK STANDARD ACCOUNTS', COLOR.INDIGO_700);
        currentRow++;
        tableHeaderRow(currentRow, ['A/c No. / Borrower Name', 'Net O/s (Rs.)', 'Risk Score', 'Risk Label', 'Audit Remark']);
        currentRow++;

        const topRiskAccounts = [...auditReportData]
            .filter(row => !String(row['Asset Classification'] || '').toUpperCase().includes('NPA'))
            .sort((a, b) => (b['Risk Score'] || 0) - (a['Risk Score'] || 0))
            .slice(0, 15);

        topRiskAccounts.forEach(row => {
            const label = row['Risk Label'] || 'Low';

            const nameCell = summarySheet.getCell(currentRow, 1);
            const acNo = String(row['A/c No.'] || row['CIF'] || '');
            const borrower = String(row['Borrower Name'] || '');
            nameCell.value = acNo ? `${acNo} — ${borrower}` : borrower;
            nameCell.font = { size: 9 };
            nameCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1, wrapText: true };

            const osCell = summarySheet.getCell(currentRow, 2);
            osCell.value = parseAmount(row['Net O/s']);
            osCell.alignment = { vertical: 'middle', horizontal: 'right' };
            osCell.font = { size: 9 };

            const scoreCell = summarySheet.getCell(currentRow, 3);
            scoreCell.value = row['Risk Score'] || 0;
            scoreCell.numFmt = '0';
            scoreCell.font = { bold: true, size: 10, color: { argb: 'FF000000' } };
            scoreCell.alignment = { vertical: 'middle', horizontal: 'center' };

            const labelCell = summarySheet.getCell(currentRow, 4);
            labelCell.value = label;
            labelCell.font = { bold: true, color: { argb: 'FF000000' }, size: 9 };
            labelCell.alignment = { vertical: 'middle', horizontal: 'center' };

            const remarkCell = summarySheet.getCell(currentRow, 5);
            remarkCell.value = row['Audit Remark'] || '';
            remarkCell.font = { italic: true, size: 9 };
            remarkCell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };

            [nameCell, osCell, scoreCell, labelCell, remarkCell].forEach(c => {
                c.border = { bottom: { style: 'hair', color: { argb: COLOR.GRAY_200 } } };
            });
            summarySheet.getRow(currentRow).height = 28;
            currentRow++;
        });

        if (topRiskAccounts.length === 0) {
            summarySheet.mergeCells(currentRow, 1, currentRow, 5);
            summarySheet.getCell(currentRow, 1).value = 'No standard accounts to display.';
            currentRow++;
        }

        currentRow += 1;
        sectionHeader(currentRow, 5, '  SECTION 5 — TOP CRITICAL NPA ACCOUNTS', COLOR.INDIGO_700);
        currentRow++;
        tableHeaderRow(currentRow, ['A/c No. / Borrower Name', 'Net O/s (Rs.)', 'Risk Score', 'Risk Label', 'Audit Remark']);
        currentRow++;

        const topNpaAccounts = [...auditReportData]
            .filter(row => String(row['Asset Classification'] || '').toUpperCase().includes('NPA') && String(row['Risk Label'] || '').includes('NPA - Critical'))
            .sort((a, b) => (b['Risk Score'] || 0) - (a['Risk Score'] || 0))
            .slice(0, 15);

        topNpaAccounts.forEach(row => {
            const label = row['Risk Label'] || 'NPA - Critical';

            const nameCell = summarySheet.getCell(currentRow, 1);
            const acNo = String(row['A/c No.'] || row['CIF'] || '');
            const borrower = String(row['Borrower Name'] || '');
            nameCell.value = acNo ? `${acNo} — ${borrower}` : borrower;
            nameCell.font = { size: 9 };
            nameCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1, wrapText: true };

            const osCell = summarySheet.getCell(currentRow, 2);
            osCell.value = parseAmount(row['Net O/s']);
            osCell.alignment = { vertical: 'middle', horizontal: 'right' };
            osCell.font = { size: 9 };

            const scoreCell = summarySheet.getCell(currentRow, 3);
            scoreCell.value = row['Risk Score'] || 0;
            scoreCell.numFmt = '0';
            scoreCell.font = { bold: true, size: 10, color: { argb: 'FF000000' } };
            scoreCell.alignment = { vertical: 'middle', horizontal: 'center' };

            const labelCell = summarySheet.getCell(currentRow, 4);
            labelCell.value = label;
            labelCell.font = { bold: true, color: { argb: 'FF000000' }, size: 9 };
            labelCell.alignment = { vertical: 'middle', horizontal: 'center' };

            const remarkCell = summarySheet.getCell(currentRow, 5);
            remarkCell.value = row['Audit Remark'] || '';
            remarkCell.font = { italic: true, size: 9 };
            remarkCell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };

            [nameCell, osCell, scoreCell, labelCell, remarkCell].forEach(c => {
                c.border = { bottom: { style: 'hair', color: { argb: COLOR.GRAY_200 } } };
            });
            summarySheet.getRow(currentRow).height = 28;
            currentRow++;
        });

        if (topNpaAccounts.length === 0) {
            summarySheet.mergeCells(currentRow, 1, currentRow, 5);
            summarySheet.getCell(currentRow, 1).value = 'No critical NPA accounts to display.';
            currentRow++;
        }
    }
    // ─────────────────────────────────────────────────────────────────────────────
    // END Audit Findings Summary
    // ─────────────────────────────────────────────────────────────────────────────


    if (targetFields.includes('Loan Type') && targetFields.includes('ROI')) {
        const roiAnalysisSheet = workbook.addWorksheet('ROI Analysis');

        // Calculate aggregations
        const roiStats = {}; // { 'Loan Type A': { min: 100, max: 0 } }

        processedData.forEach(row => {
            const loanType = row['Loan Type'];
            const roi = row['ROI'];
            const assetClassification = row['Asset Classification'];

            // Only aggregate if Loan Type exists, ROI is a valid number, and it's a Standard account
            if (loanType && typeof roi === 'number' && assetClassification === 'Standard') {
                if (!roiStats[loanType]) {
                    roiStats[loanType] = { min: roi, max: roi };
                } else {
                    if (roi < roiStats[loanType].min) roiStats[loanType].min = roi;
                    if (roi > roiStats[loanType].max) roiStats[loanType].max = roi;
                }
            }
        });

        // Prepare data for the sheet
        const roiAnalysisData = Object.keys(roiStats).map(loanType => {
            const min = roiStats[loanType].min;
            const max = roiStats[loanType].max;
            const diff = max - min;
            return {
                'Loan Type': loanType,
                'Min ROI': min,
                'Max ROI': max,
                'Difference': diff
            };
        });

        // Define Columns for ROI Analysis
        roiAnalysisSheet.columns = [
            { header: 'Loan Type', key: 'Loan Type', width: 30 },
            { header: 'Min ROI', key: 'Min ROI', width: 20 },
            { header: 'Max ROI', key: 'Max ROI', width: 20 },
            { header: 'Difference', key: 'Difference', width: 20 }
        ];

        // Add Data
        roiAnalysisSheet.addRows(roiAnalysisData);

        // Styling the Header Row
        const roiHeaderRow = roiAnalysisSheet.getRow(1);
        roiHeaderRow.height = 30;
        roiHeaderRow.eachCell((cell) => {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFE6E6FA' } // Pastel Lavender
            };
            cell.font = {
                bold: true,
                color: { argb: 'FF483D8B' }, // Dark Slate Blue
                size: 12
            };
            cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
            cell.border = {
                top: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                left: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                right: { style: 'thin', color: { argb: 'FFCCCCCC' } }
            };
        });

        // Styling the Data Rows and highlighting > 3% difference
        roiAnalysisSheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return;

            // Apply Difference Formula
            const minCellRef = `B${rowNumber}`;
            const maxCellRef = `C${rowNumber}`;
            const diffCol = 4; // Difference is column D
            const fallbackResult = roiAnalysisData[rowNumber - 2]['Difference'];
            row.getCell(diffCol).value = { formula: `${maxCellRef}-${minCellRef}`, result: fallbackResult };

            const isHighDiff = fallbackResult > 0.03; // Note: 3% is 0.03 in decimal if we stored it as fraction

            row.eachCell((cell, colNumber) => {
                const fieldName = roiAnalysisSheet.columns[colNumber - 1] ? roiAnalysisSheet.columns[colNumber - 1].key : '';
                if (!fieldName) return;

                cell.alignment = { vertical: 'middle', horizontal: colNumber === 1 ? 'left' : 'right' };
                cell.border = {
                    top: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                    left: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                    bottom: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                    right: { style: 'thin', color: { argb: 'FFEEEEEE' } }
                };

                // Number Formatting for percentages
                if (['Min ROI', 'Max ROI', 'Difference'].includes(fieldName)) {
                    cell.numFmt = '0.00%';
                }

                // Conditional formatting highlight (Red-ish pastel for high diff)
                if (isHighDiff) {
                    cell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FFFFEBEE' } // Light red/pink
                    };
                    cell.font = {
                        color: { argb: 'FFD32F2F' }, // Dark red text
                        bold: true
                    };
                }
            });
        });
    }
    const nonNpaTargetFields = baseAuditReportFields || targetFields;
    // --- Generate CC and TL Sheets ---
    // Check both Facility Code and Loan Type to catch all accounts regardless of mapping setup
    if (targetFields.includes('Facility Code') || targetFields.includes('Loan Type')) {
        let ccData = processedData.filter(row => {
            const code = String(row['Facility Code'] || '').toUpperCase();
            const loanType = String(row['Loan Type'] || '').toUpperCase();
            // Only Cash Credit - explicitly exclude OD/Overdraft
            return code === 'CC' || code.includes('CASH CREDIT')
                || loanType.includes('CASH CREDIT') || loanType === 'CC';
        });

        // Custom sorting for CC Data: Sort the Net O/s from Highest to Lowest after setting off Credit Balance
        if (ccData.length > 0 && targetFields.includes('Net O/s')) {
            ccData.sort((a, b) => {
                const netOsA = parseAmount(a['Net O/s']);
                const creditBala = targetFields.includes('Credit Balance') ? parseAmount(a['Credit Balance']) : 0;
                const effectiveOsA = netOsA - creditBala;

                const netOsB = parseAmount(b['Net O/s']);
                const creditBalb = targetFields.includes('Credit Balance') ? parseAmount(b['Credit Balance']) : 0;
                const effectiveOsB = netOsB - creditBalb;

                return effectiveOsB - effectiveOsA;
            });
        }

        // Apply netting off to the data itself
        if (ccData.length > 0 && targetFields.includes('Net O/s') && targetFields.includes('Credit Balance')) {
            ccData = ccData.map(row => {
                const netOs = parseAmount(row['Net O/s']);
                const creditBal = parseAmount(row['Credit Balance']);
                const nettedOs = netOs - creditBal;
                // We keep the netted amount in 'Net O/s', and we can clear 'Credit Balance'
                return {
                    ...row,
                    'Net O/s': nettedOs,
                    'Credit Balance': 0 // or remove it depending on logic below
                };
            });
        }

        // Adjust columns for CC Data: Overdue amount column should be after overdrawn column.
        let ccFieldsToUse = [...nonNpaTargetFields];

        // Remove excluded fields from CC output
        const ccExcludedFields = ['Credit Balance', 'A/c Open Date', 'PAN', 'Aadhar No.', 'PS Flag', 'NPA Date', 'NPA Provision', 'Outstanding EMIs', 'Theo EMI', 'Inst Due', 'Tenure'];
        ccExcludedFields.forEach(f => {
            const idx = ccFieldsToUse.indexOf(f);
            if (idx !== -1) ccFieldsToUse.splice(idx, 1);
        });

        const overdueIdx = ccFieldsToUse.indexOf('Overdue Amount');
        const overdrawnIdx = ccFieldsToUse.indexOf('Overdrawn (Yes/No)');

        if (overdueIdx !== -1 && overdrawnIdx !== -1) {
            // Remove Overdue Amount from its current position
            ccFieldsToUse.splice(overdueIdx, 1);
            // Find the new index of Overdrawn (Yes/No) since the array was modified
            const newOverdrawnIdx = ccFieldsToUse.indexOf('Overdrawn (Yes/No)');
            // Insert Overdue Amount right after Overdrawn (Yes/No)
            ccFieldsToUse.splice(newOverdrawnIdx + 1, 0, 'Overdue Amount');
        }

        // Adjust columns for TL Data: Overdue amount column should be after overdrawn column.
        let tlFieldsToUse = [...nonNpaTargetFields];

        // Remove excluded fields from TL output
        const tlExcludedFields = ['Credit Balance', 'A/c Open Date', 'PAN', 'Aadhar No.', 'PS Flag', 'NPA Date', 'NPA Provision'];
        tlExcludedFields.forEach(f => {
            const idx = tlFieldsToUse.indexOf(f);
            if (idx !== -1) tlFieldsToUse.splice(idx, 1);
        });

        const tlOverdueIdx = tlFieldsToUse.indexOf('Overdue Amount');
        const tlOverdrawnIdx = tlFieldsToUse.indexOf('Overdrawn (Yes/No)');

        if (tlOverdueIdx !== -1 && tlOverdrawnIdx !== -1) {
            // Remove Overdue Amount from its current position
            tlFieldsToUse.splice(tlOverdueIdx, 1);
            // Find new index of Overdrawn (Yes/No)
            const newTlOverdrawnIdx = tlFieldsToUse.indexOf('Overdrawn (Yes/No)');
            // Insert Overdue Amount right after Overdrawn (Yes/No)
            tlFieldsToUse.splice(newTlOverdrawnIdx + 1, 0, 'Overdue Amount');
        }

        // For TL data, calculate the Theo vs CBS requirements row by row
        const tlFieldsExtension = [
            'Tenure (Months)',
            'Instalments Due',
            'Theo EMI',
            'Theo Principal Paid',
            'Theo Interest Charged',
            'Theo O/s Balance',
            'O/s Difference',
            'Outstanding EMIs',
            'Asset Class (Theo Bal)'
        ];

        const tlExtendedFields = [...tlFieldsToUse, ...tlFieldsExtension];

        const tlData = processedData.filter(row => {
            const code = String(row['Facility Code'] || '').toUpperCase();
            const loanType = String(row['Loan Type'] || '').toUpperCase();
            const isTL = code === 'TL' || code.includes('TERM LOAN') || code.includes('HL') || code.includes('VL') || code.includes('PL') || code.includes('EDU')
                || loanType.includes('TERM LOAN') || loanType === 'TL'
                || loanType.includes('HOUSING') || loanType.includes('VEHICLE') || loanType.includes('PERSONAL') || loanType.includes('MSME');
            if (!isTL) return false;

            // TL specific calculations
            const principal = parseAmount(row['Sanction Limit']);
            const roi = typeof row['ROI'] === 'number' ? row['ROI'] * 100 : parseAmount(row['ROI']); // calculation utility expects e.g. 8.5
            const sanctionDate = parseDateString(row['Sanction Date']);
            const expiryDate = parseDateString(row['Limit Expiry Date']);
            const cbsOutstanding = Math.max(0, parseAmount(row['Net O/s']));

            const ten = calculateTenure(sanctionDate, expiryDate);
            const instDue = getInstalmentsDue(sanctionDate, auditToObj);

            const theoEmi = calculateEMI(principal, roi, ten);

            // Breakdown
            const bk = getAmortizationBreakdown(principal, roi, ten, instDue);

            // Netting off before calculating irregularity (just for data export prep)
            const creditBalance = targetFields.includes('Credit Balance') ? parseAmount(row['Credit Balance']) : 0;
            const netOs = parseAmount(row['Net O/s']);
            const nettedOs = netOs - creditBalance;

            // Difference and EMIs
            const osDiff = Math.max(0, nettedOs - bk.theoBalance);
            const emisOutstanding = (theoEmi > 0 && osDiff > 0) ? Math.floor(osDiff / theoEmi) : 0;

            // Asset Class mapped
            let bkTag = "Standard";
            if (emisOutstanding <= 1) bkTag = "Standard";
            else if (emisOutstanding <= 2) bkTag = "SMA-0";
            else if (emisOutstanding <= 3) bkTag = "SMA-1";
            else if (emisOutstanding <= 4) bkTag = "SMA-2";
            else bkTag = "NPA";

            // Append extensions to the row object for the TL sheet export
            // Replace Net O/s with the netted version, clear Credit Balance
            row['Net O/s'] = nettedOs;
            row['Credit Balance'] = 0;

            row['Tenure (Months)'] = ten;
            row['Instalments Due'] = instDue;
            row['Theo EMI'] = theoEmi;
            row['Theo Principal Paid'] = bk.principalPaid;
            row['Theo Interest Charged'] = bk.interestCharged;
            row['Theo O/s Balance'] = bk.theoBalance;
            row['O/s Difference'] = osDiff;
            row['Outstanding EMIs'] = emisOutstanding;
            row['Asset Class (Theo Bal)'] = bkTag;

            return true;
        });

        // TL Custom Sorting: Sort by Netted Net O/s (Highest to Lowest)
        if (tlData.length > 0 && tlFieldsToUse.includes('Net O/s')) {
            tlData.sort((a, b) => b['Net O/s'] - a['Net O/s']);
        }

        const styleSheet = (sheetName, data, customFields = null) => {
            if (data.length === 0) return;
            const fieldsToUse = customFields || targetFields;

            const sheet = workbook.addWorksheet(sheetName);
            sheet.columns = fieldsToUse.map(field => {
                let width = 20;
                if (field === 'Borrower Name') width = 35;
                if (field === 'CIF' || field === 'A/c No.' || field === 'PAN' || field === 'Aadhar No.') width = 18;
                if (field.includes('Date')) width = 15;
                return { header: field, key: field, width };
            });

            sheet.addRows(data);

            // Fetch column keys for calculations if it's the TL sheet
            let sanctionDateCol, expiryDateCol, sanctionLimitCol, roiCol, netOsCol;
            let tenureCol, distDueCol, theoEmiCol, theoPrinCol, theoIntCol, theoOsCol, diffCol, emiOutCol, assetTheoCol;

            if (sheetName === 'TL Accounts') {
                const getColIdx = (name) => fieldsToUse.indexOf(name) + 1;
                sanctionDateCol = getColumnLetter(getColIdx('Sanction Date'));
                expiryDateCol = getColumnLetter(getColIdx('Limit Expiry Date'));
                sanctionLimitCol = getColumnLetter(getColIdx('Sanction Limit'));
                roiCol = getColumnLetter(getColIdx('ROI'));
                netOsCol = getColumnLetter(getColIdx('Net O/s'));

                tenureCol = getColIdx('Tenure (Months)');
                distDueCol = getColIdx('Instalments Due');
                theoEmiCol = getColIdx('Theo EMI');
                theoPrinCol = getColIdx('Theo Principal Paid');
                theoIntCol = getColIdx('Theo Interest Charged');
                theoOsCol = getColIdx('Theo O/s Balance');
                diffCol = getColIdx('O/s Difference');
                emiOutCol = getColIdx('Outstanding EMIs');
                assetTheoCol = getColIdx('Asset Class (Theo Bal)');
            }

            const sHeaderRow = sheet.getRow(1);
            sHeaderRow.height = 30;
            sHeaderRow.eachCell((cell) => {
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFE6E6FA' }
                };
                cell.font = { bold: true, color: { argb: 'FF483D8B' }, size: 12 };
                cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                cell.border = {
                    top: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                    left: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                    bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                    right: { style: 'thin', color: { argb: 'FFCCCCCC' } }
                };
            });

            sheet.eachRow((row, rowNumber) => {
                if (rowNumber === 1) return;

                if (sheetName === 'TL Accounts') {
                    const fallbackData = data[rowNumber - 2];
                    const sDateRef = `${sanctionDateCol}${rowNumber}`;
                    const expDateRef = `${expiryDateCol}${rowNumber}`;
                    const limRef = `${sanctionLimitCol}${rowNumber}`;
                    const roiRef = `${roiCol}${rowNumber}`;
                    const netOsRef = `${netOsCol}${rowNumber}`;
                    const auditDateRef = 'Parameters!$B$3'; // Row 3 in Parameters sheet is Audit To Date

                    // Tenure (Months) = DATEDIF using YEAR and MONTH
                    const tenureRef = `${getColumnLetter(tenureCol)}${rowNumber}`;
                    const tenureForm = `MAX(0, (YEAR(${expDateRef})-YEAR(${sDateRef}))*12 + MONTH(${expDateRef})-MONTH(${sDateRef}))`;
                    row.getCell(tenureCol).value = { formula: tenureForm, result: fallbackData['Tenure (Months)'] };

                    // Instalments Due = MAX(0, calculated months + 1)
                    const dueRef = `${getColumnLetter(distDueCol)}${rowNumber}`;
                    const dueForm = `MAX(0, (YEAR(${auditDateRef})-YEAR(${sDateRef}))*12 + MONTH(${auditDateRef})-MONTH(${sDateRef}) + 1)`;
                    row.getCell(distDueCol).value = { formula: dueForm, result: fallbackData['Instalments Due'] };

                    // Theo EMI = PMT(rate, nper, pv)
                    const emiRef = `${getColumnLetter(theoEmiCol)}${rowNumber}`;
                    const emiForm = `IFERROR(PMT(${roiRef}/12, ${tenureRef}, -${limRef}), 0)`;
                    row.getCell(theoEmiCol).value = { formula: emiForm, result: fallbackData['Theo EMI'] };

                    // Theo Principal Paid
                    const prinRef = `${getColumnLetter(theoPrinCol)}${rowNumber}`;
                    const prinForm = `IF(MIN(${dueRef}, ${tenureRef})<=0, 0, IFERROR(IF(${roiRef}=0, (${limRef}/${tenureRef})*MIN(${dueRef}, ${tenureRef}), -CUMPRINC(${roiRef}/12, ${tenureRef}, ${limRef}, 1, MIN(${dueRef}, ${tenureRef}), 0)), 0))`;
                    row.getCell(theoPrinCol).value = { formula: prinForm, result: fallbackData['Theo Principal Paid'] };

                    // Theo Interest Charged
                    const intForm = `IFERROR(IF(${roiRef}=0, 0, -CUMIPMT(${roiRef}/12, ${tenureRef}, ${limRef}, 1, MIN(${dueRef}, ${tenureRef}), 0)), 0)`;
                    row.getCell(theoIntCol).value = { formula: intForm, result: fallbackData['Theo Interest Charged'] };

                    // Theo O/s Balance = Lim - Prin
                    const osRef = `${getColumnLetter(theoOsCol)}${rowNumber}`;
                    const osForm = `MAX(0, ${limRef} - ${prinRef})`;
                    row.getCell(theoOsCol).value = { formula: osForm, result: fallbackData['Theo O/s Balance'] };

                    // O/s Difference = MAX(0, Net O/s - Theo O/s)
                    const diffRef = `${getColumnLetter(diffCol)}${rowNumber}`;
                    const customDiffForm = `MAX(0, ${netOsRef} - ${osRef})`;
                    row.getCell(diffCol).value = { formula: customDiffForm, result: fallbackData['O/s Difference'] };

                    // Outstanding EMIs = FLOOR(Diff / EMI, 1) or 0
                    const emiOutRef = `${getColumnLetter(emiOutCol)}${rowNumber}`;
                    const emiOutForm = `IF(AND(${diffRef}>0, ${emiRef}>0), FLOOR(${diffRef}/${emiRef}, 1), 0)`;
                    row.getCell(emiOutCol).value = { formula: emiOutForm, result: fallbackData['Outstanding EMIs'] };

                    // Irregularity -> Asset Class (driven by Outstanding EMIs)
                    const getAssetForm = () => {
                        return `IF(${emiOutRef}<=1, "Standard", IF(${emiOutRef}<=2, "SMA-0", IF(${emiOutRef}<=3, "SMA-1", IF(${emiOutRef}<=4, "SMA-2", "NPA"))))`;
                    };

                    row.getCell(assetTheoCol).value = {
                        formula: getAssetForm(),
                        result: fallbackData['Asset Class (Theo Bal)']
                    };
                }

                let shouldHighlight = false;
                if (sheetName === 'TL Accounts') {
                    const fallbackData = data[rowNumber - 2];
                    const bankClass = String(fallbackData['Asset Classification'] || '').toUpperCase();
                    const theoClass = String(fallbackData['Asset Class (Theo Bal)'] || '').toUpperCase();

                    const isBankStd = bankClass.includes('STANDARD') || bankClass === 'STD';
                    const isTheoNonStd = theoClass && !theoClass.includes('STANDARD') && theoClass !== 'STD';

                    if (isBankStd && isTheoNonStd) {
                        shouldHighlight = true;
                    }
                }

                row.eachCell((cell, colNumber) => {
                    const fieldName = fieldsToUse[colNumber - 1];

                    cell.alignment = { vertical: 'middle', horizontal: 'left' };
                    cell.border = {
                        top: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                        left: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                        bottom: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                        right: { style: 'thin', color: { argb: 'FFEEEEEE' } }
                    };

                    if (shouldHighlight) {
                        cell.fill = {
                            type: 'pattern',
                            pattern: 'solid',
                            fgColor: { argb: 'FFFFE4E1' } // Lightest pastel red/pink (MistyRose)
                        };
                    }

                    if (['Sanction Limit', 'Net O/s', 'Credit Balance', 'Drawing Power', 'Primary Security', 'Collateral Security', 'Overdue Amount', 'Theo EMI', 'Theo Principal Paid', 'Theo Interest Charged', 'Theo O/s Balance', 'O/s Difference'].includes(fieldName)) {
                        cell.numFmt = '#,##0';
                        cell.alignment = { vertical: 'middle', horizontal: 'right' };
                    } else if (fieldName === 'ROI') {
                        cell.numFmt = '0.00%';
                        cell.alignment = { vertical: 'middle', horizontal: 'right' };
                    } else if (['Sanction Date', 'Limit Expiry Date', 'A/c Open Date', 'NPA Date'].includes(fieldName)) {
                        cell.numFmt = 'DD-MM-YY';
                        cell.alignment = { vertical: 'middle', horizontal: 'center' };
                    } else if (['CIF', 'A/c No.', 'PAN', 'Aadhar No.'].includes(fieldName)) {
                        cell.numFmt = '@';
                    } else if (['Overdrawn (Yes/No)', 'Sanction Type', 'PS Flag', 'Loan Type', 'Asset Classification', 'Asset Class (Theo Bal)'].includes(fieldName)) {
                        cell.alignment = { vertical: 'middle', horizontal: 'center' };
                    } else if (['Tenure (Months)', 'Instalments Due', 'Outstanding EMIs'].includes(fieldName)) {
                        cell.alignment = { vertical: 'middle', horizontal: 'right' };
                    }
                });
            });
        };

        styleSheet('CC Accounts', ccData, ccFieldsToUse);
        styleSheet('TL Accounts', tlData, tlExtendedFields);
    }

    // --- Generate Advance Stratification Sheet ---
    const stratificationSheet = workbook.addWorksheet('Advance Stratification');

    // Column letters from Audit Report for formulas (with fallbacks)
    const getSafeAuditCol = (field) => {
        const idx = auditReportFields.indexOf(field);
        return getColumnLetter(idx === -1 ? 1 : idx + 1);
    };

    const rNetOsCol = getSafeAuditCol('Net O/s');
    const rLoanTypeCol = getSafeAuditCol('Loan Type');
    const rRoiCol = getSafeAuditCol('ROI');
    const rAssetClassCol = getSafeAuditCol('Asset Classification');
    const rEmiCol = getSafeAuditCol('Outstanding EMIs');
    const auditSheetName = "'Audit Report'";

    // Dashboard Title
    stratificationSheet.mergeCells('A1:S1');
    const mainTitle = stratificationSheet.getCell('A1');
    mainTitle.value = 'ADVANCE PORTFOLIO STRATIFICATION DASHBOARD';
    mainTitle.font = { bold: true, size: 16, color: { argb: 'FF1E1B4B' } }; // Indigo 950
    mainTitle.alignment = { horizontal: 'center', vertical: 'middle' };
    stratificationSheet.getRow(1).height = 40;

    const addStratTable = (title, columns, data, formulaGenerator, startColIdx) => {

        const isEmpty = !data || data.length === 0 || data.every(row => row.count === 0 && row.sum === 0);
        if (isEmpty) {
            stratificationSheet.getColumn(startColIdx).width = 2;
            stratificationSheet.getColumn(startColIdx + 1).width = 2;
            stratificationSheet.getColumn(startColIdx + 2).width = 2;
            return;
        }
        stratificationSheet.getColumn(startColIdx).width = 25;
        stratificationSheet.getColumn(startColIdx + 1).width = 15;
        stratificationSheet.getColumn(startColIdx + 2).width = 25;
        const startColChar = getColumnLetter(startColIdx);
        const endColChar = getColumnLetter(startColIdx + 2);

        // Table Title (Merged)
        stratificationSheet.mergeCells(`${startColChar}3:${endColChar}3`);
        const titleCell = stratificationSheet.getCell(`${startColChar}3`);
        titleCell.value = title;
        titleCell.font = { bold: true, size: 11, color: { argb: 'FFFFFFFF' } };
        titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F46E5' } }; // Indigo 600
        titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
        stratificationSheet.getRow(3).height = 25;

        // Table Headers
        const headerRowIdx = 4;
        columns.forEach((colName, i) => {
            const cell = stratificationSheet.getCell(headerRowIdx, startColIdx + i);
            cell.value = colName;
            cell.font = { bold: true, size: 10, color: { argb: 'FF374151' } };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3F4F6' } };
            cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            cell.border = {
                top: { style: 'thin', color: { argb: 'FFD1D5DB' } },
                left: { style: 'thin', color: { argb: 'FFD1D5DB' } },
                bottom: { style: 'medium', color: { argb: 'FF4F46E5' } },
                right: { style: 'thin', color: { argb: 'FFD1D5DB' } }
            };
        });
        stratificationSheet.getRow(4).height = 30;

        // Table Data and Formulas
        let currentRow = 5;
        data.forEach((item, idx) => {
            const row = stratificationSheet.getRow(currentRow);
            const labelCell = row.getCell(startColIdx);
            const countCell = row.getCell(startColIdx + 1);
            const sumCell = row.getCell(startColIdx + 2);

            labelCell.value = item.label;
            countCell.value = { formula: formulaGenerator(item, 'count'), result: item.count };
            sumCell.value = { formula: formulaGenerator(item, 'sum'), result: item.sum };

            // Alignment
            labelCell.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
            countCell.alignment = { horizontal: 'center', vertical: 'middle' };
            sumCell.alignment = { horizontal: 'right', vertical: 'middle' };
            sumCell.numFmt = '#,##0';

            // Borders
            [labelCell, countCell, sumCell].forEach(cell => {
                cell.border = {
                    left: { style: 'thin', color: { argb: 'FFD1D5DB' } },
                    right: { style: 'thin', color: { argb: 'FFD1D5DB' } },
                    bottom: { style: 'thin', color: { argb: 'FFF3F4F6' } }
                };
            });

            currentRow++;
        });

        // Total Row
        const totalRow = stratificationSheet.getRow(currentRow);
        const totalLabel = totalRow.getCell(startColIdx);
        const totalCount = totalRow.getCell(startColIdx + 1);
        const totalSum = totalRow.getCell(startColIdx + 2);

        totalLabel.value = 'TOTAL';
        totalCount.value = { formula: `SUM(${getColumnLetter(startColIdx + 1)}5:${getColumnLetter(startColIdx + 1)}${currentRow - 1})` };
        totalSum.value = { formula: `SUM(${getColumnLetter(startColIdx + 2)}5:${getColumnLetter(startColIdx + 2)}${currentRow - 1})` };

        totalSum.numFmt = '#,##0';
        [totalLabel, totalCount, totalSum].forEach(cell => {
            cell.font = { bold: true, size: 10, color: { argb: 'FF111827' } };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF9FAFB' } };
            cell.alignment = { vertical: 'middle', horizontal: cell === totalLabel ? 'left' : (cell === totalCount ? 'center' : 'right') };
            if (cell === totalLabel) cell.alignment.indent = 1;
            cell.border = {
                top: { style: 'medium', color: { argb: 'FFD1D5DB' } },
                left: { style: 'thin', color: { argb: 'FFD1D5DB' } },
                bottom: { style: 'medium', color: { argb: 'FFD1D5DB' } },
                right: { style: 'thin', color: { argb: 'FFD1D5DB' } }
            };
        });
    };

    // Prepare data (Loan Size)
    const loanSizeRanges = [
        { label: 'Upto 10 Lakhs', min: -99999999999, max: 1000000 },
        { label: '10 to 25 Lakhs', min: 1000000.01, max: 2500000 },
        { label: '25 to 50 Lakhs', min: 2500000.01, max: 5000000 },
        { label: '50 to 100 Lakhs', min: 5000000.01, max: 10000000 },
        { label: 'Above 100 Lakhs', min: 10000000.01, max: 99999999999 }
    ];
    loanSizeRanges.forEach(r => {
        const filtered = processedData.filter(d => {
            const os = parseAmount(d['Net O/s']);
            return os >= r.min && os <= r.max;
        });
        r.count = filtered.length;
        r.sum = filtered.reduce((acc, d) => acc + parseAmount(d['Net O/s']), 0);
    });
    addStratTable('I. BY LOAN SIZE', ['Range', 'Accounts', 'Outstanding'], loanSizeRanges, (item, type) => {
        const col = `$${rNetOsCol}`;
        const range = `${auditSheetName}!${col}:${col}`;
        const op = type === 'count' ? 'COUNTIFS' : 'SUMIFS';
        const sumRange = type === 'sum' ? `${range}, ` : '';
        return `${op}(${sumRange}${range}, ">=${item.min}", ${range}, "<=${item.max}")`;
    }, 1);

    // Prepare data (Product)
    const productTypes = [...new Set(processedData.map(d => d['Loan Type'] || 'Unmapped'))];
    const productData = productTypes.map(type => {
        const filtered = processedData.filter(d => (d['Loan Type'] || 'Unmapped') === type);
        return { label: type, count: filtered.length, sum: filtered.reduce((acc, d) => acc + parseAmount(d['Net O/s']), 0) };
    });
    addStratTable('II. BY PRODUCT GUIDE', ['Product', 'Accounts', 'Outstanding'], productData, (item, type) => {
        const pCol = `$${rLoanTypeCol}`;
        const osCol = `$${rNetOsCol}`;
        const pRange = `${auditSheetName}!${pCol}:${pCol}`;
        const osRange = `${auditSheetName}!${osCol}:${osCol}`;
        if (type === 'count') return `COUNTIF(${pRange}, "${item.label}")`;
        return `SUMIF(${pRange}, "${item.label}", ${osRange})`;
    }, 5);

    // Prepare data (ROI)
    const roiRanges = [
        { label: 'Upto 8%', min: -99, max: 0.08 },
        { label: '8% to 10%', min: 0.080001, max: 0.10 },
        { label: '10% to 12%', min: 0.100001, max: 0.12 },
        { label: 'Above 12%', min: 0.120001, max: 99 }
    ];
    roiRanges.forEach(r => {
        const filtered = processedData.filter(d => {
            const roi = typeof d['ROI'] === 'number' ? d['ROI'] : parseAmount(d['ROI']);
            return roi >= r.min && roi <= r.max;
        });
        r.count = filtered.length;
        r.sum = filtered.reduce((acc, d) => acc + parseAmount(d['Net O/s']), 0);
    });
    addStratTable('III. BY INTEREST RATE', ['ROI Range', 'Accounts', 'Outstanding'], roiRanges, (item, type) => {
        const rCol = `$${rRoiCol}`;
        const osCol = `$${rNetOsCol}`;
        const rRange = `${auditSheetName}!${rCol}:${rCol}`;
        const osRange = `${auditSheetName}!${osCol}:${osCol}`;
        const op = type === 'count' ? 'COUNTIFS' : 'SUMIFS';
        const sumRange = type === 'sum' ? `${osRange}, ` : '';
        return `${op}(${sumRange}${rRange}, ">=${item.min}", ${rRange}, "<=${item.max}")`;
    }, 9);

    // Prepare data (IRAC)
    const iracData = ['Standard', 'NPA'].map(cls => {
        const filtered = processedData.filter(d => (d['Asset Classification'] || 'Standard') === cls);
        return { label: cls, count: filtered.length, sum: filtered.reduce((acc, d) => acc + parseAmount(d['Net O/s']), 0) };
    });
    addStratTable('IV. BY IRAC CLASSIFICATION', ['Classification', 'Accounts', 'Outstanding'], iracData, (item, type) => {
        const iCol = `$${rAssetClassCol}`;
        const osCol = `$${rNetOsCol}`;
        const iRange = `${auditSheetName}!${iCol}:${iCol}`;
        const osRange = `${auditSheetName}!${osCol}:${osCol}`;
        if (type === 'count') return `COUNTIF(${iRange}, "${item.label}")`;
        return `SUMIF(${iRange}, "${item.label}", ${osRange})`;
    }, 13);

    // Prepare data (EMI Overdue)
    const emiRanges = [
        { label: '0 EMIs', min: -999, max: 0 },
        { label: '1 EMI', min: 1, max: 1 },
        { label: '2 EMIs', min: 2, max: 2 },
        { label: '3 EMIs', min: 3, max: 3 },
        { label: 'Above 3 EMIs', min: 4, max: 9999 }
    ];
    emiRanges.forEach(r => {
        const filtered = processedData.filter(d => (d['Outstanding EMIs'] || 0) >= r.min && (d['Outstanding EMIs'] || 0) <= r.max);
        r.count = filtered.length;
        r.sum = filtered.reduce((acc, d) => acc + parseAmount(d['Net O/s']), 0);
    });
    addStratTable('V. BY EMIS OVERDUE', ['Status', 'Accounts', 'Outstanding'], emiRanges, (item, type) => {
        const eCol = `$${rEmiCol}`;
        const osCol = `$${rNetOsCol}`;
        const eRange = `${auditSheetName}!${eCol}:${eCol}`;
        const osRange = `${auditSheetName}!${osCol}:${osCol}`;
        const op = type === 'count' ? 'COUNTIFS' : 'SUMIFS';
        const sumRange = type === 'sum' ? `${osRange}, ` : '';
        return `${op}(${sumRange}${eRange}, ">=${item.min}", ${eRange}, "<=${item.max}")`;
    }, 17);

    // Helper for auto-fitting columns
    const autoFitColumns = (ws, skipCols = []) => {
        try {
            ws.columns.forEach((column, i) => {
                const colIdx = i + 1;
                if (skipCols.includes(colIdx)) return;

                let maxLength = 0;
                column.eachCell({ includeEmpty: true }, (cell) => {
                    let cellValue = "";
                    if (cell.value && typeof cell.value === 'object' && cell.value.result !== undefined) {
                        cellValue = String(cell.value.result || "");
                    } else if (cell.value && typeof cell.value === 'object' && cell.value.formula) {
                        cellValue = ""; 
                    } else {
                        cellValue = cell.value ? String(cell.value) : "";
                    }

                    if (cell.numFmt && typeof cell.value === 'number') {
                        cellValue = cell.value.toLocaleString('en-IN', { maximumFractionDigits: 0 });
                    }

                    // Override if Advance Stratification set column to 2 manually
                    if (column.width === 2) { maxLength = 0; }
                    else if (cellValue.length > maxLength) {
                        maxLength = cellValue.length;
                    }
                });
                column.width = Math.min(60, Math.max(12, maxLength + 4));
            });
        } catch (e) {
            console.warn('Auto-fit failed for sheet', ws.name, e);
        }
    };

    // Apply auto-fit to all sheets
    if (stratificationSheet) {
        autoFitColumns(stratificationSheet, [4, 8, 12, 16]);
        [4, 8, 12, 16].forEach(c => {
            const col = stratificationSheet.getColumn(c);
            if (col) col.width = 4;
        });
    }

    if (targetFields.includes('CIF')) {
        const cifSheet = workbook.getWorksheet('Potential NPA - CIF Wise');
        if (cifSheet) autoFitColumns(cifSheet);
    }

    // Auto-fit main sheets by name
    const auditSheetRef = workbook.getWorksheet('Audit Report');
    if (auditSheetRef) autoFitColumns(auditSheetRef);
    const npaSheetRef = workbook.getWorksheet('NPA Accounts');
    if (npaSheetRef) autoFitColumns(npaSheetRef);
    const tlSheetRef = workbook.getWorksheet('TL Accounts');
    if (tlSheetRef) autoFitColumns(tlSheetRef);
    const ccSheetRef = workbook.getWorksheet('CC Accounts');
    if (ccSheetRef) autoFitColumns(ccSheetRef);

    // Final buffer writes...

    // --- Generate CIF Wise Summary Sheet ---
    if (targetFields.includes('CIF')) {
        const cifSummarySheet = workbook.addWorksheet('Potential NPA - CIF Wise');

        // Group by CIF
        const cifGroups = {};
        processedData.forEach(row => {
            const cifValue = row['CIF'];
            const cif = cifValue ? String(cifValue).trim() : '';
            if (!cif || cif === 'NA') return;

            if (!cifGroups[cif]) {
                cifGroups[cif] = {
                    cif: cif,
                    name: row['Borrower Name'] || '',
                    totalAccounts: 0,
                    npaAccounts: 0
                };
            }
            cifGroups[cif].totalAccounts += 1;
            if (row['Asset Classification'] === 'NPA') {
                cifGroups[cif].npaAccounts += 1;
            }
        });

        const cifSummaryData = Object.values(cifGroups).map(group => ({
            ...group,
            difference: group.npaAccounts > 0 ? (group.totalAccounts - group.npaAccounts) : ''
        }));

        // Define Columns
        cifSummarySheet.columns = [
            { header: 'CIF', key: 'cif', width: 20 },
            { header: 'Name of the Borrower', key: 'name', width: 40 },
            { header: 'Count of No. of Total Account per CIF', key: 'totalAccounts', width: 30 },
            { header: 'Count of No. of NPA Account per CIF', key: 'npaAccounts', width: 30 },
            { header: 'Difference', key: 'difference', width: 20 }
        ];

        cifSummarySheet.addRows(cifSummaryData);

        // Get column letters from Audit Report for formulas (with fallbacks)
        const auditCifCol = getSafeAuditCol('CIF');
        const auditAssetClassCol = getSafeAuditCol('Asset Classification');

        // Styling and Formulas
        const cHeaderRow = cifSummarySheet.getRow(1);
        cHeaderRow.height = 35;
        cHeaderRow.eachCell((cell) => {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE6E6FA' } };
            cell.font = { bold: true, color: { argb: 'FF483D8B' }, size: 11 };
            cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
            cell.border = {
                top: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                left: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                right: { style: 'thin', color: { argb: 'FFCCCCCC' } }
            };
        });

        cifSummarySheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return;

            const rowData = cifSummaryData[rowNumber - 2];
            
            // Formula for Total Accounts (Col C)
            const totalCell = row.getCell(3);
            totalCell.value = { 
                formula: `COUNTIF('Audit Report'!$${auditCifCol}:$${auditCifCol}, $A${rowNumber})`, 
                result: rowData.totalAccounts || 0
            };

            // Formula for NPA Accounts (Col D)
            const npaCell = row.getCell(4);
            npaCell.value = { 
                formula: `COUNTIFS('Audit Report'!$${auditCifCol}:$${auditCifCol}, $A${rowNumber}, 'Audit Report'!$${auditAssetClassCol}:$${auditAssetClassCol}, "NPA")`, 
                result: rowData.npaAccounts || 0
            };

            const totalRef = `C${rowNumber}`;
            const npaRef = `D${rowNumber}`;
            const diffCell = row.getCell(5);
            const fallbackDiff = rowData.difference;
            diffCell.value = {
                formula: `IF(${npaRef}>0, ${totalRef}-${npaRef}, "")`,
                result: fallbackDiff
            };

            const npaCount = rowData.npaAccounts;
            const hasDifference = fallbackDiff > 0;
            // Potential NPA: when at least one account is NPA but not all are
            const isPotentialNpa = npaCount > 0 && hasDifference;

            row.eachCell((cell, colNumber) => {
                cell.border = {
                    top: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                    left: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                    bottom: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                    right: { style: 'thin', color: { argb: 'FFEEEEEE' } }
                };
                cell.alignment = { vertical: 'middle', horizontal: colNumber <= 2 ? 'left' : 'center' };

                if (isPotentialNpa) {
                    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFEBEE' } }; // Light red
                    cell.font = { color: { argb: 'FFB71C1C' }, bold: true };
                }
            });

            row.getCell(1).numFmt = '@'; // CIF as text
        });
    }

    // --- Generate NPA Accounts Sheet ---
    if (targetFields.includes('Asset Classification')) {
        // Net off Net O/s and Credit Balance for NPA data, then sort by Net O/s descending
        const npaData = processedData
            .filter(row => row['Asset Classification'] === 'NPA')
            .map(row => ({
                ...row,
                'Net O/s': parseAmount(row['Net O/s']) - parseAmount(row['Credit Balance'])
            }))
            .sort((a, b) => b['Net O/s'] - a['Net O/s']);

        const npaExcludes = ['Overdrawn (Yes/No)', 'PAN', 'Entity Type', 'Aadhar No.',
            'Credit Balance', 'A/c Open Date', 'Asset Classification', 'PS Flag'];
        const npaFields = targetFields.filter(f =>
            !npaExcludes.includes(f) &&
            (f !== 'NPA Provision' || npaConfig?.available)
        );
        const styleSheet = (sheetName, data) => {
            if (data.length === 0) return;

            const sheet = workbook.addWorksheet(sheetName);
            sheet.columns = npaFields.map(field => {
                let width = 20;
                if (field === 'Borrower Name') width = 35;
                if (field === 'CIF' || field === 'A/c No.' || field === 'PAN' || field === 'Aadhar No.') width = 18;
                if (field.includes('Date')) width = 15;
                return { header: field, key: field, width };
            });

            sheet.addRows(data);

            const sHeaderRow = sheet.getRow(1);
            sHeaderRow.height = 30;
            sHeaderRow.eachCell((cell) => {
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFE6E6FA' }
                };
                cell.font = { bold: true, color: { argb: 'FF483D8B' }, size: 12 };
                cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                cell.border = {
                    top: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                    left: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                    bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                    right: { style: 'thin', color: { argb: 'FFCCCCCC' } }
                };
            });

            sheet.eachRow((row, rowNumber) => {
                if (rowNumber === 1) return;

                row.eachCell((cell, colNumber) => {
                    const fieldName = npaFields[colNumber - 1];

                    cell.alignment = { vertical: 'middle', horizontal: 'left' };
                    cell.border = {
                        top: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                        left: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                        bottom: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                        right: { style: 'thin', color: { argb: 'FFEEEEEE' } }
                    };

                    if (['Sanction Limit', 'Net O/s', 'Credit Balance', 'Drawing Power', 'Primary Security', 'Collateral Security', 'Overdue Amount', 'NPA Provision'].includes(fieldName)) {
                        cell.numFmt = '#,##0';
                        cell.alignment = { vertical: 'middle', horizontal: 'right' };
                    } else if (fieldName === 'ROI') {
                        cell.numFmt = '0.00%';
                        cell.alignment = { vertical: 'middle', horizontal: 'right' };
                    } else if (['Sanction Date', 'Limit Expiry Date', 'A/c Open Date', 'NPA Date'].includes(fieldName)) {
                        cell.numFmt = 'DD-MM-YY';
                        cell.alignment = { vertical: 'middle', horizontal: 'center' };
                    } else if (['CIF', 'A/c No.', 'PAN', 'Aadhar No.'].includes(fieldName)) {
                        cell.numFmt = '@';
                    } else if (['Overdrawn (Yes/No)', 'Sanction Type', 'PS Flag', 'Loan Type', 'Asset Classification'].includes(fieldName)) {
                        cell.alignment = { vertical: 'middle', horizontal: 'center' };
                    }
                });
            });
        };

        styleSheet('NPA Accounts', npaData);
            
        let npaCriticalData = npaData.filter(row => String(row['Risk Label'] || '').includes('NPA - Critical'));
        if (npaCriticalData.length > 0) {
            styleSheet('Critical NPA Accounts', npaCriticalData);
        }
    }

    // --- Generate Overdrawn Accounts (>110%) Sheet ---
    if (targetFields.includes('Asset Classification') && targetFields.includes('Net O/s') && targetFields.includes('Sanction Limit')) {
        const overdrawnData = [];

        processedData.forEach(row => {
            const assetClass = row['Asset Classification'];
            if (assetClass !== 'Standard') return;

            const netOs = parseAmount(row['Net O/s']);
            const sanctionLimit = parseAmount(row['Sanction Limit']);

            if (sanctionLimit > 0 && netOs > (sanctionLimit * 1.1)) {
                const overdrawnPercent = netOs / sanctionLimit;
                overdrawnData.push({
                    ...row,
                    'Overdrawn %': overdrawnPercent
                });
            }
        });

        if (overdrawnData.length > 0) {
            // Net off Net O/s - Credit Balance, then sort by nettedOs descending
            const nettedOverdrawnData = overdrawnData.map(row => ({
                ...row,
                'Net O/s': parseAmount(row['Net O/s']) - parseAmount(row['Credit Balance'])
            }));
            nettedOverdrawnData.sort((a, b) => b['Net O/s'] - a['Net O/s']);

            const overdrawnExcludes = ['Overdrawn (Yes/No)', 'Credit Balance', 'PAN', 'Entity Type', 'Aadhar No.',
                'Asset Classification', 'NPA Date', 'A/c Open Date', 'PS Flag', 'NPA Provision'];
            const overdrawnFields = [...nonNpaTargetFields.filter(f => !overdrawnExcludes.includes(f)), 'Overdrawn %'];

            const sheet = workbook.addWorksheet('Overdrawn (>110%)');
            sheet.columns = overdrawnFields.map(field => {
                let width = 20;
                if (field === 'Borrower Name') width = 35;
                if (field === 'CIF' || field === 'A/c No.' || field === 'PAN' || field === 'Aadhar No.') width = 18;
                if (field.includes('Date')) width = 15;
                if (field === 'Overdrawn %') width = 15;
                return { header: field, key: field, width };
            });

            sheet.addRows(nettedOverdrawnData);

            const sHeaderRow = sheet.getRow(1);
            sHeaderRow.height = 30;
            sHeaderRow.eachCell((cell) => {
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFE6E6FA' }
                };
                cell.font = { bold: true, color: { argb: 'FF483D8B' }, size: 12 };
                cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                cell.border = {
                    top: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                    left: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                    bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                    right: { style: 'thin', color: { argb: 'FFCCCCCC' } }
                };
            });

            sheet.eachRow((row, rowNumber) => {
                if (rowNumber === 1) return;

                const fallbackData = nettedOverdrawnData[rowNumber - 2];

                // Add calculation formula for Overdrawn % if columns are present
                const netOsColIdx = overdrawnFields.indexOf('Net O/s') + 1;
                const sanctionColIdx = overdrawnFields.indexOf('Sanction Limit') + 1;
                const overdrawnPercentColIdx = overdrawnFields.indexOf('Overdrawn %') + 1;

                if (netOsColIdx > 0 && sanctionColIdx > 0 && overdrawnPercentColIdx > 0) {
                    const netOsRef = `${getColumnLetter(netOsColIdx)}${rowNumber}`;
                    const sanctionRef = `${getColumnLetter(sanctionColIdx)}${rowNumber}`;
                    // Formula: Net O/s / Sanction Limit
                    row.getCell(overdrawnPercentColIdx).value = {
                        formula: `IF(${sanctionRef}>0, ${netOsRef}/${sanctionRef}, 0)`,
                        result: fallbackData['Overdrawn %']
                    };
                }

                row.eachCell((cell, colNumber) => {
                    const fieldName = overdrawnFields[colNumber - 1];

                    cell.alignment = { vertical: 'middle', horizontal: 'left' };
                    cell.border = {
                        top: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                        left: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                        bottom: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                        right: { style: 'thin', color: { argb: 'FFEEEEEE' } }
                    };

                    if (['Sanction Limit', 'Net O/s', 'Credit Balance', 'Drawing Power', 'Primary Security', 'Collateral Security', 'Overdue Amount'].includes(fieldName)) {
                        cell.numFmt = '#,##0';
                        cell.alignment = { vertical: 'middle', horizontal: 'right' };
                    } else if (fieldName === 'ROI' || fieldName === 'Overdrawn %') {
                        cell.numFmt = '0.00%';
                        cell.alignment = { vertical: 'middle', horizontal: 'right' };

                        // Highlight high overdrawn percentages
                        if (fieldName === 'Overdrawn %' && fallbackData['Overdrawn %'] > 1.1) {
                            cell.fill = {
                                type: 'pattern',
                                pattern: 'solid',
                                fgColor: { argb: 'FFFFEBEE' } // Light red/pink
                            };
                            cell.font = {
                                color: { argb: 'FFD32F2F' }, // Dark red text
                                bold: true
                            };
                        }
                    } else if (['Sanction Date', 'Limit Expiry Date', 'A/c Open Date', 'NPA Date'].includes(fieldName)) {
                        cell.numFmt = 'DD-MM-YY';
                        cell.alignment = { vertical: 'middle', horizontal: 'center' };
                    } else if (['CIF', 'A/c No.', 'PAN', 'Aadhar No.'].includes(fieldName)) {
                        cell.numFmt = '@';
                    } else if (['Overdrawn (Yes/No)', 'Sanction Type', 'PS Flag', 'Loan Type', 'Asset Classification'].includes(fieldName)) {
                        cell.alignment = { vertical: 'middle', horizontal: 'center' };
                    }
                });
            });
        }
    }

    // --- Generate Gold Accounts Sheet ---
    // Check both Loan Type and Facility Code for gold
    if (targetFields.includes('Loan Type') || targetFields.includes('Facility Code')) {
        let goldData = processedData.filter(row => {
            const loanType = String(row['Loan Type'] || '').toLowerCase();
            const facilityCode = String(row['Facility Code'] || '').toLowerCase();
            return loanType.includes('gold') || facilityCode.includes('gl') || facilityCode === 'gold'
                || facilityCode.includes('gold loan');
        });

        if (goldData.length > 0) {
            if (targetFields.includes('Net O/s')) {
                // Sort by Net O/s descending
                goldData.sort((a, b) => {
                    const osA = parseAmount(a['Net O/s']);
                    const osB = parseAmount(b['Net O/s']);
                    return osB - osA;
                });
            }

            // Net off Net O/s and Credit Balance for Gold data
            goldData = goldData.map(row => ({
                ...row,
                'Net O/s': parseAmount(row['Net O/s']) - parseAmount(row['Credit Balance'])
            }));

            const goldExcludes = ['PAN', 'Entity Type', 'Aadhar No.',
                'Credit Balance', 'A/c Open Date', 'Facility Code', 'PS Flag', 'NPA Date', 'NPA Provision'];
            const goldFields = nonNpaTargetFields.filter(f => !goldExcludes.includes(f));

            const sheet = workbook.addWorksheet('Gold Accounts');
            sheet.columns = goldFields.map(field => {
                let width = 20;
                if (field === 'Borrower Name') width = 35;
                if (field === 'CIF' || field === 'A/c No.' || field === 'PAN' || field === 'Aadhar No.') width = 18;
                if (field.includes('Date')) width = 15;
                return { header: field, key: field, width };
            });

            sheet.addRows(goldData);

            const sHeaderRow = sheet.getRow(1);
            sHeaderRow.height = 30;
            sHeaderRow.eachCell((cell) => {
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFE6E6FA' }
                };
                cell.font = { bold: true, color: { argb: 'FF483D8B' }, size: 12 };
                cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                cell.border = {
                    top: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                    left: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                    bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                    right: { style: 'thin', color: { argb: 'FFCCCCCC' } }
                };
            });

            sheet.eachRow((row, rowNumber) => {
                if (rowNumber === 1) return;

                row.eachCell((cell, colNumber) => {
                    const fieldName = goldFields[colNumber - 1];

                    cell.alignment = { vertical: 'middle', horizontal: 'left' };
                    cell.border = {
                        top: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                        left: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                        bottom: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                        right: { style: 'thin', color: { argb: 'FFEEEEEE' } }
                    };

                    if (['Sanction Limit', 'Net O/s', 'Credit Balance', 'Drawing Power', 'Primary Security', 'Collateral Security', 'Overdue Amount'].includes(fieldName)) {
                        cell.numFmt = '#,##0';
                        cell.alignment = { vertical: 'middle', horizontal: 'right' };
                    } else if (fieldName === 'ROI') {
                        cell.numFmt = '0.00%';
                        cell.alignment = { vertical: 'middle', horizontal: 'right' };
                    } else if (['Sanction Date', 'Limit Expiry Date', 'A/c Open Date', 'NPA Date'].includes(fieldName)) {
                        cell.numFmt = 'DD-MM-YY';
                        cell.alignment = { vertical: 'middle', horizontal: 'center' };
                    } else if (['CIF', 'A/c No.', 'PAN', 'Aadhar No.'].includes(fieldName)) {
                        cell.numFmt = '@';
                    } else if (['Overdrawn (Yes/No)', 'Sanction Type', 'PS Flag', 'Loan Type', 'Asset Classification', 'Asset Class (Theo Bal)'].includes(fieldName)) {
                        cell.alignment = { vertical: 'middle', horizontal: 'center' };
                    }
                });
            });
        }
    }

    // --- Generate Security Shortfall Sheet ---
    if (targetFields.includes('Sanction Limit') && targetFields.includes('Primary Security')) {
        const shortfallData = [];

        processedData.forEach(row => {
            const sanctionLimit = parseAmount(row['Sanction Limit']);
            const primarySecurity = parseAmount(row['Primary Security']);

            // Only consider if both are valid numerical data and limit > security
            if (sanctionLimit > 0 && primarySecurity >= 0 && sanctionLimit > primarySecurity) {
                const shortfallAmount = sanctionLimit - primarySecurity;
                const shortfallPercent = shortfallAmount / sanctionLimit;
                shortfallData.push({
                    ...row,
                    'Shortfall Amount': shortfallAmount,
                    'Shortfall %': shortfallPercent
                });
            }
        });

        if (shortfallData.length > 0) {
            // Sort by Shortfall Amount descending
            shortfallData.sort((a, b) => b['Shortfall Amount'] - a['Shortfall Amount']);

            // Net off Net O/s and Credit Balance for Security Shortfall data
            const nettedShortfallData = shortfallData.map(row => ({
                ...row,
                'Net O/s': parseAmount(row['Net O/s']) - parseAmount(row['Credit Balance'])
            }));
            const shortfallExcludes = ['PAN', 'Entity Type', 'Aadhar No.', 'Sanction Type',
                'Credit Balance', 'A/c Open Date', 'PS Flag', 'NPA Provision'];
            const shortfallFields = [...nonNpaTargetFields.filter(f => !shortfallExcludes.includes(f)), 'Shortfall Amount', 'Shortfall %'];

            const sheet = workbook.addWorksheet('Security Shortfall');
            sheet.columns = shortfallFields.map(field => {
                let width = 20;
                if (field === 'Borrower Name') width = 35;
                if (field === 'CIF' || field === 'A/c No.' || field === 'PAN' || field === 'Aadhar No.') width = 18;
                if (field.includes('Date')) width = 15;
                if (field === 'Shortfall %') width = 15;
                return { header: field, key: field, width };
            });

            sheet.addRows(nettedShortfallData);

            const sHeaderRow = sheet.getRow(1);
            sHeaderRow.height = 30;
            sHeaderRow.eachCell((cell) => {
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFE6E6FA' }
                };
                cell.font = { bold: true, color: { argb: 'FF483D8B' }, size: 12 };
                cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                cell.border = {
                    top: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                    left: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                    bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                    right: { style: 'thin', color: { argb: 'FFCCCCCC' } }
                };
            });

            sheet.eachRow((row, rowNumber) => {
                if (rowNumber === 1) return;

                const fallbackData = nettedShortfallData[rowNumber - 2];

                // Add calculation formula for Shortfall % if columns are present
                const sanctionColIdx = shortfallFields.indexOf('Sanction Limit') + 1;
                const primarySecColIdx = shortfallFields.indexOf('Primary Security') + 1;
                const shortfallAmountColIdx = shortfallFields.indexOf('Shortfall Amount') + 1;
                const shortfallPercentColIdx = shortfallFields.indexOf('Shortfall %') + 1;

                if (sanctionColIdx > 0 && primarySecColIdx > 0 && shortfallAmountColIdx > 0 && shortfallPercentColIdx > 0) {
                    const sanctionRef = `${getColumnLetter(sanctionColIdx)}${rowNumber}`;
                    const primarySecRef = `${getColumnLetter(primarySecColIdx)}${rowNumber}`;
                    const shortfallAmountRef = `${getColumnLetter(shortfallAmountColIdx)}${rowNumber}`;

                    // Formula: Sanction Limit - Primary Security
                    row.getCell(shortfallAmountColIdx).value = {
                        formula: `MAX(0, ${sanctionRef}-${primarySecRef})`,
                        result: fallbackData['Shortfall Amount']
                    };

                    // Formula: Shortfall Amount / Sanction Limit
                    row.getCell(shortfallPercentColIdx).value = {
                        formula: `IF(${sanctionRef}>0, ${shortfallAmountRef}/${sanctionRef}, 0)`,
                        result: fallbackData['Shortfall %']
                    };
                }

                row.eachCell((cell, colNumber) => {
                    const fieldName = shortfallFields[colNumber - 1];

                    cell.alignment = { vertical: 'middle', horizontal: 'left' };
                    cell.border = {
                        top: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                        left: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                        bottom: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                        right: { style: 'thin', color: { argb: 'FFEEEEEE' } }
                    };

                    if (['Sanction Limit', 'Net O/s', 'Credit Balance', 'Drawing Power', 'Primary Security', 'Collateral Security', 'Overdue Amount', 'Shortfall Amount'].includes(fieldName)) {
                        cell.numFmt = '#,##0';
                        cell.alignment = { vertical: 'middle', horizontal: 'right' };
                    } else if (fieldName === 'ROI' || fieldName === 'Shortfall %') {
                        cell.numFmt = '0.00%';
                        cell.alignment = { vertical: 'middle', horizontal: 'right' };

                        // Highlight high shortfall percentages > 50%
                        if (fieldName === 'Shortfall %' && fallbackData['Shortfall %'] > 0.5) {
                            cell.fill = {
                                type: 'pattern',
                                pattern: 'solid',
                                fgColor: { argb: 'FFFFEBEE' } // Light red/pink
                            };
                            cell.font = {
                                color: { argb: 'FFD32F2F' }, // Dark red text
                                bold: true
                            };
                        }
                    } else if (['Sanction Date', 'Limit Expiry Date', 'A/c Open Date', 'NPA Date'].includes(fieldName)) {
                        cell.numFmt = 'DD-MM-YY';
                        cell.alignment = { vertical: 'middle', horizontal: 'center' };
                    } else if (['CIF', 'A/c No.', 'PAN', 'Aadhar No.'].includes(fieldName)) {
                        cell.numFmt = '@';
                    } else if (['Overdrawn (Yes/No)', 'Sanction Type', 'PS Flag', 'Loan Type', 'Asset Classification'].includes(fieldName)) {
                        cell.alignment = { vertical: 'middle', horizontal: 'center' };
                    }
                });
            });
        }
    }

    // --- Generate Large Advances & Top 60% Sheets ---
    if (targetFields.includes('CIF') && targetFields.includes('Net O/s')) {
        let totalAdvances = 0;
        const cifGroups = {};

        // 1. Group by CIF and calculate exposures
        processedData.forEach(row => {
            const cif = String(row['CIF'] || '').trim();
            if (!cif) return;

            const netOs = parseAmount(row['Net O/s']);
            // Include Credit Balance if available for Net exposure calculation
            const creditBalance = targetFields.includes('Credit Balance') ? parseAmount(row['Credit Balance']) : 0;
            const advance = netOs - creditBalance;

            totalAdvances += advance;

            if (!cifGroups[cif]) {
                cifGroups[cif] = {
                    borrowerName: row['Borrower Name'] || '',
                    totalLimit: 0,
                    totalBalance: 0,
                    accounts: []
                };
            }
            cifGroups[cif].totalLimit += parseAmount(row['Sanction Limit']);
            cifGroups[cif].totalBalance += advance;
            cifGroups[cif].accounts.push({
                ...row,
                _calculatedAdvance: advance // Store for later if we want account-level details
            });
        });

        if (totalAdvances > 0) {
            // Calculate Thresholds
            const tenPercentAdvances = totalAdvances * 0.10;
            const minimumAmount = 100000000; // 10,00,00,000
            const reportingThreshold = Math.min(tenPercentAdvances, minimumAmount);

            // Filter Large Advances
            const largeAdvancesCIFs = Object.keys(cifGroups).filter(cif => cifGroups[cif].totalBalance > reportingThreshold);

            if (largeAdvancesCIFs.length > 0) {
                const laSheet = workbook.addWorksheet('Large Advances');
                laSheet.columns = [
                    { header: 'CIF', key: 'cif', width: 20 },
                    { header: 'Customer Name', key: 'name', width: 40 },
                    { header: 'Limit', key: 'limit', width: 25 },
                    { header: 'Balance', key: 'balance', width: 25 }
                ];

                // Add Summary Header specific to Large Advances
                laSheet.addRow({ cif: 'Total Advances', name: totalAdvances });
                laSheet.addRow([]); // spacer
                laSheet.addRow({ cif: '10% of Advances', name: tenPercentAdvances });
                laSheet.addRow({ cif: 'Minimum Amount', name: minimumAmount });
                laSheet.addRow([]); // spacer
                laSheet.addRow({ cif: 'Lower of the Above', name: reportingThreshold });
                laSheet.addRow([]); // spacer
                laSheet.addRow([]); // spacer for table header later

                // Style the summary
                laSheet.getCell('A1').font = { bold: true };
                laSheet.getCell('A3').font = { bold: true };
                laSheet.getCell('A4').font = { bold: true };
                laSheet.getCell('A6').font = { bold: true, size: 12 };
                laSheet.getCell('B6').font = { bold: true, size: 12 };

                [1, 3, 4, 6].forEach(rIdx => {
                    const row = laSheet.getRow(rIdx);
                    row.getCell(2).numFmt = '#,##0';
                    row.getCell(2).alignment = { horizontal: 'right' };
                });

                // Add Table Header
                const tableHeaderRowIdx = 8;
                const tableHeaderString = laSheet.getRow(tableHeaderRowIdx);
                tableHeaderString.values = ['Accounts Required to be Reported'];
                laSheet.mergeCells(`A${tableHeaderRowIdx}:D${tableHeaderRowIdx}`);
                tableHeaderString.getCell(1).alignment = { horizontal: 'center' };
                tableHeaderString.getCell(1).font = { bold: true, size: 12 };
                tableHeaderString.getCell(1).border = {
                    top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }
                };

                const colHeaderRow = laSheet.getRow(tableHeaderRowIdx + 1);
                colHeaderRow.values = ['CIF', 'Customer Name', 'Limit', 'Balance'];
                colHeaderRow.font = { bold: true };
                for (let i = 1; i <= 4; i++) {
                    colHeaderRow.getCell(i).border = {
                        top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }
                    };
                }

                // Append Data
                largeAdvancesCIFs.forEach(cif => {
                    const data = cifGroups[cif];
                    const row = laSheet.addRow({
                        cif: cif,
                        name: data.borrowerName,
                        limit: data.totalLimit,
                        balance: data.totalBalance
                    });

                    row.getCell(3).numFmt = '#,##0';
                    row.getCell(4).numFmt = '#,##0';

                    for (let i = 1; i <= 4; i++) {
                        row.getCell(i).border = {
                            top: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                            left: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                            bottom: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                            right: { style: 'thin', color: { argb: 'FFEEEEEE' } }
                        };
                    }
                });
            }

            // --- Generate Top New Advances Sheet (Top 5 per Category, New only, capped at 50%) ---
            // Step 1: Filter to 'New' advances only
            const newAccounts = [];
            processedData.forEach(row => {
                const sanctionType = String(row['Sanction Type'] || '').toUpperCase();
                if (sanctionType !== 'NEW') return;

                const netOs = parseAmount(row['Net O/s']);
                const creditBalance = parseAmount(row['Credit Balance']);
                const advance = netOs - creditBalance;
                if (advance !== 0) {
                    newAccounts.push({
                        ...row,
                        '_calculatedAdvance': advance,
                        '_pctOfTotal': totalAdvances > 0 ? (advance / totalAdvances) : 0
                    });
                }
            });

            // Step 2: Group by Loan Type (category)
            const categoryMap = {};
            newAccounts.forEach(account => {
                const cat = String(account['Loan Type'] || account['Sanction Type'] || 'General').trim();
                if (!categoryMap[cat]) categoryMap[cat] = [];
                categoryMap[cat].push(account);
            });

            // Step 3: Pick top 5 per category (sorted by advance descending)
            const top60Accounts = [];
            Object.keys(categoryMap).forEach(cat => {
                const sorted = categoryMap[cat].sort((a, b) => b['_calculatedAdvance'] - a['_calculatedAdvance']);
                const picked = sorted.slice(0, 5); // Take top 5 or all if fewer
                picked.forEach(a => top60Accounts.push(a));
            });

            // Step 4: Sort the combined result and cap at 50% of total advances
            top60Accounts.sort((a, b) => b['_calculatedAdvance'] - a['_calculatedAdvance']);
            const targetExposure = totalAdvances * 0.50;
            const minExposure = totalAdvances * 0.40;
            let runningExposure = 0;
            const finalTop60 = [];
            for (const account of top60Accounts) {
                finalTop60.push(account);
                runningExposure += account['_calculatedAdvance'];
                if (runningExposure >= targetExposure) break;
            }

            // Step 5: Ensure minimum 40% coverage by adding largest remaining accounts
            if (runningExposure < minExposure) {
                const selectedAcctNos = new Set(finalTop60.map(a => a['A/c No.']));
                const remainingAccounts = processedData
                    .filter(a => !selectedAcctNos.has(a['A/c No.']))
                    .map(a => {
                        const advance = parseAmount(a['Net O/s']) - parseAmount(a['Credit Balance']);
                        return {
                            ...a,
                            '_calculatedAdvance': advance,
                            '_pctOfTotal': totalAdvances > 0 ? (advance / totalAdvances) : 0
                        };
                    })
                    .filter(a => a['_calculatedAdvance'] > 0)
                    .sort((a, b) => b['_calculatedAdvance'] - a['_calculatedAdvance']);

                for (const account of remainingAccounts) {
                    finalTop60.push(account);
                    runningExposure += account['_calculatedAdvance'];
                    if (runningExposure >= minExposure) break;
                }
            }

            if (finalTop60.length > 0) {
                const top60Excludes = ['PAN', 'Entity Type', 'Aadhar No.',
                    'Credit Balance', 'Sanction Type', 'PS Flag', 'NPA Date'];
                const top60Fields = [...nonNpaTargetFields.filter(f => !top60Excludes.includes(f)), '% of Total Adv'];
                const top60Sheet = workbook.addWorksheet('Top Advances');

                top60Sheet.columns = top60Fields.map(field => {
                    let width = 20;
                    if (field === 'Borrower Name') width = 35;
                    if (field === 'CIF' || field === 'A/c No.' || field === 'PAN' || field === 'Aadhar No.') width = 18;
                    if (field.includes('Date')) width = 15;
                    if (field === '% of Total Adv') width = 15;
                    return { header: field, key: field, width };
                });

                finalTop60.forEach(account => {
                    // Net off Net O/s and Credit Balance for Top Advances output
                    const nettedOs = parseAmount(account['Net O/s']) - parseAmount(account['Credit Balance']);
                    const rowData = { ...account, 'Net O/s': nettedOs, '% of Total Adv': account['_pctOfTotal'] };
                    top60Sheet.addRow(rowData);
                });

                const sHeaderRow = top60Sheet.getRow(1);
                sHeaderRow.height = 30;
                sHeaderRow.eachCell((cell) => {
                    cell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FFE6E6FA' }
                    };
                    cell.font = { bold: true, color: { argb: 'FF483D8B' }, size: 12 };
                    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                    cell.border = {
                        top: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                        left: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                        bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                        right: { style: 'thin', color: { argb: 'FFCCCCCC' } }
                    };
                });

                top60Sheet.eachRow((row, rowNumber) => {
                    if (rowNumber === 1) return;

                    row.eachCell((cell, colNumber) => {
                        const fieldName = top60Fields[colNumber - 1];

                        cell.alignment = { vertical: 'middle', horizontal: 'left' };
                        cell.border = {
                            top: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                            left: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                            bottom: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                            right: { style: 'thin', color: { argb: 'FFEEEEEE' } }
                        };

                        if (['Sanction Limit', 'Net O/s', 'Credit Balance', 'Drawing Power', 'Primary Security', 'Collateral Security', 'Overdue Amount'].includes(fieldName)) {
                            cell.numFmt = '#,##0';
                            cell.alignment = { vertical: 'middle', horizontal: 'right' };
                        } else if (fieldName === 'ROI' || fieldName === '% of Total Adv') {
                            cell.numFmt = '0.00%';
                            cell.alignment = { vertical: 'middle', horizontal: 'right' };
                        } else if (['Sanction Date', 'Limit Expiry Date', 'A/c Open Date', 'NPA Date'].includes(fieldName)) {
                            cell.numFmt = 'DD-MM-YY';
                            cell.alignment = { vertical: 'middle', horizontal: 'center' };
                        } else if (['CIF', 'A/c No.', 'PAN', 'Aadhar No.'].includes(fieldName)) {
                            cell.numFmt = '@';
                        } else if (['Overdrawn (Yes/No)', 'Sanction Type', 'PS Flag', 'Loan Type', 'Asset Classification'].includes(fieldName)) {
                            cell.alignment = { vertical: 'middle', horizontal: 'center' };
                        }
                    });
                });
            }
        }
    }

    // --- Generate NPA with Credit Balance Sheet ---
    if (targetFields.includes('Asset Classification') && targetFields.includes('Credit Balance')) {
        const npaCbData = processedData
            .filter(row => {
                const isNPA = row['Asset Classification'] === 'NPA';
                const netOs = parseAmount(row['Net O/s']);
                const creditBalance = parseAmount(row['Credit Balance']);
                const nettedOs = netOs - creditBalance;
                return isNPA && nettedOs < 0; // Only accounts where credit balance exceeds outstanding
            })
            .map(row => ({
                ...row,
                'Net O/s': parseAmount(row['Net O/s']) - parseAmount(row['Credit Balance'])
            }));

        if (npaCbData.length > 0) {
            // Sort by Net O/s (netted) ascending (largest negative amount first)
            npaCbData.sort((a, b) => a['Net O/s'] - b['Net O/s']);

            const npaCbExcludes = ['PAN', 'Entity Type', 'Aadhar No.', 'Sanction Type', 'Overdrawn (Yes/No)', 'Asset Classification', 'Credit Balance'];
            const npaCbFields = nonNpaTargetFields.filter(f => !npaCbExcludes.includes(f));

            const npaCbSheet = workbook.addWorksheet('NPA with Credit Balance');
            npaCbSheet.columns = npaCbFields.map(field => {
                let width = 20;
                if (field === 'Borrower Name') width = 35;
                if (field === 'CIF' || field === 'A/c No.' || field === 'PAN' || field === 'Aadhar No.') width = 18;
                if (field.includes('Date')) width = 15;
                return { header: field, key: field, width };
            });

            npaCbSheet.addRows(npaCbData);

            const sHeaderRow = npaCbSheet.getRow(1);
            sHeaderRow.height = 30;
            sHeaderRow.eachCell((cell) => {
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFE6E6FA' }
                };
                cell.font = { bold: true, color: { argb: 'FF483D8B' }, size: 12 };
                cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                cell.border = {
                    top: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                    left: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                    bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                    right: { style: 'thin', color: { argb: 'FFCCCCCC' } }
                };
            });

            npaCbSheet.eachRow((row, rowNumber) => {
                if (rowNumber === 1) return;

                row.eachCell((cell, colNumber) => {
                    const fieldName = npaCbFields[colNumber - 1];

                    cell.alignment = { vertical: 'middle', horizontal: 'left' };
                    cell.border = {
                        top: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                        left: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                        bottom: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                        right: { style: 'thin', color: { argb: 'FFEEEEEE' } }
                    };

                    if (['Sanction Limit', 'Net O/s', 'Credit Balance', 'Drawing Power', 'Primary Security', 'Collateral Security', 'Overdue Amount'].includes(fieldName)) {
                        cell.numFmt = '#,##0';
                        cell.alignment = { vertical: 'middle', horizontal: 'right' };
                    } else if (fieldName === 'ROI') {
                        cell.numFmt = '0.00%';
                        cell.alignment = { vertical: 'middle', horizontal: 'right' };
                    } else if (['Sanction Date', 'Limit Expiry Date', 'A/c Open Date', 'NPA Date'].includes(fieldName)) {
                        cell.numFmt = 'DD-MM-YY';
                        cell.alignment = { vertical: 'middle', horizontal: 'center' };
                    } else if (['CIF', 'A/c No.', 'PAN', 'Aadhar No.'].includes(fieldName)) {
                        cell.numFmt = '@';
                    } else if (['Overdrawn (Yes/No)', 'Sanction Type', 'PS Flag', 'Loan Type', 'Asset Classification'].includes(fieldName)) {
                        cell.alignment = { vertical: 'middle', horizontal: 'center' };
                    }
                });
            });
        }
    }

    // --- Generate NPA Provision Calculator Sheet ---
    if (targetFields.includes('Asset Classification') && targetFields.includes('NPA Date') && targetFields.includes('Net O/s')) {
        const npaProvisionData = [];

        processedData.forEach(row => {
            if (row['Asset Classification'] === 'NPA') {
                const npaDate = parseDateString(row['NPA Date']);
                let npaDays = 0;

                if (npaDate && auditToObj) {
                    const diffTime = Math.abs(auditToObj - npaDate);
                    npaDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
                }

                const netOs = parseAmount(row['Net O/s']);
                const creditBalance = targetFields.includes('Credit Balance') ? parseAmount(row['Credit Balance']) : 0;
                const advance = netOs - creditBalance; // True exposure (can be negative)
                const provisioningBase = Math.max(0, advance); // Use 0 if negative for calculations

                const primarySecurity = targetFields.includes('Primary Security') ? parseAmount(row['Primary Security']) : 0;

                // If Primary Security exists, limit the secured portion mathematically to the true exposure
                const securedPortion = Math.min(provisioningBase, primarySecurity);
                const unsecuredPortion = Math.max(0, provisioningBase - securedPortion);

                // Calculate Provision Percentages
                let securedProvPct = 0;
                if (npaDays <= 365) securedProvPct = 0.15;
                else if (npaDays <= 730) securedProvPct = 0.25;
                else if (npaDays <= 1460) securedProvPct = 0.40;
                else securedProvPct = 1.00;

                let unsecuredProvPct = 0;
                if (npaDays <= 365) unsecuredProvPct = 0.25;
                else unsecuredProvPct = 1.00;

                const provisionOnSecured = securedPortion * securedProvPct;
                const provisionOnUnsecured = unsecuredPortion * unsecuredProvPct;
                const totalProvision = provisionOnSecured + provisionOnUnsecured;

                npaProvisionData.push({
                    ...row,
                    'Net O/s': advance, // Use netted value
                    'NPA Days': npaDays,
                    'Secured Portion': securedPortion,
                    'Unsecured Portion': unsecuredPortion,
                    'Secured Prov %': securedProvPct,
                    'Unsecured Prov %': unsecuredProvPct,
                    'Provision Amount': totalProvision,
                    'Actual Provision': row['NPA Provision'] || 0
                });
            }
        });

        if (npaProvisionData.length > 0) {
            // Sort by Net O/s (netted) descending
            npaProvisionData.sort((a, b) => b['Net O/s'] - a['Net O/s']);

            const provExcludes = ['Overdrawn (Yes/No)', 'PAN', 'Entity Type', 'Aadhar No.',
                'Credit Balance', 'A/c Open Date', 'PS Flag'];
            const calcFields = [
                ...targetFields.filter(f => !provExcludes.includes(f) && f !== 'NPA Provision'),
                'NPA Days', 'Secured Portion', 'Unsecured Portion',
                'Secured Prov %', 'Unsecured Prov %', 'Provision Amount'
            ];

            if (npaConfig?.available) {
                calcFields.push('Actual Provision', 'Difference');
            }

            const provSheet = workbook.addWorksheet('NPA Provision Calculator');
            provSheet.columns = calcFields.map(field => {
                let width = 20;
                if (field === 'Borrower Name') width = 35;
                if (field === 'CIF' || field === 'A/c No.' || field === 'PAN' || field === 'Aadhar No.') width = 18;
                if (field.includes('Date')) width = 15;
                if (field.includes('%') || field === 'NPA Days') width = 15;
                return { header: field, key: field, width };
            });

            provSheet.addRows(npaProvisionData);

            const sHeaderRow = provSheet.getRow(1);
            sHeaderRow.height = 30;
            sHeaderRow.eachCell((cell) => {
                let bgColor = 'FFE6E6FA';
                if (['Secured Prov %', 'Unsecured Prov %', 'Provision Amount'].includes(cell.value)) {
                    bgColor = 'FFFFE4E1'; // Light red for result columns
                } else if (['NPA Days', 'Secured Portion', 'Unsecured Portion'].includes(cell.value)) {
                    bgColor = 'FFFFFACD'; // Light yellow for intermediate calc columns
                }

                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: bgColor }
                };
                cell.font = { bold: true, color: { argb: 'FF483D8B' }, size: 12 };
                cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                cell.border = {
                    top: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                    left: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                    bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                    right: { style: 'thin', color: { argb: 'FFCCCCCC' } }
                };
            });

            provSheet.eachRow((row, rowNumber) => {
                if (rowNumber === 1) return;

                const fallbackData = npaProvisionData[rowNumber - 2];

                const getRef = (colName) => {
                    const idx = calcFields.indexOf(colName);
                    return idx >= 0 ? `${getColumnLetter(idx + 1)}${rowNumber}` : null;
                };

                const npaDateRef = getRef('NPA Date');
                const netOsRef = getRef('Net O/s');
                const cbRef = getRef('Credit Balance');
                const psRef = getRef('Primary Security');

                const npaDaysRef = getRef('NPA Days');
                const secPortionRef = getRef('Secured Portion');
                const unsecPortionRef = getRef('Unsecured Portion');
                const secProvPctRef = getRef('Secured Prov %');
                const unsecProvPctRef = getRef('Unsecured Prov %');

                let exposureFormula = netOsRef;
                if (cbRef) {
                    exposureFormula = `MAX(0, ${netOsRef}-${cbRef})`;
                } else {
                    exposureFormula = `MAX(0, ${netOsRef})`;
                }

                if (secPortionRef && unsecPortionRef && secProvPctRef && unsecProvPctRef) {

                    if (npaDateRef && auditToObj) {
                        const y = auditToObj.getFullYear();
                        const m = auditToObj.getMonth() + 1;
                        const d = auditToObj.getDate();
                        const npaDaysColIdx = calcFields.indexOf('NPA Days') + 1;
                        if (npaDaysColIdx > 0) {
                            row.getCell(npaDaysColIdx).value = {
                                formula: `MAX(0, DATE(${y},${m},${d})-${npaDateRef})`,
                                result: fallbackData['NPA Days']
                            };
                        }
                    }

                    const secColIdx = calcFields.indexOf('Secured Portion') + 1;
                    if (secColIdx > 0) {
                        if (psRef) {
                            row.getCell(secColIdx).value = {
                                formula: `MIN(${exposureFormula}, ${psRef})`,
                                result: fallbackData['Secured Portion']
                            };
                        } else {
                            row.getCell(secColIdx).value = 0;
                        }
                    }

                    const unsecColIdx = calcFields.indexOf('Unsecured Portion') + 1;
                    if (unsecColIdx > 0) {
                        row.getCell(unsecColIdx).value = {
                            formula: `MAX(0, ${exposureFormula}-${secPortionRef})`,
                            result: fallbackData['Unsecured Portion']
                        };
                    }

                    const secPctColIdx = calcFields.indexOf('Secured Prov %') + 1;
                    if (secPctColIdx > 0) {
                        row.getCell(secPctColIdx).value = {
                            formula: `IF(${npaDaysRef}<=365, 0.15, IF(${npaDaysRef}<=730, 0.25, IF(${npaDaysRef}<=1460, 0.40, 1.0)))`,
                            result: fallbackData['Secured Prov %']
                        };
                    }

                    const unsecPctColIdx = calcFields.indexOf('Unsecured Prov %') + 1;
                    if (unsecPctColIdx > 0) {
                        row.getCell(unsecPctColIdx).value = {
                            formula: `IF(${npaDaysRef}<=365, 0.25, 1.0)`,
                            result: fallbackData['Unsecured Prov %']
                        };
                    }

                    const provAmtColIdx = calcFields.indexOf('Provision Amount') + 1;
                    if (provAmtColIdx > 0) {
                        row.getCell(provAmtColIdx).value = {
                            formula: `ROUND((${secPortionRef}*${secProvPctRef})+(${unsecPortionRef}*${unsecProvPctRef}), 2)`,
                            result: fallbackData['Provision Amount']
                        };
                    }

                    const actualProvColIdx = calcFields.indexOf('Actual Provision') + 1;
                    const diffColIdx = calcFields.indexOf('Difference') + 1;
                    if (actualProvColIdx > 0 && diffColIdx > 0) {
                        const actualProvRef = `${getColumnLetter(actualProvColIdx)}${rowNumber}`;
                        const calcProvRef = `${getColumnLetter(provAmtColIdx)}${rowNumber}`;
                        row.getCell(diffColIdx).value = {
                            formula: `${calcProvRef}-${actualProvRef}`,
                            result: (fallbackData['Provision Amount'] || 0) - (fallbackData['Actual Provision'] || 0)
                        };
                    }
                }

                row.eachCell((cell, colNumber) => {
                    const fieldName = calcFields[colNumber - 1];

                    cell.alignment = { vertical: 'middle', horizontal: 'left' };
                    cell.border = {
                        top: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                        left: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                        bottom: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                        right: { style: 'thin', color: { argb: 'FFEEEEEE' } }
                    };

                    if (['Sanction Limit', 'Net O/s', 'Credit Balance', 'Drawing Power', 'Primary Security', 'Collateral Security', 'Overdue Amount', 'Secured Portion', 'Unsecured Portion', 'Provision Amount', 'Actual Provision', 'Difference'].includes(fieldName)) {
                        cell.numFmt = '#,##0';
                        cell.alignment = { vertical: 'middle', horizontal: 'right' };

                        if (fieldName === 'Provision Amount') {
                            cell.font = { bold: true, color: { argb: 'FFB22222' } }; // Dark red
                        }

                        if (fieldName === 'Difference' && Math.abs(cell.value?.result || cell.value) > 100) {
                            cell.font = { color: { argb: 'FFFF0000' }, bold: true };
                        }
                    } else if (fieldName === 'ROI' || fieldName.includes('Prov %')) {
                        cell.numFmt = '0.00%';
                        cell.alignment = { vertical: 'middle', horizontal: 'right' };
                    } else if (['Sanction Date', 'Limit Expiry Date', 'A/c Open Date', 'NPA Date'].includes(fieldName)) {
                        cell.numFmt = 'DD-MM-YY';
                        cell.alignment = { vertical: 'middle', horizontal: 'center' };
                    } else if (['CIF', 'A/c No.', 'PAN', 'Aadhar No.'].includes(fieldName)) {
                        cell.numFmt = '@';
                    } else if (['Overdrawn (Yes/No)', 'Sanction Type', 'PS Flag', 'Loan Type', 'Asset Classification', 'NPA Days'].includes(fieldName)) {
                        cell.alignment = { vertical: 'middle', horizontal: 'center' };
                    }
                });
            });
        }
    }

    // --- Generate Quick Mortality Accounts Sheet ---
    {
        const qmData = processedData
            .filter(row => {
                const npaDate = parseDateString(row['NPA Date']);
                const sanctionDate = parseDateString(row['Sanction Date']);
                if (!npaDate || !sanctionDate) return false;
                const diffDays = Math.round((npaDate - sanctionDate) / (1000 * 60 * 60 * 24));
                return diffDays > 0 && diffDays < 365;
            })
            .map(row => ({
                ...row,
                'Net O/s': parseAmount(row['Net O/s']) - parseAmount(row['Credit Balance'])
            }))
            .sort((a, b) => b['Net O/s'] - a['Net O/s']);

        if (qmData.length > 0) {
            const qmSheet = workbook.addWorksheet('Quick Mortality');

            // Add a 'Days to NPA' column to the output
            const qmExcludes = ['PAN', 'Entity Type', 'Aadhar No.', 'Credit Balance', 'PS Flag', 'Overdrawn (Yes/No)', 'Sanction Type', 'A/c Open Date', 'Facility Code', 'Drawing Power', 'Asset Classification', 'Overdue Amount'];
            const qmFields = [...nonNpaTargetFields.filter(f => !qmExcludes.includes(f)), 'Days to NPA'];

            qmSheet.columns = qmFields.map(field => {
                let width = 20;
                if (field === 'Borrower Name') width = 35;
                if (field === 'CIF' || field === 'A/c No.') width = 18;
                if (field.includes('Date')) width = 15;
                if (field === 'Days to NPA') width = 14;
                return { header: field, key: field, width };
            });

            qmData.forEach(row => {
                const npaDate = parseDateString(row['NPA Date']);
                const sanctionDate = parseDateString(row['Sanction Date']);
                const daysToNpa = (npaDate && sanctionDate)
                    ? Math.round((npaDate - sanctionDate) / (1000 * 60 * 60 * 24))
                    : null;
                // Net O/s is already netted in qmData, so just use as-is
                qmSheet.addRow({ ...row, 'Days to NPA': daysToNpa });
            });

            const hRow = qmSheet.getRow(1);
            hRow.height = 30;
            hRow.eachCell((cell) => {
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFFFDAB9' } // Pastel peach/orange
                };
                cell.font = { bold: true, color: { argb: 'FF8B4513' }, size: 12 }; // Saddle brown
                cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                cell.border = {
                    top: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                    left: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                    bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                    right: { style: 'thin', color: { argb: 'FFCCCCCC' } }
                };
            });

            qmSheet.eachRow((row, rowNumber) => {
                if (rowNumber === 1) return;
                row.eachCell((cell, colNumber) => {
                    const fieldName = qmFields[colNumber - 1];
                    cell.alignment = { vertical: 'middle', horizontal: 'left' };
                    cell.border = {
                        top: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                        left: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                        bottom: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                        right: { style: 'thin', color: { argb: 'FFEEEEEE' } }
                    };
                    if (['Sanction Limit', 'Net O/s', 'Credit Balance', 'Drawing Power', 'Primary Security', 'Collateral Security', 'Overdue Amount'].includes(fieldName)) {
                        cell.numFmt = '#,##0';
                        cell.alignment = { vertical: 'middle', horizontal: 'right' };
                    } else if (fieldName === 'ROI') {
                        cell.numFmt = '0.00%';
                        cell.alignment = { vertical: 'middle', horizontal: 'right' };
                    } else if (['Sanction Date', 'Limit Expiry Date', 'A/c Open Date', 'NPA Date'].includes(fieldName)) {
                        cell.numFmt = 'DD-MM-YY';
                        cell.alignment = { vertical: 'middle', horizontal: 'center' };
                    } else if (fieldName === 'Days to NPA') {
                        cell.numFmt = '0';
                        cell.alignment = { vertical: 'middle', horizontal: 'right' };
                        cell.font = { bold: true, color: { argb: 'FF8B0000' } }; // Dark red for emphasis
                    } else if (['Overdrawn (Yes/No)', 'Sanction Type', 'PS Flag', 'Loan Type', 'Asset Classification'].includes(fieldName)) {
                        cell.alignment = { vertical: 'middle', horizontal: 'center' };
                    }
                });
            });
        }
    }

    // --- Generate SMA Accounts Sheet ---
    {
        const smaData = processedData
            .filter(row => {
                const ac = String(row['Asset Class'] || '').toUpperCase();
                return ac.includes('SMA-0') || ac.includes('SMA-1') || ac.includes('SMA-2') ||
                    ac.includes('SMA 0') || ac.includes('SMA 1') || ac.includes('SMA 2') ||
                    ac.includes('SMA0') || ac.includes('SMA1') || ac.includes('SMA2');
            })
            .map(row => ({
                ...row,
                'Net O/s': parseAmount(row['Net O/s']) - parseAmount(row['Credit Balance'])
            }))
            .sort((a, b) => b['Net O/s'] - a['Net O/s']);

        if (smaData.length > 0) {
            const smaSheet = workbook.addWorksheet('SMA Accounts');

            const smaExcludedFields = ['Credit Balance', 'PAN', 'Entity Type', 'Aadhar No.', 'PS Flag', 'Overdrawn (Yes/No)', 'Sanction Type', 'A/c Open Date', 'Facility Code', 'Drawing Power', 'Asset Classification', 'NPA Date'];
            const smaColumnFields = nonNpaTargetFields.filter(f => !smaExcludedFields.includes(f));
            smaSheet.columns = smaColumnFields.map(field => {
                let width = 20;
                if (field === 'Borrower Name') width = 35;
                if (field === 'CIF' || field === 'A/c No.' || field === 'PAN' || field === 'Aadhar No.') width = 18;
                if (field.includes('Date')) width = 15;
                return { header: field, key: field, width };
            });

            smaSheet.addRows(smaData);

            const hRow = smaSheet.getRow(1);
            hRow.height = 30;
            hRow.eachCell((cell) => {
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFFFFACD' } // Light goldenrod yellow to distinguish from standard data
                };
                cell.font = { bold: true, color: { argb: 'FF8B8000' }, size: 12 }; // Dark yellow/olive text
                cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                cell.border = {
                    top: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                    left: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                    bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                    right: { style: 'thin', color: { argb: 'FFCCCCCC' } }
                };
            });

            smaSheet.eachRow((row, rowNumber) => {
                if (rowNumber === 1) return;

                row.eachCell((cell, colNumber) => {
                    const fieldName = smaColumnFields[colNumber - 1];

                    cell.alignment = { vertical: 'middle', horizontal: 'left' };
                    cell.border = {
                        top: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                        left: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                        bottom: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                        right: { style: 'thin', color: { argb: 'FFEEEEEE' } }
                    };

                    if (['Sanction Limit', 'Net O/s', 'Credit Balance', 'Drawing Power', 'Primary Security', 'Collateral Security', 'Overdue Amount'].includes(fieldName)) {
                        cell.numFmt = '#,##0';
                        cell.alignment = { vertical: 'middle', horizontal: 'right' };
                    } else if (fieldName === 'ROI') {
                        cell.numFmt = '0.00%';
                        cell.alignment = { vertical: 'middle', horizontal: 'right' };
                    } else if (['Sanction Date', 'Limit Expiry Date', 'A/c Open Date', 'NPA Date'].includes(fieldName)) {
                        cell.numFmt = 'DD-MM-YY';
                        cell.alignment = { vertical: 'middle', horizontal: 'center' };
                    } else if (['CIF', 'A/c No.', 'PAN', 'Aadhar No.'].includes(fieldName)) {
                        cell.numFmt = '@';
                    } else if (['Overdrawn (Yes/No)', 'Sanction Type', 'PS Flag', 'Loan Type', 'Asset Classification'].includes(fieldName)) {
                        cell.alignment = { vertical: 'middle', horizontal: 'center' };
                    }
                });
            });
        }
    }

    // --- Generate Expired Limits Sheet ---
    if (auditToObj) {
        const expiredData = processedData
            .filter(row => {
                const expDate = parseDateString(row['Limit Expiry Date']);
                if (!expDate) return false;
                // Must be Standard Asset
                const ac = String(row['Asset Classification'] || '').toUpperCase();
                const isStandard = ac.includes('STANDARD') || ac === 'STD';
                return expDate < auditToObj && isStandard;
            })
            .map(row => ({
                ...row,
                'Net O/s': parseAmount(row['Net O/s']) - parseAmount(row['Credit Balance'])
            }))
            .sort((a, b) => b['Net O/s'] - a['Net O/s']);

        if (expiredData.length > 0) {
            const expiredSheet = workbook.addWorksheet('Expired Limits');

            const expiredExcludedFields = ['Credit Balance', 'PAN', 'Entity Type', 'Aadhar No.', 'PS Flag', 'Overdrawn (Yes/No)', 'Sanction Type', 'A/c Open Date', 'Facility Code', 'Drawing Power', 'Asset Classification', 'NPA Date'];
            const expiredColumnFields = nonNpaTargetFields.filter(f => !expiredExcludedFields.includes(f));
            expiredSheet.columns = expiredColumnFields.map(field => {
                let width = 20;
                if (field === 'Borrower Name') width = 35;
                if (field === 'CIF' || field === 'A/c No.' || field === 'PAN' || field === 'Aadhar No.') width = 18;
                if (field.includes('Date')) width = 15;
                return { header: field, key: field, width };
            });

            expiredSheet.addRows(expiredData);

            const hRow = expiredSheet.getRow(1);
            hRow.height = 30;
            hRow.eachCell((cell) => {
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFFFE4E1' } // Lightest pastel red/pink (MistyRose)
                };
                cell.font = { bold: true, color: { argb: 'FF8B0000' }, size: 12 }; // Dark red text
                cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                cell.border = {
                    top: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                    left: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                    bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } },
                    right: { style: 'thin', color: { argb: 'FFCCCCCC' } }
                };
            });

            expiredSheet.eachRow((row, rowNumber) => {
                if (rowNumber === 1) return;

                row.eachCell((cell, colNumber) => {
                    const fieldName = expiredColumnFields[colNumber - 1];

                    cell.alignment = { vertical: 'middle', horizontal: 'left' };
                    cell.border = {
                        top: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                        left: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                        bottom: { style: 'thin', color: { argb: 'FFEEEEEE' } },
                        right: { style: 'thin', color: { argb: 'FFEEEEEE' } }
                    };

                    if (['Sanction Limit', 'Net O/s', 'Credit Balance', 'Drawing Power', 'Primary Security', 'Collateral Security', 'Overdue Amount'].includes(fieldName)) {
                        cell.numFmt = '#,##0';
                        cell.alignment = { vertical: 'middle', horizontal: 'right' };
                    } else if (fieldName === 'ROI') {
                        cell.numFmt = '0.00%';
                        cell.alignment = { vertical: 'middle', horizontal: 'right' };
                    } else if (['Sanction Date', 'Limit Expiry Date', 'A/c Open Date', 'NPA Date'].includes(fieldName)) {
                        cell.numFmt = 'DD-MM-YY';
                        cell.alignment = { vertical: 'middle', horizontal: 'center' };
                    } else if (['CIF', 'A/c No.', 'PAN', 'Aadhar No.'].includes(fieldName)) {
                        cell.numFmt = '@';
                    } else if (['Overdrawn (Yes/No)', 'Sanction Type', 'PS Flag', 'Loan Type', 'Asset Classification'].includes(fieldName)) {
                        cell.alignment = { vertical: 'middle', horizontal: 'center' };
                    }
                });
            });
        }
    }

    // --- Generate Sheet Index in Parameters ---
    const indexParamSheet = workbook.getWorksheet('Parameters');
    if (indexParamSheet) {
        let nextRow = indexParamSheet.rowCount + 2;
        workbook.worksheets.forEach(sheet => {
            if (sheet.name !== 'Parameters') {
                const linkCell = indexParamSheet.getCell(`A${nextRow}`);
                linkCell.value = {
                    text: `Go to ${sheet.name}`,
                    hyperlink: `#'${sheet.name}'!A1`
                };
                linkCell.font = {
                    color: { argb: 'FF0000FF' },
                    underline: true,
                    size: 11
                };
                linkCell.alignment = { vertical: 'middle', horizontal: 'left' };
                nextRow++;
            }
        });
    }

    // Apply Auto-Filter and Column Auto-Width to ALL generated sheets
    workbook.worksheets.forEach(sheet => {
        // Skip 'Parameters' and 'Large Advances' (has custom layout without proper headers on row 1 in a simple way)
        if (sheet.name === 'Parameters' || sheet.name === 'Large Advances') {
            return;
        }

        const rowCount = sheet.rowCount;
        const colCount = sheet.columnCount;

        if (rowCount > 0 && colCount > 0) {
            // Add AutoFilter
            sheet.autoFilter = {
                from: { row: 1, column: 1 },
                to: { row: rowCount, column: colCount }
            };
        }

        // Auto-adjust column width based on content (up to max width map of 40)
        sheet.columns.forEach(column => {
            let maxLength = 10; // Minimum width
            column.eachCell({ includeEmpty: false }, cell => {
                let cellLength = cell.value ? cell.value.toString().length : 0;

                // For number fields that have comma formatting, adjust length estimate
                if (typeof cell.value === 'number') {
                    cellLength = cell.value.toLocaleString('en-IN').length + 2;
                } else if (cell.value instanceof Date) {
                    // Dates formatted as DD-MM-YY take up around 8 chars. Give it 10 to be safe
                    cellLength = 10;
                }

                if (cellLength > maxLength) {
                    maxLength = cellLength;
                }
            });

            // Set width, with max of 40 and buffer
            column.width = Math.min(40, maxLength + 2);
        });
    });

    // Export buffer
    
        // Apply Universal Formatting across all sheets
        workbook.worksheets.forEach(ws => {
            ws.columns.forEach((col, i) => {
                const headerCell = ws.getCell(1, i + 1);
                if (headerCell && (headerCell.value === 'A/c No.' || headerCell.value === 'CIF' || headerCell.value === 'PAN')) {
                    col.eachCell({ includeEmpty: false }, cell => { if (cell.row > 1) cell.numFmt = '@'; }); // Force text format
                }
            });
            
            ws.eachRow(row => {
                row.eachCell({ includeEmpty: true }, cell => {
                    let cellVal = cell.value;
                    let isCalc = false;
                    if (cellVal && typeof cellVal === 'object' && cellVal.result !== undefined) {
                        cellVal = cellVal.result; // Extract calculated result
                        isCalc = true;
                    }
                    
                    if (typeof cellVal === 'number' && !isNaN(cellVal)) {
                        // Numeric types (either direct or calculated)
                        if (!cell.numFmt || cell.numFmt === 'General' || String(cell.numFmt).includes('#,##0')) {
                            // Ignore specific percentages if they exist
                            if (!String(cell.numFmt).includes('%') && !String(cell.numFmt).includes('0%')) {
                                cell.numFmt = '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)';
                            }
                        }
                    } else if (cellVal instanceof Date) {
                        if (!cell.numFmt || cell.numFmt === 'General') {
                            cell.numFmt = 'DD-MM-YY';
                        }
                    } else if (typeof cellVal === 'string' && cellVal.trim() !== '') {
                        // Apply string safety on columns
                    }
                });
            });
        });

        const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    // Include audit date in filename for easy identification
    const dateStr = auditToObj.toLocaleDateString('en-IN', { day: '2-digit', month: '2-digit', year: 'numeric' }).replace(/\//g, '-');
    saveAs(blob, `Audit_Report_${dateStr}.xlsx`);
};


