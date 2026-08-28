/**
 * Incoming Operations & Supplier Directory Handlers
 */

function getIncomingSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return getSheetByNames_(ss, [SHEET_NAMES.INCOMING, SHEET_NAMES.LEGACY_INCOMING, 'Incoming DB', 'IncDB']);
}

function fetchSupplierDatabase() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const supplierSet = new Set();

    const dbSheet = ss.getSheetByName(SHEET_NAMES.DATABASE) || ss.getSheetByName('database') || ss.getSheetByName('DB');
    if (dbSheet) {
      const data = dbSheet.getDataRange().getValues();
      if (data.length > 1) {
        const headers = data[0].map(h => String(h).trim().toLowerCase());
        let supplierColIdx = headers.findIndex(h => h.includes('supplier'));
        if (supplierColIdx === -1) supplierColIdx = 4;
        for (let i = 1; i < data.length; i++) {
          const val = String(data[i][supplierColIdx] || '').trim();
          if (val && val !== '-' && val.toLowerCase() !== 'n/a') supplierSet.add(val);
        }
      }
    }

    const incSheet = getIncomingSheet_();
    if (incSheet) {
      const incData = incSheet.getDataRange().getValues();
      for (let i = 1; i < incData.length; i++) {
        const sup = String(incData[i][2] || '').trim();
        if (sup && sup !== '-' && sup.toLowerCase() !== 'n/a') supplierSet.add(sup);
      }
    }

    return { supplierList: Array.from(supplierSet).sort() };
  } catch (err) {
    return { supplierList: [], error: err.message };
  }
}

function fetchIncomingData() {
  try {
    const sheet = getIncomingSheet_();
    if (!sheet) return { totals: {}, monthlyIncoming: {}, transactions: [], pastRecords: [], error: 'Incoming sheet not found. Expected Incoming or IncDB.' };

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      const emptyTotals = {};
      MATERIAL_COLUMNS.forEach(mat => emptyTotals[mat.key] = 0);
      return { totals: emptyTotals, monthlyIncoming: {}, transactions: [], pastRecords: [] };
    }

    const headers = data[0].map(h => String(h || '').trim().toLowerCase());
    const materialColumnMap = {};
    MATERIAL_COLUMNS.forEach(mat => {
      let idx = headers.findIndex(h => h === mat.key.toLowerCase().trim());
      if (idx === -1) idx = headers.findIndex(h => h === mat.name.toLowerCase().trim());
      materialColumnMap[mat.key] = idx;
    });

    const totals = {};
    MATERIAL_COLUMNS.forEach(mat => totals[mat.key] = 0);
    const monthlyIncoming = {};
    const transactions = [];
    const pastRecords = [];

    const idCol = headers.findIndex(h => h === 'id');
    const dateCol = headers.findIndex(h => h === 'date');
    const supplierCol = headers.findIndex(h => h === 'supplier' || h === 'supplier name' || h === 'party');
    const drCol = headers.findIndex(h => h === 'dr #' || h === 'dr no' || h === 'dr number' || h === 'dr' || h.includes('delivery receipt'));

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rawDate = dateCol !== -1 ? row[dateCol] : row[1];
      if (!rawDate) continue;
      const dateObj = rawDate instanceof Date ? rawDate : new Date(rawDate);
      if (isNaN(dateObj.getTime())) continue;

      const tz = Session.getScriptTimeZone() || 'GMT+8';
      const formattedDate = Utilities.formatDate(dateObj, tz, 'M/d/yyyy');
      const monthKey = Utilities.formatDate(dateObj, tz, 'MMM yy');
      if (!monthlyIncoming[monthKey]) monthlyIncoming[monthKey] = {};

      const supplierName = supplierCol !== -1 ? String(row[supplierCol] || '-').trim() : String(row[2] || '-').trim();
      let drNum = drCol !== -1 ? String(row[drCol] || '-').trim() : String(row[row.length - 1] || '-').trim();
      if (!drNum) drNum = '-';

      const recordItems = {};
      MATERIAL_COLUMNS.forEach(mat => {
        const idx = materialColumnMap[mat.key];
        const val = idx !== -1 ? Number(row[idx]) || 0 : 0;
        recordItems[mat.key] = val;
        totals[mat.key] += val;
        monthlyIncoming[monthKey][mat.key] = (monthlyIncoming[monthKey][mat.key] || 0) + val;
      });

      transactions.push({ type: 'Incoming', date: formattedDate, rawDate: dateObj.getTime(), rowIndex: i + 1, party: supplierName, ref: drNum !== '-' ? drNum : 'N/A' });
      pastRecords.push({ id: i + 1, recordId: idCol !== -1 && row[idCol] ? row[idCol] : ('INC-' + (i + 1)), date: formattedDate, supplier: supplierName, drNumber: drNum !== '-' ? drNum : '', items: recordItems });
    }

    pastRecords.reverse();
    return { totals, monthlyIncoming, transactions, pastRecords };
  } catch (err) {
    console.error('fetchIncomingData error:', err);
    return { totals: {}, monthlyIncoming: {}, transactions: [], pastRecords: [], error: err.message };
  }
}

function recordIncoming(formData) {
  try {
    const sheet = getIncomingSheet_();
    if (!sheet) throw new Error('Incoming sheet not found. Expected Incoming or IncDB.');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h || '').trim());
    const newRow = new Array(headers.length).fill('');
    const newId = 'INC-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT+8', 'yyyyMMdd-HHmmss');

    headers.forEach((header, index) => {
      const h = header.toLowerCase().trim();
      if (h === 'id') newRow[index] = newId;
      else if (h === 'date') newRow[index] = formData.date ? new Date(formData.date) : new Date();
      else if (h === 'supplier' || h === 'supplier name' || h === 'party') newRow[index] = formData.party || '';
      else if (h === 'dr #' || h === 'dr no' || h === 'dr number' || h === 'dr' || h.includes('delivery receipt')) newRow[index] = formData.drNumber || '';
      else {
        const material = MATERIAL_COLUMNS.find(mat => mat.key.toLowerCase().trim() === h || mat.name.toLowerCase().trim() === h);
        if (material) newRow[index] = formData.items && formData.items[material.key] ? Number(formData.items[material.key]) : 0;
      }
    });
    sheet.appendRow(newRow);
    return { success: true, id: newId };
  } catch (err) {
    console.error('recordIncoming error:', err);
    return { success: false, error: err.message };
  }
}

function updateIncoming(rowIndex, formData) {
  try {
    const sheet = getIncomingSheet_();
    if (!sheet) throw new Error('Incoming sheet not found. Expected Incoming or IncDB.');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h || '').trim());
    const existingRow = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
    const updatedRow = existingRow.slice();
    headers.forEach((header, index) => {
      const h = header.toLowerCase().trim();
      if (h === 'date') updatedRow[index] = formData.date ? new Date(formData.date) : new Date();
      else if (h === 'supplier' || h === 'supplier name' || h === 'party') updatedRow[index] = formData.party || '';
      else if (h === 'dr #' || h === 'dr no' || h === 'dr number' || h === 'dr' || h.includes('delivery receipt')) updatedRow[index] = formData.drNumber || '';
      else {
        const material = MATERIAL_COLUMNS.find(mat => mat.key.toLowerCase().trim() === h || mat.name.toLowerCase().trim() === h);
        if (material) updatedRow[index] = formData.items && formData.items[material.key] ? Number(formData.items[material.key]) : 0;
      }
    });
    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([updatedRow]);
    return { success: true };
  } catch (err) {
    console.error('updateIncoming error:', err);
    return { success: false, error: err.message };
  }
}
