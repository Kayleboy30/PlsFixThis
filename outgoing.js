/**
 * ==========================================
 * OUTGOING MODULE (Dispatches / Gate Pass)
 * ==========================================
 */

function getOutgoingSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return getSheetByNames_(ss, [SHEET_NAMES.OUTGOING, SHEET_NAMES.LEGACY_OUTGOING, 'Outgoing DB', 'OutDB']);
}

/**
 * Reads Cluster, Cluster Head, Contact, Base Station, and Complete Address from Database.
 */
function fetchRegionDirectory() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const directory = {};
    const clusterNamesSet = new Set();
    const avpMasterMap = {};

    function normalizeCluster(str) {
      if (!str) return '';
      const s = String(str).trim();
      const numMatch = s.match(/\d+/);
      return numMatch ? ('CLUSTER ' + numMatch[0]) : s.toUpperCase();
    }

    const dbSheet = ss.getSheetByName(SHEET_NAMES.DATABASE) || ss.getSheetByName('database') || ss.getSheetByName('DB');
    if (dbSheet) {
      const data = dbSheet.getDataRange().getValues();
      if (data.length > 1) {
        const headers = data[0].map(h => String(h).trim().toLowerCase());
        let avpColIdx = headers.findIndex(h => h.includes('avp'));
        let divColIdx = headers.findIndex(h => h.includes('division'));
        let opColIdx = headers.findIndex(h => h.includes('operation'));
        let destColIdx = headers.findIndex(h => h.includes('destination') || h.includes('complete address'));
        if (avpColIdx === -1) avpColIdx = 0;
        if (divColIdx === -1) divColIdx = 1;
        if (opColIdx === -1) opColIdx = 2;
        if (destColIdx === -1) destColIdx = 3;

        for (let i = 1; i < data.length; i++) {
          const aName = String(data[i][avpColIdx] || '').trim();
          if (!aName || aName === '-' || aName.toLowerCase() === 'n/a') continue;
          const info = { division: String(data[i][divColIdx] || '').trim(), operation: String(data[i][opColIdx] || '').trim(), destination: String(data[i][destColIdx] || '').trim() };
          avpMasterMap[aName.toLowerCase()] = info;
          avpMasterMap[aName.toLowerCase().replace(/\s+/g, '')] = info;
        }

        let clusterIdx = headers.findIndex(h => (h.includes('cluster') || h.includes('region')) && !h.includes('head') && !h.includes('contact'));
        let headIdx = headers.findIndex(h => h.includes('head') && !h.includes('contact'));
        let divIdx = headers.findIndex(h => h.includes('division'));
        let addrIdx = headers.findIndex(h => h.includes('address') || h.includes('destination'));
        const colXIdx = 23;
        let contactIdx = headers.findIndex(h => h.includes('cluster') && h.includes('contact'));
        if (contactIdx === -1) contactIdx = headers.findIndex((h, i) => h.includes('contact') && i > clusterIdx);
        let branchIdx = headers.findIndex(h => h.includes('branch') || h.includes('station') || h.includes('base'));
        if (branchIdx !== -1 && branchIdx < clusterIdx) branchIdx = headers.findIndex((h, i) => (h.includes('branch') || h.includes('station') || h.includes('base')) && i > clusterIdx);

        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          const rawCluster = clusterIdx !== -1 ? String(row[clusterIdx] || '').trim() : '';
          if (!rawCluster || rawCluster === '-' || rawCluster.toLowerCase() === 'n/a') continue;
          const norm = normalizeCluster(rawCluster);
          clusterNamesSet.add(norm);
          const assignedAvp = row[colXIdx] && String(row[colXIdx]).trim() !== '-' ? String(row[colXIdx]).trim() : '';
          const avpDetails = avpMasterMap[assignedAvp.toLowerCase()] || avpMasterMap[assignedAvp.toLowerCase().replace(/\s+/g, '')] || {};
          directory[norm] = {
            cluster: norm,
            avpName: assignedAvp,
            division: avpDetails.division || '',
            operation: avpDetails.operation || '',
            destination: avpDetails.destination || '',
            clusterHead: headIdx !== -1 ? String(row[headIdx] || '').trim() : '',
            clusterHeadContact: contactIdx !== -1 ? String(row[contactIdx] || '').trim() : '',
            baseStation: branchIdx !== -1 ? String(row[branchIdx] || '').trim() : ''
          };
        }
      }
    }

    const outSheet = getOutgoingSheet_();
    if (outSheet) {
      const outData = outSheet.getDataRange().getValues();
      if (outData.length > 1) {
        const headers = outData[0].map(h => String(h).trim().toLowerCase());
        const clusterIdx = headers.findIndex(h => (h.includes('cluster') || h.includes('region')) && !h.includes('head') && !h.includes('contact'));
        const headIdx = headers.findIndex(h => h.includes('head') && !h.includes('contact'));
        const contactIdx = headers.findIndex(h => h.includes('contact'));
        const branchIdx = headers.findIndex(h => h.includes('branch') || h.includes('station') || h.includes('base'));
        const avpIdx = headers.findIndex(h => h.includes('avp'));
        const divIdx = headers.findIndex(h => h.includes('division'));
        const opIdx = headers.findIndex(h => h.includes('operation'));
        const destIdx = headers.findIndex(h => h.includes('destination'));
        for (let i = outData.length - 1; i >= 1; i--) {
          const row = outData[i];
          const rawCluster = clusterIdx !== -1 ? String(row[clusterIdx] || '').trim() : '';
          if (!rawCluster || rawCluster === '-' || rawCluster.toLowerCase() === 'n/a') continue;
          const norm = normalizeCluster(rawCluster);
          clusterNamesSet.add(norm);
          if (!directory[norm]) {
            const outAvp = avpIdx !== -1 ? String(row[avpIdx] || '').trim() : '';
            const avpDetails = avpMasterMap[outAvp.toLowerCase()] || {};
            directory[norm] = {
              cluster: norm, avpName: outAvp,
              division: avpDetails.division || (divIdx !== -1 ? String(row[divIdx] || '').trim() : ''),
              operation: avpDetails.operation || (opIdx !== -1 ? String(row[opIdx] || '').trim() : ''),
              destination: avpDetails.destination || (destIdx !== -1 ? String(row[destIdx] || '').trim() : ''),
              clusterHead: headIdx !== -1 ? String(row[headIdx] || '').trim() : '',
              clusterHeadContact: contactIdx !== -1 ? String(row[contactIdx] || '').trim() : '',
              baseStation: branchIdx !== -1 ? String(row[branchIdx] || '').trim() : ''
            };
          }
        }
      }
    }

    return { directory, clusterList: Array.from(clusterNamesSet).sort((a, b) => (parseInt(a.replace(/\D/g, '')) || 0) - (parseInt(b.replace(/\D/g, '')) || 0)) };
  } catch (err) {
    return { directory: {}, clusterList: [], error: err.message };
  }
}

function fetchAvpDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.DATABASE) || ss.getSheetByName('database') || ss.getSheetByName('DB');
  const avpDirectory = {};
  const avpList = [];
  if (!sheet || sheet.getLastRow() <= 1) return { avpDirectory, avpList };
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim().toLowerCase());
  const avpIdx = headers.findIndex(h => h.includes('avp'));
  const divIdx = headers.findIndex(h => h.includes('division'));
  const opIdx = headers.findIndex(h => h.includes('operation'));
  const destIdx = headers.findIndex(h => h.includes('destination'));
  const addressIdx = headers.findIndex(h => h.includes('complete address') || h === 'address');
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const avpName = avpIdx !== -1 ? String(row[avpIdx] || '').trim() : '';
    if (!avpName) continue;
    const info = { division: divIdx !== -1 ? String(row[divIdx] || '').trim() : '', operation: opIdx !== -1 ? String(row[opIdx] || '').trim() : '', destination: destIdx !== -1 ? String(row[destIdx] || '').trim() : '', completeAddress: addressIdx !== -1 ? String(row[addressIdx] || '').trim() : '' };
    if (!avpDirectory[avpName]) { avpDirectory[avpName] = info; avpDirectory[avpName.toLowerCase()] = info; avpList.push(avpName); }
  }
  return { avpDirectory, avpList: avpList.sort() };
}

function parseOutgoingTransactionDate_(value) {
  if (!value) return null;
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  const str = String(value).trim();
  const match = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (match) {
    const p1 = Number(match[1]), p2 = Number(match[2]), year = Number(match[3]);
    let month, day;
    if (p1 > 12) { day = p1; month = p2; }
    else if (p2 > 12) { month = p1; day = p2; }
    else { day = p1; month = p2; }
    return new Date(year, month - 1, day);
  }
  const parsed = new Date(str);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function fetchOutgoingData() {
  const outSheet = getOutgoingSheet_();
  const totals = {};
  const monthlyOutgoing = {};
  const transactions = [];
  const pastRecords = [];
  MATERIAL_COLUMNS.forEach(mat => totals[mat.key] = 0);
  if (!outSheet || outSheet.getLastRow() <= 1) return { totals, monthlyOutgoing, transactions, pastRecords };

  const values = outSheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const lower = headers.map(h => h.toLowerCase());
  const dateIdx = lower.findIndex(h => h === 'date');
  const avpIdx = lower.findIndex(h => h === 'avp name' || h === 'avp');
  const divIdx = lower.findIndex(h => h === 'division');
  const destIdx = lower.findIndex(h => h === 'destination');
  const controlIdx = lower.findIndex(h => h.includes('control'));
  const regIdx = lower.findIndex(h => h.includes('region') || (h.includes('cluster') && !h.includes('head')));
  const opIdx = lower.findIndex(h => h === 'operation');
  const headIdx = lower.findIndex(h => h === 'cluster head');
  const contactIdx = lower.findIndex(h => h.includes('head contact') || h.includes('contact'));
  const baseIdx = lower.findIndex(h => h === 'base station' || h.includes('base'));
  const notesIdx = lower.findIndex(h => h === 'notes');

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const transactionDate = dateIdx !== -1 ? parseOutgoingTransactionDate_(row[dateIdx]) : null;
    const rawDate = transactionDate ? transactionDate.getTime() : 0;
    const monthStr = transactionDate ? Utilities.formatDate(transactionDate, Session.getScriptTimeZone() || 'GMT+8', 'MMM yy') : '';
    const avpVal = avpIdx !== -1 ? String(row[avpIdx] || '').trim() : '-';
    const divVal = divIdx !== -1 ? String(row[divIdx] || '').trim() : '-';
    const destVal = destIdx !== -1 ? String(row[destIdx] || '').trim() : '-';
    const controlVal = controlIdx !== -1 ? String(row[controlIdx] || '').trim() : '-';
    const regVal = regIdx !== -1 ? String(row[regIdx] || '').trim() : '-';
    const opVal = opIdx !== -1 ? String(row[opIdx] || '').trim() : '';
    const headVal = headIdx !== -1 ? String(row[headIdx] || '').trim() : '';
    const contactVal = contactIdx !== -1 ? String(row[contactIdx] || '').trim() : '';
    const baseVal = baseIdx !== -1 ? String(row[baseIdx] || '').trim() : '';
    const notesVal = notesIdx !== -1 ? String(row[notesIdx] || '').trim() : '';

    if (transactionDate && !monthlyOutgoing[monthStr]) monthlyOutgoing[monthStr] = {};
    MATERIAL_COLUMNS.forEach(mat => {
      const idx = lower.findIndex(h => h === mat.key.toLowerCase() || h === mat.name.toLowerCase());
      const val = idx !== -1 ? Number(row[idx]) || 0 : 0;
      totals[mat.key] += val;
      if (transactionDate) monthlyOutgoing[monthStr][mat.key] = (monthlyOutgoing[monthStr][mat.key] || 0) + val;
    });

    if (!transactionDate) continue;
    const formattedDate = Utilities.formatDate(transactionDate, Session.getScriptTimeZone() || 'GMT+8', 'M/d/yyyy');
    transactions.push({ type: 'Outgoing', date: formattedDate, rawDate, rowIndex: r + 1, party: destVal !== '-' ? destVal : avpVal, ref: controlVal !== '-' ? controlVal : regVal });
    const items = {};
    MATERIAL_COLUMNS.forEach(mat => {
      const idx = lower.findIndex(h => h === mat.key.toLowerCase() || h === mat.name.toLowerCase());
      items[mat.key] = idx !== -1 ? Number(row[idx]) || 0 : 0;
    });
    pastRecords.push({ id: r + 1, recordId: r + 1, date: formattedDate, avpName: avpVal, division: divVal, destination: destVal, controlNo: controlVal, cluster: regVal, operation: opVal, clusterHead: headVal, clusterHeadContact: contactVal, baseStation: baseVal, notes: notesVal, items });
  }
  pastRecords.reverse();
  return { totals, monthlyOutgoing, transactions, pastRecords };
}

function recordOutgoing(formData) {
  try {
    const sheet = getOutgoingSheet_();
    if (!sheet) throw new Error('Outgoing sheet not found. Expected Outgoing or OutDB.');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h || '').trim());
    const newRow = new Array(headers.length).fill('');
    const newId = 'OUT-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT+8', 'yyyyMMdd-HHmmss');
    headers.forEach((header, index) => {
      const h = header.toLowerCase().trim();
      if (h === 'id') newRow[index] = newId;
      else if (h === 'date') newRow[index] = formData.date ? new Date(formData.date) : new Date();
      else if (h === 'avp name' || h === 'avp') newRow[index] = formData.avpName || '';
      else if (h === 'division') newRow[index] = formData.division || '';
      else if (h === 'destination') newRow[index] = formData.destination || '';
      else if (h.includes('control')) newRow[index] = formData.controlNo || '';
      else if (h === 'region' || h === 'cluster') newRow[index] = formData.region || '';
      else if (h === 'operation') newRow[index] = formData.operation || '';
      else if (h === 'cluster head') newRow[index] = formData.clusterHead || '';
      else if (h.includes('head contact') || h === 'contact') newRow[index] = formData.clusterHeadContact || '';
      else if (h === 'base station' || h.includes('base')) newRow[index] = formData.baseStation || '';
      else if (h === 'notes') newRow[index] = formData.notes || '';
      else {
        const material = MATERIAL_COLUMNS.find(mat => mat.key.toLowerCase() === h || mat.name.toLowerCase() === h);
        if (material) newRow[index] = formData.items && formData.items[material.key] ? Number(formData.items[material.key]) : 0;
      }
    });
    sheet.appendRow(newRow);
    return { success: true, id: newId };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function updateOutgoing(rowIndex, formData) {
  try {
    const sheet = getOutgoingSheet_();
    if (!sheet) throw new Error('Outgoing sheet not found. Expected Outgoing or OutDB.');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h || '').trim());
    const updatedRow = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
    headers.forEach((header, index) => {
      const h = header.toLowerCase().trim();
      if (h === 'date') updatedRow[index] = formData.date ? new Date(formData.date) : new Date();
      else if (h === 'avp name' || h === 'avp') updatedRow[index] = formData.avpName || '';
      else if (h === 'division') updatedRow[index] = formData.division || '';
      else if (h === 'destination') updatedRow[index] = formData.destination || '';
      else if (h.includes('control')) updatedRow[index] = formData.controlNo || '';
      else if (h === 'region' || h === 'cluster') updatedRow[index] = formData.region || '';
      else if (h === 'operation') updatedRow[index] = formData.operation || '';
      else if (h === 'cluster head') updatedRow[index] = formData.clusterHead || '';
      else if (h.includes('head contact') || h === 'contact') updatedRow[index] = formData.clusterHeadContact || '';
      else if (h === 'base station' || h.includes('base')) updatedRow[index] = formData.baseStation || '';
      else if (h === 'notes') updatedRow[index] = formData.notes || '';
      else {
        const material = MATERIAL_COLUMNS.find(mat => mat.key.toLowerCase() === h || mat.name.toLowerCase() === h);
        if (material) updatedRow[index] = formData.items && formData.items[material.key] ? Number(formData.items[material.key]) : 0;
      }
    });
    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([updatedRow]);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}
