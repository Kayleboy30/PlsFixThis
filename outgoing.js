/**
 * ==========================================
 * OUTGOING MODULE (Dispatches / Gate Pass)
 * ==========================================
 */

/**
 * Reads Cluster, Cluster Head, Contact, Base Station, and Complete Address from 'Database' sheet
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

    // ── STEP 1: Build Master AVP Info Map (Columns A-D) ─────────────────────
    const dbSheet = ss.getSheetByName(SHEET_NAMES.DATABASE || 'Database');
    if (dbSheet) {
      const data = dbSheet.getDataRange().getValues();
      if (data.length > 1) {
        const headers = data[0].map(h => String(h).trim().toLowerCase());
        let avpColIdx  = headers.findIndex(h => h.includes('avp'));
        let divColIdx  = headers.findIndex(h => h.includes('division'));
        let opColIdx   = headers.findIndex(h => h.includes('operation'));
        let destColIdx = headers.findIndex(h => h.includes('destination') || h.includes('complete address'));

        if (avpColIdx === -1) avpColIdx = 0;
        if (divColIdx === -1) divColIdx = 1;
        if (opColIdx === -1)  opColIdx = 2;
        if (destColIdx === -1) destColIdx = 3;

        for (let i = 1; i < data.length; i++) {
          const aName = String(data[i][avpColIdx] || '').trim();
          if (!aName || aName === '-' || aName.toLowerCase() === 'n/a') continue;

          const aDiv  = String(data[i][divColIdx] || '').trim();
          const aOp   = String(data[i][opColIdx] || '').trim();
          const aDest = String(data[i][destColIdx] || '').trim();

          const info = { division: aDiv, operation: aOp, destination: aDest };
          avpMasterMap[aName.toLowerCase()] = info;
          avpMasterMap[aName.toLowerCase().replace(/\s+/g, '')] = info;
        }

        // ── STEP 2: Read Cluster Directory (Columns T-X) ───────────────────
        let clusterIdx = headers.findIndex(h => (h.includes('cluster') || h.includes('region')) && !h.includes('head') && !h.includes('contact'));
        let headIdx    = headers.findIndex(h => h.includes('head') && !h.includes('contact'));
        let divIdx     = headers.findIndex(h => h.includes('division'));
        let opIdx      = headers.findIndex(h => h.includes('operation'));
        let addrIdx    = headers.findIndex(h => h.includes('address') || h.includes('destination'));
        const colXIdx  = 23; // Column X (AVP assigned to Cluster)
        // ⚠️ FIX: Find contact column that appears AFTER the cluster column (not AVP contact)
        let contactIdx = headers.findIndex(h => h.includes('cluster') && h.includes('contact'));
        if (contactIdx === -1) {
          // Find any contact column AFTER the cluster column to avoid picking AVP contact
          contactIdx = headers.findIndex((h, i) => h.includes('contact') && i > clusterIdx);
        }
        // ⚠️ FIX: Find branch column that appears AFTER the cluster column
        let branchIdx = headers.findIndex(h => h.includes('branch') || h.includes('station') || h.includes('base'));
        if (branchIdx !== -1 && branchIdx < clusterIdx) {
          // Found a branch-like column before the cluster section — skip it and find one after
          branchIdx = headers.findIndex((h, i) => (h.includes('branch') || h.includes('station') || h.includes('base')) && i > clusterIdx);
        }

        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          const rawCluster = clusterIdx !== -1 ? String(row[clusterIdx] || '').trim() : '';
          if (!rawCluster || rawCluster === '-' || rawCluster.toLowerCase() === 'n/a') continue;

          const norm = normalizeCluster(rawCluster);
          clusterNamesSet.add(norm);

          // Get Assigned AVP from Column X
          const assignedAvp = (row[colXIdx] && String(row[colXIdx]).trim() && String(row[colXIdx]).trim() !== '-')
            ? String(row[colXIdx]).trim()
            : '';

          // Look up this AVP's real division & operation from Master AVP Map
          const avpKey = assignedAvp.toLowerCase();
          const avpKeyClean = avpKey.replace(/\s+/g, '');
          const avpDetails = avpMasterMap[avpKey] || avpMasterMap[avpKeyClean] || {};

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

    // ── STEP 3: Fallback from OutDB for any missing contact / station ────────
    const outSheet = ss.getSheetByName('OutDB') || ss.getSheetByName(SHEET_NAMES.OUTGOING || 'Outgoing');
    if (outSheet) {
      const outData = outSheet.getDataRange().getValues();
      if (outData.length > 1) {
        const headers = outData[0].map(h => String(h).trim().toLowerCase());
        let clusterIdx = headers.findIndex(h => (h.includes('cluster') || h.includes('region')) && !h.includes('head') && !h.includes('contact'));
        let headIdx    = headers.findIndex(h => h.includes('head') && !h.includes('contact'));
        let contactIdx = headers.findIndex(h => h.includes('contact'));
        let branchIdx  = headers.findIndex(h => h.includes('branch') || h.includes('station') || h.includes('base'));
        let avpIdx     = headers.findIndex(h => h.includes('avp'));
        let divIdx     = headers.findIndex(h => h.includes('division'));
        let opIdx      = headers.findIndex(h => h.includes('operation'));
        let destIdx    = headers.findIndex(h => h.includes('destination'));

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
              cluster: norm,
              avpName: outAvp,
              division: avpDetails.division || (divIdx !== -1 ? String(row[divIdx] || '').trim() : ''),
              operation: avpDetails.operation || (opIdx !== -1 ? String(row[opIdx] || '').trim() : ''),
              destination: avpDetails.destination || (destIdx !== -1 ? String(row[destIdx] || '').trim() : ''),
              clusterHead: headIdx !== -1 ? String(row[headIdx] || '').trim() : '',
              clusterHeadContact: contactIdx !== -1 ? String(row[contactIdx] || '').trim() : '',
              baseStation: branchIdx !== -1 ? String(row[branchIdx] || '').trim() : ''
            };
          } else {
            const existing = directory[norm];
            if (!existing.clusterHeadContact && contactIdx !== -1 && row[contactIdx]) existing.clusterHeadContact = String(row[contactIdx]).trim();
            if (!existing.baseStation && branchIdx !== -1 && row[branchIdx]) existing.baseStation = String(row[branchIdx]).trim();
            if (!existing.clusterHead && headIdx !== -1 && row[headIdx]) existing.clusterHead = String(row[headIdx]).trim();
          }
        }
      }
    }

    return {
      directory: directory,
      clusterList: Array.from(clusterNamesSet).sort((a, b) => {
        const numA = parseInt(a.replace(/\D/g, '')) || 0;
        const numB = parseInt(b.replace(/\D/g, '')) || 0;
        return numA - numB;
      })
    };
  } catch (err) {
    return { directory: {}, clusterList: [], error: err.message };
  }
}

/**
 * Reads AVP, Division, Operation, Destination, and Complete Address from 'Database' sheet
 */
function fetchAvpDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Database') || 
              ss.getSheetByName('database') || 
              ss.getSheetByName('DB');

  const avpDirectory = {};
  const avpList = [];

  if (!sheet || sheet.getLastRow() <= 1) {
    return { avpDirectory, avpList };
  }

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim().toLowerCase());

  const avpIdx     = headers.findIndex(h => h.includes('avp'));
  const divIdx     = headers.findIndex(h => h.includes('division'));
  const opIdx      = headers.findIndex(h => h.includes('operation'));
  const destIdx    = headers.findIndex(h => h.includes('destination'));
  const addressIdx = headers.findIndex(h => h.includes('complete address') || h === 'address');

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const avpName     = (avpIdx !== -1 && row[avpIdx] !== undefined) ? String(row[avpIdx]).trim() : '';
    const division    = (divIdx !== -1 && row[divIdx] !== undefined) ? String(row[divIdx]).trim() : '';
    const operation   = (opIdx !== -1 && row[opIdx] !== undefined) ? String(row[opIdx]).trim() : '';
    const destination = (destIdx !== -1 && row[destIdx] !== undefined) ? String(row[destIdx]).trim() : '';
    const address     = (addressIdx !== -1 && row[addressIdx] !== undefined) ? String(row[addressIdx]).trim() : '';

    if (avpName) {
      const info = {
        division: division,
        operation: operation,
        destination: destination,
        completeAddress: address
      };

      if (!avpDirectory[avpName]) {
        avpDirectory[avpName] = info;
        avpDirectory[avpName.toLowerCase()] = info;
        avpList.push(avpName);
      }
    }
  }

  return { avpDirectory, avpList: avpList.sort() };
}

function fetchOutgoingData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const outSheet = ss.getSheetByName(SHEET_NAMES.OUTGOING);

  const totals = {};
  const monthlyOutgoing = {};
  const transactions = [];
  const pastRecords = [];

  MATERIAL_COLUMNS.forEach(mat => { totals[mat.key] = 0; });

  if (!outSheet || outSheet.getLastRow() <= 1) {
    return { totals, monthlyOutgoing, transactions, pastRecords };
  }

  const values  = outSheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());

  const dateIdx    = headers.findIndex(h => h.toLowerCase() === 'date');
  const avpIdx     = headers.findIndex(h => h.toLowerCase() === 'avp name' || h.toLowerCase() === 'avp');
  const divIdx     = headers.findIndex(h => h.toLowerCase() === 'division');
  const destIdx    = headers.findIndex(h => h.toLowerCase() === 'destination');
  const controlIdx = headers.findIndex(h => h.toLowerCase().includes('control'));
  const regIdx     = headers.findIndex(h => h.toLowerCase().includes('region') || (h.toLowerCase().includes('cluster') && !h.toLowerCase().includes('head')));
  const opIdx      = headers.findIndex(h => h.toLowerCase() === 'operation');
  const headIdx    = headers.findIndex(h => h.toLowerCase() === 'cluster head');
  const contactIdx = headers.findIndex(h => h.toLowerCase().includes('head contact') || h.toLowerCase().includes('contact'));
  const baseIdx    = headers.findIndex(h => h.toLowerCase() === 'base station' || h.toLowerCase().includes('base'));
  const notesIdx   = headers.findIndex(h => h.toLowerCase() === 'notes');

  function parseTransactionDate(value) {
    if (!value) return null;

    if (Object.prototype.toString.call(value) === '[object Date]') {
      if (!isNaN(value.getTime())) return value;
    }

    const str = String(value).trim();
    if (!str) return null;

    // Smart DD/MM/YYYY vs MM/DD/YYYY detection
    const match = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
    if (match) {
      const p1 = Number(match[1]), p2 = Number(match[2]), year = Number(match[3]);
      let month, day;

      if (p1 > 12)      { day = p1; month = p2; }   // DD/MM/YYYY
      else if (p2 > 12) { month = p1; day = p2; }   // MM/DD/YYYY
      else              { day = p1; month = p2; }   // Both ≤ 12 → DD/MM (PH standard)

      const date = new Date(year, month - 1, day);
      if (!isNaN(date.getTime())) return date;
    }

    const parsed = new Date(str);
    return isNaN(parsed.getTime()) ? null : parsed;
  }

  for (let r = 1; r < values.length; r++) {
    const row             = values[r];
    const transactionDate = parseTransactionDate(dateIdx !== -1 ? row[dateIdx] : null);
    const rawDate         = transactionDate ? transactionDate.getTime() : 0;
    const monthStr        = transactionDate ? Utilities.formatDate(transactionDate, Session.getScriptTimeZone(), 'MMM yy') : '';

    const avpVal     = avpIdx     !== -1 && row[avpIdx]    ? String(row[avpIdx]).trim()    : '-';
    const divVal     = divIdx     !== -1 && row[divIdx]    ? String(row[divIdx]).trim()    : '-';
    const destVal    = destIdx    !== -1 && row[destIdx]   ? String(row[destIdx]).trim()   : '-';
    const controlVal = controlIdx !== -1 && row[controlIdx]? String(row[controlIdx]).trim(): '-';
    const regVal     = regIdx     !== -1 && row[regIdx]    ? String(row[regIdx]).trim()    : '-';
    const opVal      = opIdx      !== -1 && row[opIdx]     ? String(row[opIdx]).trim()     : '';
    const headVal    = headIdx    !== -1 && row[headIdx]   ? String(row[headIdx]).trim()   : '';
    const contactVal = contactIdx !== -1 && row[contactIdx]? String(row[contactIdx]).trim(): '';
    const baseVal    = baseIdx    !== -1 && row[baseIdx]   ? String(row[baseIdx]).trim()   : '';
    const notesVal   = notesIdx   !== -1 && row[notesIdx]  ? String(row[notesIdx]).trim()  : '';

    const itemQuantities = {};
    let rowHasData = false;

    MATERIAL_COLUMNS.forEach(mat => {
      const colIdx = headers.findIndex(h => h.toLowerCase() === mat.key.toLowerCase());
      if (colIdx !== -1) {
        const qty = parseFloat(row[colIdx]) || 0;
        if (qty > 0) {
          rowHasData = true;
          itemQuantities[mat.key] = qty;
          totals[mat.key] += qty;
          if (monthStr) {
            if (!monthlyOutgoing[monthStr]) monthlyOutgoing[monthStr] = {};
            if (!monthlyOutgoing[monthStr][mat.key]) monthlyOutgoing[monthStr][mat.key] = 0;
            monthlyOutgoing[monthStr][mat.key] += qty;
          }
        }
      }
    });

    if (rowHasData) {
      const dateString = transactionDate
        ? Utilities.formatDate(transactionDate, Session.getScriptTimeZone(), 'M/d/yyyy')
        : 'N/A';

      transactions.push({ type: 'Outgoing', date: dateString, party: `${destVal} (${avpVal})`.trim(), ref: controlVal, rawDate, rowIndex: r + 1 });

      pastRecords.push({ id: r, date: dateString, controlNo: controlVal, avpName: avpVal, division: divVal, destination: destVal, cluster: regVal, operation: opVal, clusterHead: headVal, clusterHeadContact: contactVal, baseStation: baseVal, notes: notesVal, items: itemQuantities, rawDate, rowIndex: r + 1 });
    }
  }

  pastRecords.sort((a, b) => b.rawDate !== a.rawDate ? b.rawDate - a.rawDate : b.rowIndex - a.rowIndex);
  transactions.sort((a, b) => b.rawDate !== a.rawDate ? b.rawDate - a.rawDate : b.rowIndex - a.rowIndex);

  return { totals, monthlyOutgoing, transactions, pastRecords };
}



/**
 * Appends a NEW outgoing record to the 'Outgoing' sheet
 */
function recordOutgoing(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.OUTGOING);

  if (!sheet) throw new Error(`Sheet '${SHEET_NAMES.OUTGOING}' not found.`);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const newRow = new Array(headers.length).fill('');
  const nextId = sheet.getLastRow();

  headers.forEach((header, idx) => {
    const hLower = header.toLowerCase();
    if      (hLower === 'id')            newRow[idx] = nextId;
    else if (hLower === 'date')          newRow[idx] = formData.date;
    else if (hLower === 'avp name')      newRow[idx] = formData.avpName || '';
    else if (hLower === 'division')      newRow[idx] = formData.division || '';
    else if (hLower.includes('region') || (hLower.includes('cluster') && !hLower.includes('head')))
                                         newRow[idx] = formData.regionCluster || '';
    else if (hLower === 'destination')   newRow[idx] = formData.destination || '';
    else if (hLower.includes('control')) newRow[idx] = formData.controlNo || '';
    else if (hLower === 'operation')     newRow[idx] = formData.operation || '';
    else if (hLower.includes('head contact') || hLower.includes('contact'))
                                         newRow[idx] = formData.clusterHeadContact || '';
    else if (hLower.includes('cluster head') || hLower === 'head')
                                         newRow[idx] = formData.clusterHead || '';
    else if (hLower === 'base station' || hLower.includes('base') || hLower.includes('station'))
                                         newRow[idx] = formData.baseStation || '';
    else if (hLower === 'notes')         newRow[idx] = formData.notes || '';
    else {
      const mat = MATERIAL_COLUMNS.find(m => m.key.toLowerCase() === hLower);
      if (mat && formData.items && formData.items[mat.key]) {
        newRow[idx] = parseFloat(formData.items[mat.key]) || 0;
      }
    }
  });

  sheet.appendRow(newRow);
  return { success: true };
}

/**
 * UPDATES an EXISTING outgoing record in the 'Outgoing' sheet by row index
 */
function updateOutgoing(rowIndex, formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.OUTGOING);

  if (!sheet) throw new Error(`Sheet '${SHEET_NAMES.OUTGOING}' not found.`);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const existingRow = sheet.getRange(rowIndex + 1, 1, 1, headers.length).getValues()[0];
  const updatedRow = existingRow.slice();

  headers.forEach((header, idx) => {
    const hLower = header.toLowerCase();
    if      (hLower === 'id')            { /* Keep ID unchanged */ }
    else if (hLower === 'date')          updatedRow[idx] = formData.date;
    else if (hLower === 'avp name')      updatedRow[idx] = formData.avpName || '';
    else if (hLower === 'division')      updatedRow[idx] = formData.division || '';
    else if (hLower.includes('region') || (hLower.includes('cluster') && !hLower.includes('head')))
                                         updatedRow[idx] = formData.regionCluster || '';
    else if (hLower === 'destination')   updatedRow[idx] = formData.destination || '';
    else if (hLower.includes('control')) updatedRow[idx] = formData.controlNo || '';
    else if (hLower === 'operation')     updatedRow[idx] = formData.operation || '';
    else if (hLower.includes('head contact') || hLower.includes('contact'))
                                         updatedRow[idx] = formData.clusterHeadContact || '';
    else if (hLower.includes('cluster head') || hLower === 'head')
                                         updatedRow[idx] = formData.clusterHead || '';
    else if (hLower === 'base station' || hLower.includes('base') || hLower.includes('station'))
                                         updatedRow[idx] = formData.baseStation || '';
    else if (hLower === 'notes')         updatedRow[idx] = formData.notes || '';
    else {
      const mat = MATERIAL_COLUMNS.find(m => m.key.toLowerCase() === hLower);
      if (mat) {
        updatedRow[idx] = (formData.items && formData.items[mat.key]) ? parseFloat(formData.items[mat.key]) || 0 : 0;
      }
    }
  });

  sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([updatedRow]);
  return { success: true };
}
