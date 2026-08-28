/**
 * ============================================================================
 * CLUSTER & ADDRESS DIRECTORY (ClusterAddress.gs)
 * ============================================================================
 */

function fetchClusterAddressData() {
  try {
    const csvUrl = 'https://docs.google.com/spreadsheets/d/18YPh-vQ6EN4P5sLxtVOpYHvWWiEzrXWVCAsKjO21l4s/export?format=csv&gid=1633170149';
    const response = UrlFetchApp.fetch(csvUrl, { muteHttpExceptions: true });
    const csvText = response.getContentText();
    const rows = Utilities.parseCsv(csvText);

    if (!rows || rows.length < 3) return { success: true, data: [] };

    const records = [];
    let currentAvp = '';
    let currentDivision = '';

    for (let i = 2; i < rows.length; i++) {
      const row = rows[i];
      if (row[0] && row[0].trim() !== '') currentAvp = row[0].trim();
      if (row[1] && row[1].trim() !== '') currentDivision = row[1].trim();

      const cluster = row[2] ? row[2].trim() : '';
      const idNum = row[3] ? row[3].trim() : '';
      const head = row[4] ? row[4].trim() : '';
      const contact = row[5] ? row[5].trim() : '';
      const email = row[6] ? row[6].trim() : '';
      const station = row[7] ? row[7].trim() : '';
      const address = row[8] ? row[8].trim() : '';

      if (!cluster && !station && !head) continue;

      records.push({
        avpName: currentAvp || (row[9] ? row[9].trim() : ''),
        division: currentDivision,
        cluster: cluster,
        idNumber: idNum,
        clusterHead: head,
        contact: contact,
        email: email,
        baseStation: station,
        completeAddress: address
      });
    }

    return { success: true, data: records, count: records.length };
  } catch (err) {
    return { success: false, error: err.message };
  }
}
