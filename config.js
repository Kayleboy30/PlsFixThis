/**
 * Global Configuration & 13 Tracked Materials
 *
 * The project has existed with both the newer Incoming/Outgoing sheet names
 * and the original IncDB/OutDB names. Keep the canonical names here while
 * allowing the data handlers to fall back to the legacy sheets.
 */
const SHEET_NAMES = {
  INCOMING: 'Incoming',
  OUTGOING: 'Outgoing',
  INVENTORY: 'Inventory',
  DATABASE: 'Database',
  USERS: 'Users',
  DIVISION_BUDGET: 'Division Budget',

  // Legacy sheet names used by the original inventory system.
  LEGACY_INCOMING: 'IncDB',
  LEGACY_OUTGOING: 'OutDB',
  LEGACY_INVENTORY: 'Stock Balance'
};

const MATERIAL_COLUMNS = [
  { key: 'FAF', name: 'Financial Agreement Form' },
  { key: 'Passbook', name: 'Passbook' },
  { key: 'Desk Calendar', name: 'Desk Calendar' },
  { key: 'GTR', name: 'Group Treasurer Register' },
  { key: 'Wall Calendar', name: 'Wall Calendar' },
  { key: 'Guide Book', name: 'Guide Book' },
  { key: 'FAF Barmm', name: 'Financial Agreement Form (BARMM)' },
  { key: 'Passbook Barmm', name: 'Passbook (BARMM)' },
  { key: 'GTR Barmm', name: 'Group Treasurer Register (BARMM)' },
  { key: 'Poster', name: 'Poster with Acrylic' },
  { key: 'GTR New', name: 'GTR New' },
  { key: 'Enrolment', name: 'Asa Insurance Enrollment Forms' },
  { key: 'Coverage', name: 'Asa Insurance Coverage' }
];

function formatDateMonth(d) {
  if (!d) return null;
  const date = new Date(d);
  if (isNaN(date.getTime())) return null;
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const yr = String(date.getFullYear()).slice(-2);
  return `${months[date.getMonth()]} ${yr}`;
}

/**
 * Return the first existing sheet from the supplied names.
 * This keeps the application compatible with both the current and legacy
 * spreadsheet layouts without requiring the user to rename existing data.
 */
function getSheetByNames_(ss, names) {
  for (const name of names) {
    if (!name) continue;
    const sheet = ss.getSheetByName(name);
    if (sheet) return sheet;
  }
  return null;
}
