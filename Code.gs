// ============================================================
// LMS MICRO MODULE DASHBOARD v8 (OPTIMIZED FOR HUGE DATA)
// Data Source: Single CSV file in Google Drive
// Structure:
//   Cols 0-10 : CasperFHRID, AgentName, HubName, Role, Region,
//               AM, RM, GM, CEC, LBP, Zone
//   Col 11+   : Module columns (L onwards) — dynamic, blank = not assigned
// ============================================================

const CONFIG = {
  CSV_FILE_ID:      '1tUm5BhG1UZv2p6-KwjvGV7QYLxeLZE3K',
  FM_CSV_FILE_ID:   '1tzAETzsJhmbzgji0BcpDkO3Eq5KV703K',
  ADMIN_EMAILS:     ['animesh.jana@flipkart.com'],
  EXPORT_FOLDER_ID: '1w9Nj7hOG1CYqeNE-6-vN7QaRrx5fdorq',
  CACHE_KEY:        'LMS_CSV_DATA_v8_COMPRESSED',
  FM_CACHE_KEY:     'LMS_FM_DATA_v1_COMPRESSED',
  CACHE_SECS:       21600,
  CHUNK_SIZE:       90000
};

const COL = {
  ID: 0, AGENT: 1, HUB: 2, ROLE: 3, REGION: 4, AM: 5, RM: 6, GM: 7, CEC: 8, LBP: 9, ZONE: 10, MOD_START: 11
};

// FM CSV column mapping (0-indexed)
const FM_COL = {
  GM: 0, RM: 1, CEC: 2, ID: 3, AGENT: 4, HUBZONE: 6, MOD_START: 8
};

// ============================================================
// WEB APP ENTRY
// ============================================================
function doGet(e) {
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('LMS Micro Module Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSessionInfo() {
  try {
    const email   = Session.getActiveUser().getEmail().toLowerCase();
    const isAdmin = CONFIG.ADMIN_EMAILS.includes(email);
    return { email, isAdmin };
  } catch(e) {
    return { email: '', isAdmin: false };
  }
}

// ============================================================
// CHUNKED CACHE HELPERS
// ============================================================
function putCache(key, data) {
  const str    = JSON.stringify(data);
  const chunks = [];
  for (let i = 0; i < str.length; i += CONFIG.CHUNK_SIZE)
    chunks.push(str.slice(i, i + CONFIG.CHUNK_SIZE));
  const cache = CacheService.getScriptCache();
  const pairs = {};
  chunks.forEach((c, i) => { pairs[key + '_' + i] = c; });
  pairs[key + '_n'] = String(chunks.length);
  try { cache.putAll(pairs, CONFIG.CACHE_SECS); } catch(e) {}
}

function getCache(key) {
  const cache = CacheService.getScriptCache();
  const n     = cache.get(key + '_n');
  if (!n) return null;
  const count  = Number(n);
  const keys   = Array.from({length: count}, (_, i) => key + '_' + i);
  const values = cache.getAll(keys);
  const chunks = [];
  for (let i = 0; i < count; i++) {
    if (!values[key + '_' + i]) return null;
    chunks.push(values[key + '_' + i]);
  }
  try { return JSON.parse(chunks.join('')); } catch(e) { return null; }
}

function clearCache() {
  const cache = CacheService.getScriptCache();
  [CONFIG.CACHE_KEY, CONFIG.FM_CACHE_KEY].forEach(baseKey => {
    const n = cache.get(baseKey + '_n');
    if (!n) return;
    const count = Number(n);
    const keys  = [baseKey + '_n'];
    for (let i = 0; i < count; i++) keys.push(baseKey + '_' + i);
    cache.removeAll(keys);
  });
}

function readModuleDateMap() {
  try {
    const ss    = SpreadsheetApp.openById('1_KgXFJEPHoOp9VmscVuOagr1pdoUO40-');
    const sheet = ss.getSheetByName('Module Details') || ss.getSheets()[0];
    const data  = sheet.getDataRange().getValues();
    const map   = {};

    for (let i = 0; i < data.length; i++) {
      const row     = data[i];
      const modName = String(row[1] || '').trim();
      const rawDate = row[2];
      const rawDur  = row[3];
      if (!modName) continue;

      let monthKey = '', dateStr = '';
      if (rawDate instanceof Date) {
        monthKey = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), 'MMM-yy');
        dateStr  = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), 'dd-MMM-yy');
      } else if (rawDate) {
        const s = String(rawDate).trim();
        const match = s.match(/([A-Za-z]{3}-\d{2})/);
        if (match) monthKey = match[1];
        dateStr = s;
      }

      let duration = '';
      if (typeof rawDur === 'string' && rawDur.trim()) {
        duration = rawDur.trim();
      } else if (typeof rawDur === 'number' && rawDur > 0) {
        const totalSecs = Math.round(rawDur * 86400);
        const h = Math.floor(totalSecs / 3600);
        const m = Math.floor((totalSecs % 3600) / 60);
        const s = totalSecs % 60;
        duration = h + ':' + String(m).padStart(2,'0') + ':' + String(s).padStart(2,'0');
      } else if (rawDur instanceof Date) {
        const h = rawDur.getUTCHours();
        const m = rawDur.getUTCMinutes();
        const s = rawDur.getUTCSeconds();
        duration = h + ':' + String(m).padStart(2,'0') + ':' + String(s).padStart(2,'0');
      }

      if (modName) map[modName] = { month: monthKey, date: dateStr, duration };
    }
    return map;
  } catch(e) {
    return {};
  }
}

// ============================================================
// MAIN CSV — READ & COMPRESS
// ============================================================
function readCSV() {
  const moduleDateMap = readModuleDateMap();
  const file    = DriveApp.getFileById(CONFIG.CSV_FILE_ID);
  const content = file.getBlob().getDataAsString('UTF-8');
  const parsed  = Utilities.parseCsv(content);

  if (parsed.length < 2) return null;
  const header = parsed[0].map(h => h.trim());

  const modules = [];
  for (let c = COL.MOD_START; c < header.length; c++) {
    const name = header[c] ? header[c].trim() : '';
    if (!name) continue;
    const info = moduleDateMap[name] || {};
    modules.push({ idx: c - COL.MOD_START, name, month: info.month || '', date: info.date || '', duration: info.duration || '' });
  }

  const dict = { hub: [], role: [], region: [], am: [], rm: [], gm: [], cec: [], lbp: [], zone: [] };
  const getIdx = (key, val) => {
    if (!val) return -1;
    let i = dict[key].indexOf(val);
    if (i === -1) { i = dict[key].length; dict[key].push(val); }
    return i;
  };

  const rows = [];
  for (let i = 1; i < parsed.length; i++) {
    const r  = parsed[i];
    const id = String(r[COL.ID] || '').trim();
    if (!id) continue;

    const modStats = modules.map(m => {
      const v = String(r[COL.MOD_START + m.idx] || '').trim().toLowerCase();
      if (v === 'completed') return 'C';
      if (v === 'in progress') return 'I';
      if (v === 'not started') return 'N';
      return '';
    });

    if (modStats.every(s => !s)) continue;

    rows.push([
      id,
      String(r[COL.AGENT] || '').trim(),
      getIdx('hub',    String(r[COL.HUB]    || '').trim()),
      getIdx('role',   String(r[COL.ROLE]   || '').trim()),
      getIdx('region', String(r[COL.REGION] || '').trim()),
      getIdx('am',     String(r[COL.AM]     || '').trim()),
      getIdx('rm',     String(r[COL.RM]     || '').trim()),
      getIdx('gm',     String(r[COL.GM]     || '').trim()),
      getIdx('cec',    String(r[COL.CEC]    || '').trim()),
      getIdx('lbp',    String(r[COL.LBP]    || '').trim()),
      getIdx('zone',   String(r[COL.ZONE]   || '').trim()),
      ...modStats
    ]);
  }

  return { rows, modules, dict };
}

function loadDashboardData() {
  try {
    const cached = getCache(CONFIG.CACHE_KEY);
    if (cached) return { ...cached, error: null };

    const data = readCSV();
    if (!data) return { error: '❌ CSV file is empty or unreadable.', rows: [], modules: [], dict: {} };

    putCache(CONFIG.CACHE_KEY, data);
    return { ...data, error: null };
  } catch(e) {
    return { error: '❌ ' + e.message, rows: [], modules: [], dict: {} };
  }
}

function refreshData() {
  try {
    clearCache();
    const start = new Date();
    const data  = readCSV();
    if (!data) return { success: false, message: '❌ CSV file is empty or unreadable.' };

    putCache(CONFIG.CACHE_KEY, data);
    const secs = ((new Date() - start) / 1000).toFixed(1);

    return {
      success: true,
      message: '✅ Data refreshed!',
      details: data.modules.length + ' modules · ' + data.rows.length.toLocaleString() + ' employees · ' + secs + 's',
    };
  } catch(e) {
    return { success: false, message: '❌ ' + e.message };
  }
}

// ============================================================
// FM CSV — READ & COMPRESS
// FM Row: [id, agent, gm_idx, rm_idx, cec_idx, hubzone_idx, ...modStats]
// ============================================================
function readFMCSV() {
  const file    = DriveApp.getFileById(CONFIG.FM_CSV_FILE_ID);
  const content = file.getBlob().getDataAsString('UTF-8');
  const parsed  = Utilities.parseCsv(content);

  if (parsed.length < 2) return null;
  const header = parsed[0].map(h => h.trim());

  const modules = [];
  for (let c = FM_COL.MOD_START; c < header.length; c++) {
    const name = header[c] ? header[c].trim() : '';
    if (!name) continue;
    modules.push({ idx: c - FM_COL.MOD_START, name, month: '', date: '', duration: '' });
  }

  const dict = { gm: [], rm: [], cec: [], hubzone: [] };
  const getIdx = (key, val) => {
    if (!val) return -1;
    let i = dict[key].indexOf(val);
    if (i === -1) { i = dict[key].length; dict[key].push(val); }
    return i;
  };

  const rows = [];
  for (let i = 1; i < parsed.length; i++) {
    const r  = parsed[i];
    const id = String(r[FM_COL.ID] || '').trim();
    if (!id) continue;

    const modStats = modules.map(m => {
      const v = String(r[FM_COL.MOD_START + m.idx] || '').trim().toLowerCase();
      if (!v || v === '-') return '';  // dash = not assigned, exclude
      if (v === 'completed') return 'C';
      if (v === 'in progress') return 'I';
      if (v === 'not started') return 'N';
      return '';
    });

    if (modStats.every(s => !s)) continue;

    rows.push([
      id,
      String(r[FM_COL.AGENT]   || '').trim(),
      getIdx('gm',      String(r[FM_COL.GM]      || '').trim()),
      getIdx('rm',      String(r[FM_COL.RM]       || '').trim()),
      getIdx('cec',     String(r[FM_COL.CEC]      || '').trim()),
      getIdx('hubzone', String(r[FM_COL.HUBZONE]  || '').trim()),
      ...modStats
    ]);
  }

  return { rows, modules, dict };
}

function loadFMDashboardData() {
  try {
    const cached = getCache(CONFIG.FM_CACHE_KEY);
    if (cached) return { ...cached, error: null };

    const data = readFMCSV();
    if (!data) return { error: '❌ FM CSV file is empty or unreadable.', rows: [], modules: [], dict: {} };

    putCache(CONFIG.FM_CACHE_KEY, data);
    return { ...data, error: null };
  } catch(e) {
    return { error: '❌ ' + e.message, rows: [], modules: [], dict: {} };
  }
}

// ============================================================
// DOWNLOAD GENERATOR
// ============================================================
function generateDownload(filters, selectedModuleNames, selectedStatuses) {
  try {
    const data = getCache(CONFIG.CACHE_KEY) || readCSV();
    if (!data) return { error: 'No data. Please reload the dashboard first.' };

    const { rows, modules, dict } = data;
    const getVal = (key, idx) => idx === -1 ? '' : dict[key][idx];
    const mapMod = (v) => v === 'C' ? 'Completed' : v === 'I' ? 'In Progress' : v === 'N' ? 'Not Started' : '';

    const selMods = selectedModuleNames && selectedModuleNames.length
      ? modules.filter(m => selectedModuleNames.includes(m.name))
      : modules;

    const sl = selectedStatuses ? selectedStatuses.map(s => s.toLowerCase()) : [];

    const expHeaders = [
      'Employee ID','Agent Name','Hub','Role','Region','AM','RM','GM','CEC','LBP','Zone',
      ...selMods.map(m => m.name)
    ];

    const escapeCsv = (val) => {
      let str = String(val || '');
      if (str.includes(',') || str.includes('"') || str.includes('\n')) {
        return '"' + str.replace(/"/g, '""') + '"';
      }
      return str;
    };

    const csvRows = [];
    csvRows.push(expHeaders.map(escapeCsv).join(','));

    let count = 0;
    rows.forEach(r => {
      const rowZone   = getVal('zone', r[10]);
      const rowGm     = getVal('gm', r[7]);
      const rowRegion = getVal('region', r[4]);
      const rowRm     = getVal('rm', r[6]);
      const rowAm     = getVal('am', r[5]);
      const rowLbp    = getVal('lbp', r[9]);
      const rowCec    = getVal('cec', r[8]);

      if (filters.zone   && filters.zone   !== 'All' && rowZone   !== filters.zone)   return;
      if (filters.gm     && filters.gm     !== 'All' && rowGm     !== filters.gm)     return;
      if (filters.region && filters.region !== 'All' && rowRegion !== filters.region) return;
      if (filters.rm     && filters.rm     !== 'All' && rowRm     !== filters.rm)     return;
      if (filters.am     && filters.am     !== 'All' && rowAm     !== filters.am)     return;
      if (filters.lbp    && filters.lbp    !== 'All' && rowLbp    !== filters.lbp)    return;
      if (filters.cec    && filters.cec    !== 'All' && rowCec    !== filters.cec)    return;

      const modVals = selMods.map(m => mapMod(r[11 + m.idx]) || '');

      if (sl.length) {
        const hasMatch = modVals.some(v => v && sl.includes(v.toLowerCase()));
        if (!hasMatch) return;
      } else {
        if (modVals.every(v => !v)) return;
      }

      const modDisplay = modVals.map(v => v || 'Not Assigned');
      const rowArr = [
        r[0], r[1], getVal('hub', r[2]), getVal('role', r[3]), rowRegion, rowAm, rowRm, rowGm, rowCec, rowLbp, rowZone,
        ...modDisplay
      ];

      csvRows.push(rowArr.map(escapeCsv).join(','));
      count++;
    });

    if (count === 0) return { error: 'No records match the selected filters.' };

    const csvContent = csvRows.join('\n');
    const fileName = 'LMS_Export_' + new Date().toISOString().slice(0,10) + '.csv';

    const folder = DriveApp.getFolderById(CONFIG.EXPORT_FOLDER_ID);
    const file = folder.createFile(fileName, csvContent, MimeType.CSV);

    return { url: file.getUrl(), name: file.getName(), count: count };
  } catch(e) {
    return { error: '❌ ' + e.message };
  }
}
