/**
 * FACULTY eLOG V2
 * Google Apps Script JSON API for the static GitHub Pages frontend.
 *
 * Bind this script to the existing Google Spreadsheet and deploy it as a Web App.
 * Google Sheets remains the source of truth.
 */

const SHEETS = {
  LOGBOOK: 'Logbook',
  CONSUMABLES: 'Consumables',
  GLASSWARE: 'Glassware',
  EQUIPMENT: 'Equipment',
  SPECIALIZED: 'Specialized Equipment',
  FORMS: 'Laboratory Forms',
  ROOMS: 'Rooms',
  SETTINGS: 'Settings'
};

const APP_TIME_ZONE = 'Asia/Manila';

const HEADERS = {
  'Logbook': [
    'Timestamp','Faculty Name','Department','Room Name','Room Number',
    'Equipment Used','Date','Time In','Time Out','Activity Type',
    'Purpose','Remarks'
  ],
  'Consumables': [
    'Item Name',
    'Quantity',
    'Unit',
    'Availability',
    'Expiration Date',
    'Remarks'
  ],
  'Glassware': ['Glassware Name','Quantity','Unit','Remarks'],
  'Equipment': ['Equipment Name','Quantity','Unit','Condition','Remarks'],
  'Specialized Equipment': ['Equipment Name','Quantity','Unit','Location','Condition','Remarks'],
  'Laboratory Forms': ['Form Name','Description','File / Link','Status'],
  'Rooms': ['Room Number','Room Name','Building','Status'],
  'Settings': ['Category','Value','Status']
};

/** Handles read-only API actions from the GitHub Pages frontend. */
function doGet(event) {
  const parameters = event && event.parameter ? event.parameter : {};
  const action = clean_(parameters.action).toLowerCase();

  try {
    let data;
    if (action === 'initialdata') {
      data = getInitialData();
    } else if (action === 'readonlydata') {
      data = getReadOnlyData();
    } else if (action === 'logbook') {
      data = getLogbook({
        search: parameters.search,
        department: parameters.department,
        startDate: parameters.startDate,
        endDate: parameters.endDate
      });
    } else if (action === 'reports') {
      data = getReports(parameters.month);
    } else {
      return jsonResponse_({
        success: false,
        error: 'Unsupported action. Use initialData, readOnlyData, logbook, or reports.'
      });
    }

    return jsonResponse_({
      success: true,
      data: data,
      updatedAt: new Date().toISOString()
    });
  } catch (error) {
    logApiError_('GET ' + (action || '(missing action)'), error);
    return jsonResponse_({
      success: false,
      error: 'Unable to load Faculty eLog data. Please try again.'
    });
  }
}

/** Handles the Logbook's only website write action. */
function doPost(event) {
  let action = '';
  try {
    const request = parsePostRequest_(event);
    action = clean_(request.action).toLowerCase();
    if (action !== 'submitlog') {
      return jsonResponse_({
        success: false,
        error: 'Unsupported POST action. Use submitLog.'
      });
    }

    const result = submitLog({
      facultyName: request.facultyName,
      department: request.department,
      roomNumber: request.roomNumber,
      roomName: request.roomName,
      equipmentUsed: request.equipmentUsed,
      timeOut: request.timeOut,
      activityType: request.activityType,
      purpose: request.purpose,
      remarks: request.remarks
    });

    return jsonResponse_({
      success: true,
      data: result,
      updatedAt: new Date().toISOString()
    });
  } catch (error) {
    logApiError_('POST ' + (action || '(missing action)'), error);
    return jsonResponse_({
      success: false,
      error: friendlySubmitError_(error)
    });
  }
}

function parsePostRequest_(event) {
  const parameters = event && event.parameter ? event.parameter : {};
  const request = {};
  Object.keys(parameters).forEach(key => request[key] = parameters[key]);

  const contentType = clean_(event && event.postData && event.postData.type).toLowerCase();
  const contents = clean_(event && event.postData && event.postData.contents);
  if (contents && contentType.indexOf('application/json') === 0) {
    const parsed = JSON.parse(contents);
    if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
      throw new Error('Invalid JSON request body.');
    }
    Object.keys(parsed).forEach(key => request[key] = parsed[key]);
  }
  return request;
}

function friendlySubmitError_(error) {
  const message = clean_(error && error.message);
  if (/ is required\.$/.test(message) || message === 'Time Out must be later than the server-recorded Time In.') {
    return message;
  }
  return 'Unable to save your log. Please try again.';
}

function jsonResponse_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function logApiError_(context, error) {
  console.error('[Faculty eLog API] ' + context + ': ' +
    (error && error.stack ? error.stack : error));
}

/**
 * Run once after binding the project to the target spreadsheet.
 * Creates missing sheets and headers. Existing data is preserved.
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSpreadsheetTimeZone_(ss);

  Object.keys(HEADERS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);

    const headers = HEADERS[name];
    const firstRow = sh.getRange(1, 1, 1, headers.length).getValues()[0];
    const isBlank = firstRow.every(v => String(v).trim() === '');

    if (isBlank) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#0f172a')
      .setFontColor('#ffffff');

    sh.autoResizeColumns(1, headers.length);
  });

  // Keep Logbook date/time values sortable while displaying them in the required formats.
  const logbookSheet = ss.getSheetByName(SHEETS.LOGBOOK);
  if (logbookSheet) {
    const rowCount = Math.max(logbookSheet.getLastRow() - 1, 1);
    logbookSheet.getRange(2, 1, rowCount, 1).setNumberFormat('M/d/yyyy h:mm:ss AM/PM');
    logbookSheet.getRange(2, 7, rowCount, 1).setNumberFormat('M/d/yyyy');
    logbookSheet.getRange(2, 8, rowCount, 2).setNumberFormat('h:mm AM/PM');
  }

  seedSettings_();
  return 'Faculty eLog sheets are ready.';
}

function getInitialData() {
  return {
    rooms: getSheetObjects_(SHEETS.ROOMS),
    consumables: getSheetObjects_(SHEETS.CONSUMABLES),
    glassware: getSheetObjects_(SHEETS.GLASSWARE),
    equipment: getSheetObjects_(SHEETS.EQUIPMENT),
    specializedEquipment: getSheetObjects_(SHEETS.SPECIALIZED),
    laboratoryForms: getSheetObjects_(SHEETS.FORMS),
    settings: getSheetObjects_(SHEETS.SETTINGS),
    logbook: getSheetObjects_(SHEETS.LOGBOOK)
  };
}

function getReadOnlyData() {
  return {
    rooms: getSheetObjects_(SHEETS.ROOMS),
    consumables: getSheetObjects_(SHEETS.CONSUMABLES),
    glassware: getSheetObjects_(SHEETS.GLASSWARE),
    equipment: getSheetObjects_(SHEETS.EQUIPMENT),
    specializedEquipment: getSheetObjects_(SHEETS.SPECIALIZED),
    laboratoryForms: getSheetObjects_(SHEETS.FORMS),
    settings: getSheetObjects_(SHEETS.SETTINGS)
  };
}

function submitLog(entry) {
  const now = new Date();
  const timeOut = validateLog_(entry, now);
  const sh = getRequiredSheet_(SHEETS.LOGBOOK);
  ensureSpreadsheetTimeZone_(sh.getParent());
  const row = [
    now,
    clean_(entry.facultyName),
    clean_(entry.department),
    clean_(entry.roomName),
    clean_(entry.roomNumber),
    clean_(entry.equipmentUsed),
    now,
    now,
    timeOut,
    clean_(entry.activityType),
    clean_(entry.purpose),
    clean_(entry.remarks)
  ];

  sh.appendRow(row);

  const logRow = sh.getLastRow();
  sh.getRange(logRow, 1).setNumberFormat('M/d/yyyy h:mm:ss AM/PM');
  sh.getRange(logRow, 7).setNumberFormat('M/d/yyyy');
  sh.getRange(logRow, 8, 1, 2).setNumberFormat('h:mm AM/PM');

  return {
    success: true,
    message: 'Logbook entry saved successfully.',
    entry: row.map(v => v instanceof Date ? v.toISOString() : v)
  };
}

function getLogbook(filters) {
  let rows = getSheetObjects_(SHEETS.LOGBOOK);

  filters = filters || {};
  const q = clean_(filters.search).toLowerCase();
  const dept = clean_(filters.department);
  const start = clean_(filters.startDate);
  const end = clean_(filters.endDate);

  return rows.filter(r => {
    if (dept && clean_(r['Department']) !== dept) return false;

    const d = normalizeDate_(r['Date']);
    if (start && d < start) return false;
    if (end && d > end) return false;

    if (!q) return true;

    return [
      r['Faculty Name'], r['Department'], r['Room Name'], r['Room Number'],
      r['Equipment Used'], r['Activity Type'], r['Purpose'], r['Remarks']
    ].some(v => String(v || '').toLowerCase().includes(q));
  });
}

function getReports(month) {
  const logs = getSheetObjects_(SHEETS.LOGBOOK);
  const selected = clean_(month);

  const filtered = selected
    ? logs.filter(r => normalizeDate_(r['Date']).startsWith(selected))
    : logs;

  const rooms = {};
  const faculty = {};
  const equipment = {};
  const facultyHours = {};
  const experiments = {};
  let totalHours = 0;

  filtered.forEach(r => {
    increment_(rooms, r['Room Name']);
    increment_(faculty, r['Faculty Name']);

    const equipmentText = clean_(r['Equipment Used']);
    if (equipmentText) increment_(equipment, equipmentText);

    const mins = minutesBetween_(r['Time In'], r['Time Out']);
    if (mins > 0) {
      const hours = mins / 60;
      increment_(facultyHours, r['Faculty Name'], hours);
      totalHours += hours;
    }

    if (clean_(r['Activity Type']).toUpperCase() === 'EXPERIMENT') {
      increment_(experiments, r['Faculty Name']);
    }
  });

  return {
    totalLogs: filtered.length,
    totalHours: Number(totalHours.toFixed(2)),
    rooms: sortObject_(rooms),
    faculty: sortObject_(faculty),
    equipment: sortObject_(equipment),
    facultyHours: sortObject_(facultyHours),
    experiments: sortObject_(experiments)
  };
}

function getSheetObjects_(sheetName) {
  const sh = getRequiredSheet_(sheetName);
  const values = sh.getDataRange().getDisplayValues();

  if (values.length <= 1) return [];

  const headers = values[0];
  return values.slice(1)
    .filter(row => row.some(v => String(v).trim() !== ''))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i] ?? '');
      return obj;
    });
}

function getRequiredSheet_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('Missing sheet: ' + name + '. Run setupSheets() first.');
  return sh;
}

function ensureSpreadsheetTimeZone_(spreadsheet) {
  if (spreadsheet.getSpreadsheetTimeZone() !== APP_TIME_ZONE) {
    spreadsheet.setSpreadsheetTimeZone(APP_TIME_ZONE);
  }
}

function validateLog_(e, now) {
  const required = [
    ['facultyName','Faculty Name'],
    ['department','Department'],
    ['roomNumber','Room Number'],
    ['timeOut','Time Out'],
    ['activityType','Activity Type'],
    ['purpose','Purpose']
  ];

  required.forEach(([key,label]) => {
    if (!e || !clean_(e[key])) throw new Error(label + ' is required.');
  });

  const timeOut = timeOnManilaDate_(now, e.timeOut);
  if (timeOut.getTime() <= now.getTime()) {
    throw new Error('Time Out must be later than the server-recorded Time In.');
  }
  return timeOut;
}

function timeOnManilaDate_(now, value) {
  const time = parseClockTime_(value);
  if (!time) throw new Error('Time Out is required.');
  const date = Utilities.formatDate(now, APP_TIME_ZONE, 'yyyy-MM-dd');
  return new Date(`${date}T${String(time.hour).padStart(2, '0')}:${String(time.minute).padStart(2, '0')}:00+08:00`);
}

function parseClockTime_(value) {
  const match = clean_(value).toUpperCase().match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)?$/);
  if (!match) return null;
  let hour = Number(match[1]);
  const minute = Number(match[2]);
  const period = match[3];
  if (minute > 59 || (period ? hour < 1 || hour > 12 : hour > 23)) return null;
  if (period) {
    if (hour === 12) hour = 0;
    if (period === 'PM') hour += 12;
  }
  return { hour: hour, minute: minute };
}

function seedSettings_() {
  const sh = getRequiredSheet_(SHEETS.SETTINGS);
  if (sh.getLastRow() > 1) return;

  const rows = [
    ['Activity Type','EXPERIMENT','Active'],
    ['Activity Type','CLASS','Active'],
    ['Activity Type','RESEARCH','Active'],
    ['Activity Type','PREPARATION','Active'],
    ['Activity Type','OTHERS','Active']
  ];

  sh.getRange(2,1,rows.length,3).setValues(rows);
}

function clean_(v) {
  return String(v == null ? '' : v).trim();
}

function normalizeDate_(v) {
  const s = clean_(v);
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return `${m[3]}-${String(m[1]).padStart(2,'0')}-${String(m[2]).padStart(2,'0')}`;

  const d = new Date(s);
  if (!isNaN(d)) {
    return Utilities.formatDate(d, APP_TIME_ZONE, 'yyyy-MM-dd');
  }
  return s;
}

function minutesBetween_(a,b) {
  const start = parseClockTime_(a);
  const end = parseClockTime_(b);
  if (!start || !end) return 0;
  return (end.hour * 60 + end.minute) - (start.hour * 60 + start.minute);
}

function increment_(obj, key, amount) {
  key = clean_(key) || 'Unspecified';
  obj[key] = (obj[key] || 0) + (amount == null ? 1 : amount);
}

function sortObject_(obj) {
  return Object.fromEntries(
    Object.entries(obj).sort((a,b) => b[1] - a[1])
  );
}
