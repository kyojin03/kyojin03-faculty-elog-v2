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
      data = getReports({
        month: parameters.month,
        startDate: parameters.startDate,
        endDate: parameters.endDate
      });
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

function getReports(filters) {
  filters = typeof filters === 'string' ? { month: filters } : (filters || {});
  const range = resolveReportRange_(filters);
  const logs = getSheetObjects_(SHEETS.LOGBOOK);
  const included = [];
  let missingDates = 0;
  let invalidDates = 0;

  logs.forEach(row => {
    const rawDate = reportField_(row, ['Date']);
    if (!clean_(rawDate)) {
      missingDates += 1;
      return;
    }
    const date = parseReportDate_(rawDate);
    if (!date) {
      invalidDates += 1;
      return;
    }
    if (range.startDate && date.iso < range.startDate) return;
    if (range.endDate && date.iso > range.endDate) return;
    included.push({ row: row, date: date });
  });

  if (!range.startDate && included.length) {
    range.startDate = included.reduce((minimum, item) => minimum < item.date.iso ? minimum : item.date.iso, included[0].date.iso);
  }
  if (!range.endDate && included.length) {
    range.endDate = included.reduce((maximum, item) => maximum > item.date.iso ? maximum : item.date.iso, included[0].date.iso);
  }
  range.label = reportRangeLabel_(range.startDate, range.endDate);

  const totalLogs = included.length;
  const faculty = {};
  const departments = {};
  const rooms = {};
  const activities = {};
  const equipment = {};
  const facultyHours = {};
  const experiments = {};
  const daily = {};
  const monthly = {};
  const weekdayCounts = [0, 0, 0, 0, 0, 0, 0];
  let missingFaculty = 0;
  let missingDepartments = 0;
  let missingRooms = 0;
  let missingActivities = 0;
  let missingDurationPairs = 0;
  let invalidDurationPairs = 0;
  let usableDurationEntries = 0;
  let totalDurationMinutes = 0;

  included.forEach(item => {
    const row = item.row;
    const facultyLabel = normalizeReportLabel_(reportField_(row, ['Faculty Name']));
    const departmentLabel = normalizeReportLabel_(reportField_(row, ['Department']));
    const roomLabel = reportRoomLabel_(row);
    const activityLabel = normalizeReportLabel_(reportField_(row, ['Activity Type']));
    const equipmentLabel = normalizeReportLabel_(reportField_(row, ['Equipment Used']));

    if (facultyLabel) addReportCount_(faculty, facultyLabel);
    else missingFaculty += 1;

    if (departmentLabel) addReportCount_(departments, departmentLabel);
    else missingDepartments += 1;

    if (roomLabel) addReportCount_(rooms, roomLabel);
    else missingRooms += 1;

    if (activityLabel) addReportCount_(activities, activityLabel);
    else missingActivities += 1;

    if (equipmentLabel) addReportCount_(equipment, equipmentLabel);

    addReportCount_(daily, item.date.iso);
    addReportCount_(monthly, item.date.iso.slice(0, 7));
    weekdayCounts[item.date.weekday] += 1;

    const timeIn = clean_(reportField_(row, ['Time In']));
    const timeOut = clean_(reportField_(row, ['Time Out']));
    if (!timeIn || !timeOut) {
      missingDurationPairs += 1;
    } else {
      const minutes = minutesBetween_(timeIn, timeOut);
      if (minutes > 0) {
        usableDurationEntries += 1;
        totalDurationMinutes += minutes;
        if (facultyLabel) addReportCount_(facultyHours, facultyLabel, minutes / 60);
      } else {
        invalidDurationPairs += 1;
      }
    }

    if (activityLabel === 'Experiment' && facultyLabel) {
      addReportCount_(experiments, facultyLabel);
    }
  });

  const facultyEntries = reportEntries_(faculty, totalLogs);
  const departmentEntries = reportEntries_(departments, totalLogs);
  const roomEntries = reportEntries_(rooms, totalLogs);
  const activityEntries = reportEntries_(activities, totalLogs);
  const equipmentEntries = reportEntries_(equipment, totalLogs);
  const partialMonths = reportPartialMonths_(range.startDate, range.endDate);
  const monthKeys = enumerateReportMonths_(range.startDate, range.endDate, Object.keys(monthly));
  const monthlyEntries = monthKeys.map(month => {
    const monthLabel = reportMonthLabel_(month);
    return {
      month: month,
      label: partialMonths.indexOf(monthLabel) >= 0 ? monthLabel + ' (partial)' : monthLabel,
      count: reportCount_(monthly, month),
      share: reportShare_(reportCount_(monthly, month), totalLogs)
    };
  });
  const calendarWeekdays = reportWeekdayOccurrences_(range.startDate, range.endDate);
  const weekdayNames = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  const weekdays = weekdayNames.map((label, index) => ({
    label: label,
    count: weekdayCounts[index],
    share: reportShare_(weekdayCounts[index], totalLogs),
    calendarOccurrences: calendarWeekdays[index],
    averagePerOccurrence: calendarWeekdays[index]
      ? Number((weekdayCounts[index] / calendarWeekdays[index]).toFixed(2))
      : 0
  }));
  const maximumWeekdayCount = Math.max.apply(null, weekdayCounts.concat([0]));
  const peakWeekdays = maximumWeekdayCount
    ? weekdays.filter(item => item.count === maximumWeekdayCount)
    : [];
  const monthlyPeaks = monthlyEntries.map(month => {
    const dates = Object.keys(daily).filter(date => date.slice(0, 7) === month.month);
    const peakCount = dates.reduce((maximum, date) => Math.max(maximum, reportCount_(daily, date)), 0);
    return {
      month: month.month,
      label: month.label,
      totalLogs: month.count,
      peakCount: peakCount,
      peakDates: dates.filter(date => reportCount_(daily, date) === peakCount && peakCount > 0).sort(),
      shareOfMonth: reportShare_(peakCount, month.count)
    };
  });
  const allDates = Object.keys(daily);
  const overallPeakCount = allDates.reduce((maximum, date) => Math.max(maximum, reportCount_(daily, date)), 0);
  const overallPeakDates = allDates
    .filter(date => reportCount_(daily, date) === overallPeakCount && overallPeakCount > 0)
    .sort();
  const totalHours = Number((totalDurationMinutes / 60).toFixed(2));
  const narratives = buildReportNarratives_({
    totalLogs: totalLogs,
    faculty: facultyEntries,
    departments: departmentEntries,
    rooms: roomEntries,
    activities: activityEntries,
    peakWeekdays: peakWeekdays,
    overallPeakDates: overallPeakDates,
    overallPeakCount: overallPeakCount,
    totalHours: totalHours,
    usableDurationEntries: usableDurationEntries,
    invalidDates: invalidDates,
    missingDates: missingDates,
    partialMonths: partialMonths
  });

  const generatedAt = new Date();
  return {
    range: {
      startDate: range.startDate,
      endDate: range.endDate,
      label: range.label,
      timezone: APP_TIME_ZONE,
      partialMonths: partialMonths
    },
    generatedAt: generatedAt.toISOString(),
    generatedAtDisplay: Utilities.formatDate(generatedAt, APP_TIME_ZONE, 'M/d/yyyy h:mm:ss a'),
    counts: {
      sourceRecords: logs.length,
      includedRecords: totalLogs,
      excludedInvalidDates: invalidDates,
      excludedMissingDates: missingDates
    },
    overview: {
      totalLogs: totalLogs,
      distinctFaculty: facultyEntries.length,
      distinctRooms: roomEntries.length,
      distinctActivities: activityEntries.length,
      totalRecordedHours: totalHours,
      usableDurationEntries: usableDurationEntries
    },
    rankings: {
      faculty: facultyEntries.slice(0, 5),
      departments: departmentEntries.slice(0, 5),
      rooms: roomEntries.slice(0, 5),
      equipment: equipmentEntries.slice(0, 5)
    },
    activities: activityEntries,
    months: monthlyEntries,
    weekdays: weekdays,
    peakWeekdays: peakWeekdays,
    monthlyPeaks: monthlyPeaks,
    overallPeak: {
      dates: overallPeakDates,
      count: overallPeakCount,
      share: reportShare_(overallPeakCount, totalLogs)
    },
    duration: {
      totalHours: totalHours,
      usableEntries: usableDurationEntries,
      missingPairs: missingDurationPairs,
      invalidPairs: invalidDurationPairs,
      averageHours: usableDurationEntries
        ? Number((totalHours / usableDurationEntries).toFixed(2))
        : 0
    },
    dataQuality: {
      missingFaculty: missingFaculty,
      missingDepartments: missingDepartments,
      missingRooms: missingRooms,
      missingActivities: missingActivities,
      missingDates: missingDates,
      invalidDates: invalidDates,
      missingDurationPairs: missingDurationPairs,
      invalidDurationPairs: invalidDurationPairs
    },
    narratives: narratives,
    recommendations: buildReportRecommendations_(narratives, peakWeekdays, roomEntries, invalidDates + missingDates, invalidDurationPairs + missingDurationPairs),
    limitations: [
      'Log counts describe recorded laboratory use; they do not measure capacity, occupancy, adoption, or completion.',
      'Multiple entries on the same date are counted as separate valid records because the Logbook has no reliable record identifier for confirmed-copy deduplication.',
      'The Logbook records Department but does not contain a separate Building field for location analysis.',
      'Purpose and remarks remain free-text fields and are not grouped automatically.',
      'Equipment results use exact normalized labels and do not infer that differently written labels are the same item.'
    ],

    // Backward-compatible fields retained for existing clients.
    totalLogs: totalLogs,
    totalHours: totalHours,
    rooms: reportEntriesObject_(roomEntries),
    faculty: reportEntriesObject_(facultyEntries),
    equipment: reportEntriesObject_(equipmentEntries),
    facultyHours: sortObject_(reportCountObject_(facultyHours)),
    experiments: sortObject_(reportCountObject_(experiments))
  };
}

function resolveReportRange_(filters) {
  const month = clean_(filters.month);
  let startDate = clean_(filters.startDate);
  let endDate = clean_(filters.endDate);

  if (!startDate && !endDate && /^\d{4}-\d{2}$/.test(month)) {
    startDate = month + '-01';
    endDate = month + '-' + String(reportDaysInMonth_(month)).padStart(2, '0');
  }
  if ((startDate && !parseReportDate_(startDate)) || (endDate && !parseReportDate_(endDate))) {
    throw new Error('Report dates must use YYYY-MM-DD.');
  }
  if ((startDate && !endDate) || (!startDate && endDate)) {
    throw new Error('Both report start and end dates are required.');
  }
  if (startDate && endDate && startDate > endDate) {
    throw new Error('Report end date cannot be earlier than the start date.');
  }
  return { startDate: startDate, endDate: endDate, label: '' };
}

function reportField_(row, names) {
  for (let index = 0; index < names.length; index += 1) {
    if (Object.prototype.hasOwnProperty.call(row, names[index])) return row[names[index]];
  }
  return '';
}

function parseReportDate_(value) {
  if (value instanceof Date && !isNaN(value)) {
    return parseReportDate_(Utilities.formatDate(value, APP_TIME_ZONE, 'yyyy-MM-dd'));
  }
  const text = clean_(value);
  let year;
  let month;
  let day;
  let match = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (match) {
    year = Number(match[1]);
    month = Number(match[2]);
    day = Number(match[3]);
  } else {
    match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (!match) return null;
    month = Number(match[1]);
    day = Number(match[2]);
    year = Number(match[3]);
  }
  const utc = new Date(Date.UTC(year, month - 1, day));
  if (utc.getUTCFullYear() !== year || utc.getUTCMonth() !== month - 1 || utc.getUTCDate() !== day) return null;
  return {
    year: year,
    month: month,
    day: day,
    weekday: utc.getUTCDay(),
    iso: `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`
  };
}

function normalizeReportLabel_(value) {
  const text = clean_(value).replace(/\s+/g, ' ');
  if (!text) return '';
  return text.toLowerCase().replace(/(^|[\s\-'])\w/g, part => part.toUpperCase());
}

function reportRoomLabel_(row) {
  const name = normalizeReportLabel_(reportField_(row, ['Room Name']));
  const number = clean_(reportField_(row, ['Room Number'])).replace(/\s+/g, ' ').toUpperCase();
  if (name && number) return `${name} (${number})`;
  return name || number;
}

function addReportCount_(object, label, amount) {
  const key = clean_(label).toUpperCase();
  if (!key) return;
  if (!object[key]) object[key] = { label: label, count: 0 };
  object[key].count += amount == null ? 1 : amount;
}

function reportCount_(object, key) {
  return object[key] ? object[key].count : 0;
}

function reportShare_(count, total) {
  return total ? Number(((count / total) * 100).toFixed(1)) : 0;
}

function reportEntries_(object, total) {
  return Object.keys(object).map(key => ({
    label: object[key].label,
    count: Number(Number(object[key].count).toFixed(2)),
    share: reportShare_(object[key].count, total)
  })).sort((a, b) => b.count - a.count || a.label.localeCompare(b.label));
}

function reportEntriesObject_(entries) {
  const object = {};
  entries.forEach(item => object[item.label] = item.count);
  return object;
}

function reportCountObject_(counts) {
  const object = {};
  Object.keys(counts).forEach(key => {
    object[counts[key].label] = Number(Number(counts[key].count).toFixed(2));
  });
  return object;
}

function enumerateReportMonths_(startDate, endDate, fallbackMonths) {
  if (!startDate || !endDate) return fallbackMonths.slice().sort();
  const start = parseReportDate_(startDate);
  const end = parseReportDate_(endDate);
  const months = [];
  let year = start.year;
  let month = start.month;
  while (year < end.year || (year === end.year && month <= end.month)) {
    months.push(`${year}-${String(month).padStart(2, '0')}`);
    month += 1;
    if (month === 13) {
      month = 1;
      year += 1;
    }
  }
  return months;
}

function reportWeekdayOccurrences_(startDate, endDate) {
  const counts = [0, 0, 0, 0, 0, 0, 0];
  if (!startDate || !endDate) return counts;
  const start = parseReportDate_(startDate);
  const end = parseReportDate_(endDate);
  const cursor = new Date(Date.UTC(start.year, start.month - 1, start.day));
  const last = Date.UTC(end.year, end.month - 1, end.day);
  while (cursor.getTime() <= last) {
    counts[cursor.getUTCDay()] += 1;
    cursor.setUTCDate(cursor.getUTCDate() + 1);
  }
  return counts;
}

function reportDaysInMonth_(month) {
  const match = clean_(month).match(/^(\d{4})-(\d{2})$/);
  if (!match) return 0;
  return new Date(Date.UTC(Number(match[1]), Number(match[2]), 0)).getUTCDate();
}

function reportMonthLabel_(month) {
  const names = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  const match = clean_(month).match(/^(\d{4})-(\d{2})$/);
  return match ? `${names[Number(match[2]) - 1]} ${match[1]}` : month;
}

function reportDateLabel_(date) {
  const parsed = parseReportDate_(date);
  if (!parsed) return clean_(date);
  return `${reportMonthLabel_(parsed.iso.slice(0, 7)).replace(' ' + parsed.year, '')} ${parsed.day}, ${parsed.year}`;
}

function reportRangeLabel_(startDate, endDate) {
  if (!startDate || !endDate) return 'All valid recorded dates';
  if (startDate === endDate) return reportDateLabel_(startDate);
  return `${reportDateLabel_(startDate)} to ${reportDateLabel_(endDate)}`;
}

function reportPartialMonths_(startDate, endDate) {
  if (!startDate || !endDate) return [];
  const start = parseReportDate_(startDate);
  const end = parseReportDate_(endDate);
  const months = [];
  if (start.day !== 1) months.push(reportMonthLabel_(start.iso.slice(0, 7)));
  if (end.day !== reportDaysInMonth_(end.iso.slice(0, 7))) months.push(reportMonthLabel_(end.iso.slice(0, 7)));
  return months.filter((value, index, array) => array.indexOf(value) === index);
}

function buildReportNarratives_(data) {
  const total = data.totalLogs;
  if (!total) {
    return {
      overview: 'No recorded entries fall within the selected period.',
      who: 'No faculty labels were recorded in the selected period.',
      department: 'No department labels were recorded in the selected period.',
      what: 'No activity types were recorded in the selected period.',
      when: 'No peak date or weekday can be identified because the selected period has no recorded entries.',
      where: 'No room use was recorded in the selected period.',
      duration: 'No usable Time In and Time Out pairs were available for duration analysis.',
      coverage: data.invalidDates + data.missingDates
        ? `${data.invalidDates + data.missingDates} source records were excluded because Date was missing or invalid.`
        : 'All source records had usable Date values.'
    };
  }
  const topFaculty = data.faculty[0];
  const topDepartments = data.departments.filter(item => item.count === (data.departments[0] && data.departments[0].count));
  const topRooms = data.rooms.filter(item => item.count === (data.rooms[0] && data.rooms[0].count));
  const topActivities = data.activities.filter(item => item.count === (data.activities[0] && data.activities[0].count));
  const weekdayNames = data.peakWeekdays.map(item => item.label).join(' and ');
  const topDates = data.overallPeakDates.map(reportDateLabel_).join(' and ');
  return {
    overview: `${total} log ${total === 1 ? 'entry was' : 'entries were'} included in the selected period.`,
    who: topFaculty
      ? `${topFaculty.label} recorded ${topFaculty.count} of ${total} logs (${topFaculty.share}%).`
      : 'Faculty labels were missing from all included records.',
    department: topDepartments.length
      ? `${topDepartments.map(item => item.label).join(' and ')} ${topDepartments.length > 1 ? 'were tied as the leading departments' : 'was the leading department'}, with ${topDepartments[0].count} of ${total} logs (${topDepartments[0].share}%).`
      : 'Department was missing from all included records.',
    what: topActivities.length
      ? `${topActivities.map(item => item.label).join(' and ')} ${topActivities.length > 1 ? 'were tied as the leading activities' : 'was the leading activity'}, with ${topActivities[0].count} of ${total} logs (${topActivities[0].share}%).`
      : 'Activity Type was missing from all included records.',
    when: data.peakWeekdays.length
      ? `${weekdayNames} ${data.peakWeekdays.length > 1 ? 'were tied for the highest recorded activity' : 'had the highest recorded activity'}, with ${data.peakWeekdays[0].count} of ${total} logs (${data.peakWeekdays[0].share}%). The overall peak ${data.overallPeakDates.length > 1 ? 'dates were' : 'date was'} ${topDates}, with ${data.overallPeakCount} ${data.overallPeakCount === 1 ? 'log' : 'logs'}.`
      : 'No peak date or weekday could be identified.',
    where: topRooms.length
      ? `${topRooms.map(item => item.label).join(' and ')} ${topRooms.length > 1 ? 'were tied as the most frequently recorded rooms' : 'was the most frequently recorded room'}, with ${topRooms[0].count} of ${total} logs (${topRooms[0].share}%).`
      : 'Room Name and Room Number were missing from all included records.',
    duration: `${data.usableDurationEntries} of ${total} entries had usable Time In and Time Out pairs, totaling ${data.totalHours} recorded hours.`,
    coverage: `${data.invalidDates + data.missingDates} source ${data.invalidDates + data.missingDates === 1 ? 'record was' : 'records were'} excluded because Date was missing or invalid.${data.partialMonths.length ? ' The selected range includes partial coverage for ' + data.partialMonths.join(' and ') + '.' : ''}`
  };
}

function buildReportRecommendations_(narratives, peakWeekdays, rooms, dateIssues, durationIssues) {
  const recommendations = [];
  if (peakWeekdays.length) {
    recommendations.push(`Review staffing, preparation, and equipment readiness for ${peakWeekdays.map(item => item.label).join(' and ')}, the highest-volume recorded weekday${peakWeekdays.length > 1 ? 's' : ''}.`);
  }
  if (rooms.length) {
    recommendations.push(`Prioritize routine readiness checks for ${rooms[0].label}, the most frequently recorded room in this period.`);
  }
  if (dateIssues) recommendations.push('Review records with missing or invalid Date values so future period reports have more complete coverage.');
  if (durationIssues) recommendations.push('Review missing or invalid Time In and Time Out pairs before using recorded hours for operational planning.');
  if (!recommendations.length) recommendations.push('Continue monitoring the same measures across comparable periods before making capacity or occupancy conclusions.');
  return recommendations;
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
