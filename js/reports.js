  const Reports = (() => {
    const MANILA_TIME_ZONE = 'Asia/Manila';
    const MONTH_NAMES = ['January','February','March','April','May','June','July','August','September','October','November','December'];

    function manilaToday() {
      const parts = new Intl.DateTimeFormat('en-CA', {
        timeZone: MANILA_TIME_ZONE,
        year: 'numeric',
        month: '2-digit',
        day: '2-digit'
      }).formatToParts(new Date()).reduce((result, part) => {
        result[part.type] = part.value;
        return result;
      }, {});
      return `${parts.year}-${parts.month}-${parts.day}`;
    }

    function monthFromDate(date) {
      return String(date || '').slice(0, 7);
    }

    function monthBounds(month) {
      const match = String(month || '').match(/^(\d{4})-(\d{2})$/);
      if (!match) return null;
      const year = Number(match[1]);
      const number = Number(match[2]);
      const lastDay = new Date(Date.UTC(year, number, 0)).getUTCDate();
      return { startDate: `${month}-01`, endDate: `${month}-${String(lastDay).padStart(2, '0')}` };
    }

    function shiftedMonth(month, offset) {
      const match = String(month || '').match(/^(\d{4})-(\d{2})$/);
      if (!match) return '';
      const date = new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1 + offset, 1));
      return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, '0')}`;
    }

    function shiftedDate(date, days) {
      const match = String(date || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
      if (!match) return '';
      const value = new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3])));
      value.setUTCDate(value.getUTCDate() + days);
      return `${value.getUTCFullYear()}-${String(value.getUTCMonth() + 1).padStart(2, '0')}-${String(value.getUTCDate()).padStart(2, '0')}`;
    }

    function initialize() {
      const today = manilaToday();
      const month = monthFromDate(today);
      const bounds = monthBounds(month);
      App.byId('reportPreset').value = 'current-month';
      App.byId('reportMonth').value = month;
      App.byId('reportStartDate').value = bounds.startDate;
      App.byId('reportEndDate').value = bounds.endDate;

      App.byId('reportPreset').addEventListener('change', applyPreset);
      App.byId('reportMonth').addEventListener('change', applySelectedMonth);
      App.byId('reportStartDate').addEventListener('change', applyCustomRange);
      App.byId('reportEndDate').addEventListener('change', applyCustomRange);
      App.byId('refreshReports').addEventListener('click', () => load(true));
      App.byId('exportReports').addEventListener('click', exportExcel);
      App.byId('printReports').addEventListener('click', printReport);
    }

    function applyPreset() {
      const preset = App.value('reportPreset');
      const today = manilaToday();
      const currentMonth = monthFromDate(today);
      let range = null;
      if (preset === 'current-month') {
        App.byId('reportMonth').value = currentMonth;
        range = monthBounds(currentMonth);
      } else if (preset === 'previous-month') {
        const previous = shiftedMonth(currentMonth, -1);
        App.byId('reportMonth').value = previous;
        range = monthBounds(previous);
      } else if (preset === 'last-30-days') {
        App.byId('reportMonth').value = '';
        range = { startDate: shiftedDate(today, -29), endDate: today };
      }
      if (!range) return;
      App.byId('reportStartDate').value = range.startDate;
      App.byId('reportEndDate').value = range.endDate;
      load(false);
    }

    function applySelectedMonth() {
      const month = App.value('reportMonth');
      const range = monthBounds(month);
      if (!range) return;
      App.byId('reportPreset').value = 'custom';
      App.byId('reportStartDate').value = range.startDate;
      App.byId('reportEndDate').value = range.endDate;
      load(false);
    }

    function applyCustomRange() {
      App.byId('reportPreset').value = 'custom';
      App.byId('reportMonth').value = exactMonthForRange(App.value('reportStartDate'), App.value('reportEndDate'));
      if (App.value('reportStartDate') && App.value('reportEndDate')) load(false);
    }

    function exactMonthForRange(startDate, endDate) {
      const month = monthFromDate(startDate);
      const bounds = monthBounds(month);
      return bounds && bounds.startDate === startDate && bounds.endDate === endDate ? month : '';
    }

    function currentRange() {
      const startDate = App.value('reportStartDate');
      const endDate = App.value('reportEndDate');
      if (!startDate || !endDate) {
        showError('Choose both a start date and an end date.');
        return null;
      }
      if (endDate < startDate) {
        showError('End date cannot be earlier than start date.');
        return null;
      }
      clearError();
      return { startDate, endDate };
    }

    function rangeKey(range) {
      return range ? `${range.startDate}|${range.endDate}` : '';
    }

    function ensureLoaded() {
      const range = currentRange();
      if (!range) return;
      if (!App.state.reportData || App.state.reportsRangeKey !== rangeKey(range)) load(false);
    }

    function setActionsDisabled(disabled) {
      App.byId('refreshReports').disabled = disabled;
      App.byId('exportReports').disabled = disabled || !App.state.reportData;
      App.byId('printReports').disabled = disabled || !App.state.reportData;
    }

    function showLoading() {
      ['statLogs', 'statFaculty', 'statRooms', 'statHours'].forEach(id => App.byId(id).textContent = '--');
      App.byId('statHoursDetail').textContent = 'Loading duration coverage';
      App.byId('reportPeriodLabel').textContent = 'Loading...';
      App.byId('reportRecordCount').textContent = '--';
      App.byId('reportLastRefresh').textContent = '--';
      ['reportFaculty','reportDepartments','reportRooms','reportMonthly','reportWeekdays','reportActivities','reportEquipment'].forEach(id => {
        App.byId(id).innerHTML = '<div class="report-empty" role="status"><div><div class="inline-loader"></div>Loading report...</div></div>';
      });
      App.byId('reportExecutiveBody').innerHTML = '<tr><td colspan="3">Loading report analysis...</td></tr>';
      App.byId('reportPeakDates').innerHTML = '<tr><td colspan="5">Loading peak dates...</td></tr>';
    }

    async function load(showConfirmation) {
      const range = currentRange();
      if (!range) return;
      const requestId = ++App.state.reportRequestId;
      showLoading();
      setActionsDisabled(true);
      try {
        const report = await App.runServer('getReports', range);
        if (requestId !== App.state.reportRequestId) return;
        App.state.reportData = report && typeof report === 'object' ? report : {};
        App.state.reportsRangeKey = rangeKey(range);
        render(App.state.reportData);
        if (showConfirmation) App.showToast('Report data is up to date.', 'success', 'Refreshed');
      } catch (error) {
        if (requestId !== App.state.reportRequestId) return;
        App.logTechnicalError('Report load failed.', error);
        App.state.reportData = null;
        renderEmpty();
        showError('Unable to generate the report. The previous report was not reused. Please try again.');
      } finally {
        if (requestId === App.state.reportRequestId) setActionsDisabled(false);
      }
    }

    function numericValue(value) {
      const numberValue = Number(value);
      return Number.isFinite(numberValue) ? numberValue : 0;
    }

    function number(value, maximumFractionDigits = 0) {
      return numericValue(value).toLocaleString(undefined, { maximumFractionDigits });
    }

    function percent(value) {
      return `${numericValue(value).toLocaleString(undefined, { minimumFractionDigits: 1, maximumFractionDigits: 1 })}%`;
    }

    function dateLabel(iso) {
      const match = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
      if (!match) return String(iso || '');
      return `${MONTH_NAMES[Number(match[2]) - 1]} ${Number(match[3])}, ${match[1]}`;
    }

    function showError(message) {
      const error = App.byId('reportError');
      error.textContent = message;
      error.hidden = false;
    }

    function clearError() {
      const error = App.byId('reportError');
      error.textContent = '';
      error.hidden = true;
    }

    function safeArray(value) {
      return Array.isArray(value) ? value : [];
    }

    function render(report) {
      clearError();
      const overview = report.overview || {};
      const counts = report.counts || {};
      const range = report.range || {};
      const duration = report.duration || {};
      const rankings = report.rankings || {};
      const narratives = report.narratives || {};
      const total = numericValue(overview.totalLogs);

      App.byId('reportPeriodLabel').textContent = range.label || 'Selected range';
      App.byId('reportRecordCount').textContent = `${number(counts.includedRecords)} included`;
      App.byId('reportLastRefresh').textContent = report.generatedAtDisplay || 'Just now';
      App.byId('statLogs').textContent = number(total);
      App.byId('statFaculty').textContent = number(overview.distinctFaculty);
      App.byId('statRooms').textContent = number(overview.distinctRooms);
      App.byId('statHours').textContent = number(overview.totalRecordedHours, 2);
      App.byId('statHoursDetail').textContent = `${number(duration.usableEntries)} of ${number(total)} usable duration pairs`;
      App.byId('printCoverPeriod').textContent = range.label || 'Selected reporting period';
      App.byId('printCoverMeta').textContent = `Generated ${report.generatedAtDisplay || ''} | ${number(total)} included records | Asia/Manila`;
      App.byId('printClosingPeriod').textContent = range.label || 'Selected reporting period';

      renderRanking('reportFaculty', rankings.faculty, { limit: 5 });
      renderRanking('reportDepartments', rankings.departments, { limit: 5 });
      renderRanking('reportRooms', rankings.rooms, { limit: 5 });
      renderRanking('reportMonthly', report.months, { limit: 24 });
      renderRanking('reportWeekdays', report.weekdays, {
        limit: 7,
        detail: item => `${number(item.count)} logs | ${percent(item.share)} | avg ${number(item.averagePerOccurrence, 2)} across ${number(item.calendarOccurrences)} ${String(item.label || '').toLowerCase()}s`
      });
      renderRanking('reportActivities', report.activities, { limit: 8 });
      renderRanking('reportEquipment', rankings.equipment, { limit: 5 });

      App.byId('facultyAnalysis').textContent = [narratives.who, narratives.department].filter(Boolean).join(' ');
      App.byId('roomAnalysis').textContent = narratives.where || '';
      App.byId('whenAnalysis').textContent = narratives.when || '';
      App.byId('peakAnalysis').textContent = overallPeakSentence(report.overallPeak, total);
      App.byId('activityAnalysis').textContent = narratives.what || '';
      App.byId('durationAnalysis').textContent = narratives.duration || '';
      App.byId('reportOverviewNarrative').textContent = narratives.overview || '';

      renderExecutive(report);
      renderPeakDates(report.monthlyPeaks);
      renderQuality(report);
      renderLists(report);
    }

    function renderEmpty() {
      render({
        range: { label: 'Selected range' }, counts: {}, overview: {}, duration: {}, rankings: {},
        activities: [], months: [], weekdays: [], monthlyPeaks: [], narratives: { overview: 'No report is currently available.' },
        recommendations: [], limitations: []
      });
    }

    function renderRanking(id, entries, options = {}) {
      const container = App.byId(id);
      const rows = safeArray(entries);
      if (!rows.length) {
        container.innerHTML = '<div class="report-empty">No recorded entries for this measure.</div>';
        return;
      }
      const limit = options.limit || rows.length;
      const visible = rows.slice(0, limit);
      const maximum = Math.max(...visible.map(item => numericValue(item.count)), 1);
      container.innerHTML = visible.map(item => {
        const width = Math.max(item.count ? 3 : 0, Math.round((numericValue(item.count) / maximum) * 100));
        const detail = options.detail
          ? options.detail(item)
          : `${number(item.count)} ${numericValue(item.count) === 1 ? 'log' : 'logs'} | ${percent(item.share)}`;
        return `<div class="ranking-row" title="${App.escapeAttribute(item.label + ': ' + detail)}">
          <span class="ranking-label">${App.escapeHtml(item.label)}</span>
          <span class="ranking-bar" aria-hidden="true"><span style="width:${width}%"></span></span>
          <strong class="ranking-value">${App.escapeHtml(detail)}</strong>
        </div>`;
      }).join('') + (rows.length > visible.length ? `<div class="ranking-more">${rows.length - visible.length} additional categories are not shown in this concise view.</div>` : '');
    }

    function leadingEvidence(entries) {
      const first = safeArray(entries)[0];
      const total = App.state.reportData?.overview?.totalLogs;
      return first
        ? { result: first.label, evidence: `${number(first.count)} of ${number(total)} logs (${percent(first.share)})` }
        : { result: 'No recorded entries', evidence: '0 included records' };
    }

    function renderExecutive(report) {
      const rankings = report.rankings || {};
      const faculty = leadingEvidence(rankings.faculty);
      const department = leadingEvidence(rankings.departments);
      const activity = leadingEvidence(report.activities);
      const room = leadingEvidence(rankings.rooms);
      const weekday = leadingEvidence(report.peakWeekdays);
      const peak = report.overallPeak || {};
      const total = numericValue(report.overview?.totalLogs);
      const rows = [
        ['WHO', faculty.result, faculty.evidence],
        ['DEPARTMENT', department.result, department.evidence],
        ['WHAT', activity.result, activity.evidence],
        ['WHEN', weekday.result, weekday.evidence],
        ['PEAK DATE', safeArray(peak.dates).length ? safeArray(peak.dates).map(dateLabel).join(' and ') : 'No recorded entries', safeArray(peak.dates).length ? `${number(peak.count)} of ${number(total)} logs (${percent(peak.share)})` : '0 included records'],
        ['WHERE', room.result, room.evidence],
        ['DURATION', `${number(report.duration?.totalHours, 2)} recorded hours`, `${number(report.duration?.usableEntries)} of ${number(total)} entries had usable time pairs`]
      ];
      App.byId('reportExecutiveBody').innerHTML = rows.map(row => `<tr><th scope="row">${App.escapeHtml(row[0])}</th><td>${App.escapeHtml(row[1])}</td><td>${App.escapeHtml(row[2])}</td></tr>`).join('');
    }

    function overallPeakSentence(peak, total) {
      const dates = safeArray(peak?.dates);
      if (!dates.length) return 'No peak date can be identified because the selected period has no recorded entries.';
      return `${dates.length > 1 ? 'The overall peak dates were' : 'The overall peak date was'} ${dates.map(dateLabel).join(' and ')}, with ${number(peak.count)} of ${number(total)} logs (${percent(peak.share)}).`;
    }

    function renderPeakDates(peaks) {
      const rows = safeArray(peaks);
      App.byId('reportPeakDates').innerHTML = rows.length ? rows.map(item => `<tr>
        <th scope="row">${App.escapeHtml(item.label)}</th><td>${number(item.totalLogs)}</td>
        <td>${item.peakDates?.length ? item.peakDates.map(dateLabel).map(App.escapeHtml).join('<br>') : 'No recorded entries'}</td>
        <td>${number(item.peakCount)}</td><td>${percent(item.shareOfMonth)}</td>
      </tr>`).join('') : '<tr><td colspan="5">No months fall within the selected period.</td></tr>';
    }

    function qualityItem(label, value, note) {
      return `<div class="quality-item"><span>${App.escapeHtml(label)}</span><strong>${App.escapeHtml(String(value))}</strong><small>${App.escapeHtml(note)}</small></div>`;
    }

    function renderQuality(report) {
      const duration = report.duration || {};
      const quality = report.dataQuality || {};
      const counts = report.counts || {};
      App.byId('reportDuration').innerHTML = [
        qualityItem('Usable duration pairs', number(duration.usableEntries), 'Valid Time In and Time Out pairs'),
        qualityItem('Recorded hours', number(duration.totalHours, 2), 'Sum of valid recorded durations'),
        qualityItem('Average recorded hours', number(duration.averageHours, 2), 'Per entry with a usable duration pair'),
        qualityItem('Missing time pairs', number(duration.missingPairs), 'One or both time fields missing'),
        qualityItem('Invalid time pairs', number(duration.invalidPairs), 'Unparseable or non-positive duration')
      ].join('');
      App.byId('reportDataQuality').innerHTML = [
        qualityItem('Source records', number(counts.sourceRecords), 'Nonblank Logbook rows read'),
        qualityItem('Included records', number(counts.includedRecords), 'Valid Date inside both boundaries'),
        qualityItem('Missing Date', number(counts.excludedMissingDates), 'Excluded from date-based reporting'),
        qualityItem('Invalid Date', number(counts.excludedInvalidDates), 'Excluded without silently repairing'),
        qualityItem('Missing faculty / department', `${number(quality.missingFaculty)} / ${number(quality.missingDepartments)}`, 'Included but flagged for grouping'),
        qualityItem('Missing room / activity', `${number(quality.missingRooms)} / ${number(quality.missingActivities)}`, 'Included but flagged for grouping')
      ].join('');
    }

    function listHtml(items, emptyText) {
      const values = safeArray(items).filter(Boolean);
      return values.length ? values.map(item => `<li>${App.escapeHtml(item)}</li>`).join('') : `<li>${App.escapeHtml(emptyText)}</li>`;
    }

    function renderLists(report) {
      const narratives = report.narratives || {};
      const findings = [narratives.overview,narratives.who,narratives.department,narratives.what,narratives.when,narratives.where,narratives.duration,narratives.coverage];
      App.byId('reportFindings').innerHTML = listHtml(findings, 'No findings are available for this period.');
      App.byId('reportRecommendations').innerHTML = listHtml(report.recommendations, 'Continue collecting consistent records before making operational conclusions.');
      App.byId('reportLimitations').innerHTML = listHtml(report.limitations, 'No additional limitations were supplied.');
    }

    async function printReport() {
      const range = currentRange();
      if (!range) return;
      if (!App.state.reportData || App.state.reportsRangeKey !== rangeKey(range)) await load(false);
      if (App.state.reportData) window.print();
    }

    function xmlEscape(value) {
      return String(value == null ? '' : value).replace(/[&<>"']/g, character => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&apos;' })[character]);
    }

    function xmlCell(value, style) {
      const isNumber = typeof value === 'number' && Number.isFinite(value);
      return `<Cell${style ? ` ss:StyleID="${style}"` : ''}><Data ss:Type="${isNumber ? 'Number' : 'String'}">${xmlEscape(value)}</Data></Cell>`;
    }

    function xmlRow(values, style) {
      return `<Row>${values.map(value => xmlCell(value, style)).join('')}</Row>`;
    }

    function xmlSheet(name, rows) {
      return `<Worksheet ss:Name="${xmlEscape(name.slice(0, 31))}"><Table>${rows.join('')}</Table><WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel"><FreezePanes/><FrozenNoSplit/><SplitHorizontal>1</SplitHorizontal><TopRowBottomPane>1</TopRowBottomPane></WorksheetOptions></Worksheet>`;
    }

    function buildExcelXml(report) {
      const sheets = [];
      const summary = [xmlRow(['Faculty eLog Utilization Report'], 'Title'),xmlRow(['Period',report.range?.label || '']),xmlRow(['Generated',report.generatedAtDisplay || '']),xmlRow(['Metric','Value'], 'Header')];
      [['Included logs',report.overview?.totalLogs],['Distinct faculty labels',report.overview?.distinctFaculty],['Distinct rooms',report.overview?.distinctRooms],['Recorded hours',report.duration?.totalHours],['Usable duration entries',report.duration?.usableEntries],['Missing dates',report.counts?.excludedMissingDates],['Invalid dates',report.counts?.excludedInvalidDates]].forEach(row => summary.push(xmlRow(row)));
      summary.push(xmlRow([]),xmlRow(['Key findings'], 'Header'));
      const narratives = report.narratives || {};
      [narratives.overview,narratives.who,narratives.department,narratives.what,narratives.when,narratives.where,narratives.duration,narratives.coverage].filter(Boolean).forEach(item => summary.push(xmlRow([item])));
      summary.push(xmlRow([]),xmlRow(['Recommendations'], 'Header'));
      safeArray(report.recommendations).forEach(item => summary.push(xmlRow([item])));
      sheets.push(xmlSheet('Executive Summary', summary));

      const rows = (header, entries, mapper) => [xmlRow(header, 'Header')].concat(safeArray(entries).map(item => xmlRow(mapper(item))));
      sheets.push(xmlSheet('Faculty', rows(['Faculty label','Logs','Share percent'],report.rankings?.faculty,item => [item.label,item.count,item.share])));
      sheets.push(xmlSheet('Departments', rows(['Department','Logs','Share percent'],report.rankings?.departments,item => [item.label,item.count,item.share])));
      sheets.push(xmlSheet('Rooms', rows(['Room','Logs','Share percent'],report.rankings?.rooms,item => [item.label,item.count,item.share])));
      sheets.push(xmlSheet('Activities', rows(['Activity type','Logs','Share percent'],report.activities,item => [item.label,item.count,item.share])));
      sheets.push(xmlSheet('Monthly Trend', rows(['Month','Logs','Share percent'],report.months,item => [item.label,item.count,item.share])));
      sheets.push(xmlSheet('Weekdays', rows(['Weekday','Logs','Share percent','Calendar occurrences','Average per occurrence'],report.weekdays,item => [item.label,item.count,item.share,item.calendarOccurrences,item.averagePerOccurrence])));
      sheets.push(xmlSheet('Peak Dates', rows(['Month','Monthly logs','Peak dates','Peak logs','Share of month'],report.monthlyPeaks,item => [item.label,item.totalLogs,safeArray(item.peakDates).map(dateLabel).join('; '),item.peakCount,item.shareOfMonth])));
      sheets.push(xmlSheet('Data Quality', [xmlRow(['Measure','Count'], 'Header'),xmlRow(['Source records',report.counts?.sourceRecords || 0]),xmlRow(['Included records',report.counts?.includedRecords || 0]),xmlRow(['Missing Date',report.counts?.excludedMissingDates || 0]),xmlRow(['Invalid Date',report.counts?.excludedInvalidDates || 0]),xmlRow(['Usable duration pairs',report.duration?.usableEntries || 0]),xmlRow(['Missing time pairs',report.duration?.missingPairs || 0]),xmlRow(['Invalid time pairs',report.duration?.invalidPairs || 0])]));

      return `<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"><Styles><Style ss:ID="Default" ss:Name="Normal"><Alignment ss:Vertical="Bottom"/><Font ss:FontName="Arial" ss:Size="10"/></Style><Style ss:ID="Title"><Font ss:FontName="Arial" ss:Size="16" ss:Bold="1" ss:Color="#0B2447"/></Style><Style ss:ID="Header"><Font ss:FontName="Arial" ss:Bold="1" ss:Color="#FFFFFF"/><Interior ss:Color="#0B2447" ss:Pattern="Solid"/></Style></Styles>${sheets.join('')}</Workbook>`;
    }

    function exportExcel() {
      const report = App.state.reportData;
      if (!report) {
        App.showToast('Generate the report before exporting it.', 'error');
        return;
      }
      const xml = buildExcelXml(report);
      const blob = new Blob(['\ufeff', xml], { type: 'application/vnd.ms-excel;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      const start = report.range?.startDate || 'all';
      const end = report.range?.endDate || 'dates';
      link.href = url;
      link.download = `faculty-elog-utilization-${start}-to-${end}.xml`;
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.setTimeout(() => URL.revokeObjectURL(url), 1000);
      App.showToast('Excel-compatible report exported.', 'success', 'Export ready');
    }

    return { initialize, ensureLoaded, load };
  })();
