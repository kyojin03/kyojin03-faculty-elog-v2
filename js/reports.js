  const Reports = (() => {
    function currentMonth() {
      const now = new Date();
      return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
    }

    function initialize() {
      App.byId('reportMonth').value = currentMonth();
      App.byId('reportMonth').addEventListener('change', () => load(false));
      App.byId('refreshReports').addEventListener('click', () => load(true));
      App.byId('printReports').addEventListener('click', () => window.print());
    }

    function ensureLoaded() {
      const month = App.value('reportMonth');
      if (!App.state.reportData || App.state.reportsMonth !== month) load(false);
    }

    function showLoading() {
      ['statLogs', 'statHours', 'statExperiments'].forEach(id => App.byId(id).textContent = '--');
      ['reportRooms', 'reportFaculty', 'reportEquipment', 'reportHours', 'reportExperiments'].forEach(id => {
        App.byId(id).innerHTML = '<div class="report-empty" role="status"><div><div class="inline-loader"></div>Loading report...</div></div>';
      });
    }

    async function load(showConfirmation) {
      const requestId = ++App.state.reportRequestId;
      const month = App.value('reportMonth');
      showLoading();
      App.byId('refreshReports').disabled = true;
      try {
        const report = await App.runServer('getReports', month);
        if (requestId !== App.state.reportRequestId) return;
        App.state.reportData = report && typeof report === 'object' ? report : {};
        App.state.reportsMonth = month;
        render(App.state.reportData);
        if (showConfirmation) App.showToast('Report data is up to date.', 'success', 'Refreshed');
      } catch (error) {
        if (requestId !== App.state.reportRequestId) return;
        App.logTechnicalError('Report load failed.', error);
        render({});
        App.showToast('Unable to load reports. Please try again.', 'error');
      } finally {
        if (requestId === App.state.reportRequestId) App.byId('refreshReports').disabled = false;
      }
    }

    function numericValue(value) {
      const number = Number(value);
      return Number.isFinite(number) ? number : 0;
    }

    function render(report) {
      const experiments = report.experiments && typeof report.experiments === 'object' ? report.experiments : {};
      const experimentTotal = Object.values(experiments).reduce((sum, value) => sum + numericValue(value), 0);
      App.byId('statLogs').textContent = numericValue(report.totalLogs).toLocaleString();
      App.byId('statHours').textContent = numericValue(report.totalHours).toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 2 });
      App.byId('statExperiments').textContent = experimentTotal.toLocaleString();
      renderRanking('reportRooms', report.rooms, 'sessions');
      renderRanking('reportFaculty', report.faculty, 'sessions');
      renderRanking('reportEquipment', report.equipment, 'uses');
      renderRanking('reportHours', report.facultyHours, 'hours');
      renderRanking('reportExperiments', experiments, 'experiments');
    }

    function renderRanking(id, object, unit) {
      const container = App.byId(id);
      const entries = Object.entries(object && typeof object === 'object' ? object : {})
        .map(([label, value]) => [label, numericValue(value)])
        .sort((a, b) => b[1] - a[1]);

      if (!entries.length) {
        container.innerHTML = '<div class="report-empty">No report data is available for this month.</div>';
        return;
      }

      const maximum = Math.max(...entries.map(entry => entry[1]), 1);
      container.innerHTML = entries.map(([label, value]) => {
        const width = Math.max(3, Math.round((value / maximum) * 100));
        const formatted = unit === 'hours'
          ? value.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 2 }) + ' h'
          : value.toLocaleString();
        return `<div class="ranking-row" title="${App.escapeAttribute(label)}: ${App.escapeAttribute(formatted)}">
          <span class="ranking-label">${App.escapeHtml(label)}</span>
          <span class="ranking-bar" aria-hidden="true"><span style="width:${width}%"></span></span>
          <strong class="ranking-value" aria-label="${App.escapeAttribute(formatted + ' ' + unit)}">${App.escapeHtml(formatted)}</strong>
        </div>`;
      }).join('');
    }

    return { initialize, ensureLoaded, load };
  })();
