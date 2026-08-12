'use strict';

window.FacultyElogApi = (() => {
  const ACTIONS = Object.freeze({
    getInitialData: { action: 'initialData', method: 'GET' },
    getReadOnlyData: { action: 'readOnlyData', method: 'GET' },
    getLogbook: { action: 'logbook', method: 'GET' },
    submitLog: { action: 'submitLog', method: 'POST' },
    getReports: { action: 'reports', method: 'GET' }
  });

  const config = window.FACULTY_ELOG_CONFIG || {};
  const apiUrl = String(config.apiUrl || '').trim();
  const timeoutMs = Number(config.requestTimeoutMs) > 0 ? Number(config.requestTimeoutMs) : 30000;

  function assertConfiguration() {
    const isAppsScriptWebApp = /^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec$/i.test(apiUrl);
    const isLocalTestApi = /^http:\/\/(127\.0\.0\.1|localhost)(:\d+)?\/exec$/i.test(apiUrl);
    if (!isAppsScriptWebApp && !isLocalTestApi) {
      throw new Error('The Laboratory Log In API URL is not configured.');
    }
  }

  function endpoint(action, parameters = {}) {
    assertConfiguration();
    const url = new URL(apiUrl);
    url.searchParams.set('action', action);
    Object.entries(parameters).forEach(([key, value]) => {
      if (value !== undefined && value !== null && String(value).trim() !== '') {
        url.searchParams.set(key, String(value));
      }
    });
    return url.toString();
  }

  async function request(functionName, ...args) {
    const route = ACTIONS[functionName];
    if (!route) throw new Error('Unsupported Laboratory Log In API operation.');

    const controller = new AbortController();
    const timeout = window.setTimeout(() => controller.abort(), timeoutMs);

    let url = endpoint(route.action);
    const options = {
      method: route.method,
      headers: { Accept: 'application/json' },
      cache: 'no-store',
      redirect: 'follow',
      signal: controller.signal
    };

    if (functionName === 'getLogbook') {
      const filters = args[0] || {};
      url = endpoint(route.action, {
        search: filters.search,
        department: filters.department,
        startDate: filters.startDate,
        endDate: filters.endDate
      });
    } else if (functionName === 'getReports') {
      url = endpoint(route.action, { month: args[0] || '' });
    } else if (functionName === 'submitLog') {
      const entry = args[0] || {};
      options.headers['Content-Type'] = 'application/x-www-form-urlencoded;charset=UTF-8';
      options.body = new URLSearchParams({
        action: route.action,
        facultyName: entry.facultyName || '',
        department: entry.department || '',
        roomNumber: entry.roomNumber || '',
        roomName: entry.roomName || '',
        equipmentUsed: entry.equipmentUsed || '',
        date: entry.date || '',
        timeIn: entry.timeIn || '',
        timeOut: entry.timeOut || '',
        activityType: entry.activityType || '',
        purpose: entry.purpose || '',
        remarks: entry.remarks || ''
      });
    }

    try {
      const response = await fetch(url, options);
      let payload;
      try {
        payload = await response.json();
      } catch (_error) {
        throw new Error('The Laboratory Log In service returned an invalid response.');
      }

      if (!response.ok || payload?.success === false) {
        throw new Error(payload?.message || payload?.error || `The Laboratory Log In service returned ${response.status}.`);
      }

      return Object.prototype.hasOwnProperty.call(payload || {}, 'data') ? payload.data : payload;
    } catch (error) {
      if (error?.name === 'AbortError') throw new Error('The Laboratory Log In request timed out.');
      throw error;
    } finally {
      window.clearTimeout(timeout);
    }
  }

  return { request, endpoint };
})();
