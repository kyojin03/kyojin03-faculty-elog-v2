  const Logbook = (() => {
    let filtersReady = false;

    function initialize() {
      App.byId('logForm').addEventListener('submit', handleSubmit);
      App.byId('roomNumber').addEventListener('change', updateRoomName);
      App.byId('timeOut').addEventListener('input', clearTimeOutError);
      App.byId('clearFilters').addEventListener('click', clearFilters);
      App.byId('refreshLogs').addEventListener('click', () => loadLogs(true));

      const debouncedFilter = App.debounce(() => loadLogs(false), 380);
      App.byId('search').addEventListener('input', debouncedFilter);
      App.byId('deptFilter').addEventListener('change', () => loadLogs(false));
      App.byId('startDate').addEventListener('change', () => loadLogs(false));
      App.byId('endDate').addEventListener('change', () => loadLogs(false));
    }

    function hydrate() {
      populateRooms();
      populateActivityTypes();
      populateFacultyAndDepartments();
      renderLogs(App.state.displayedLogs);
      filtersReady = true;
    }

    function activeRows(rows) {
      return rows.filter(row => !row.Status || String(row.Status).trim().toLowerCase() === 'active');
    }

    function populateRooms() {
      const select = App.byId('roomNumber');
      const options = activeRows(App.state.rooms).map(room => {
        const number = room['Room Number'] || '';
        const name = room['Room Name'] || '';
        const building = room.Building || '';
        const detail = [name, building].filter(Boolean).join(' - ');
        return `<option value="${App.escapeAttribute(number)}" data-room-name="${App.escapeAttribute(name)}">${App.escapeHtml(number)}${detail ? ' | ' + App.escapeHtml(detail) : ''}</option>`;
      });
      select.innerHTML = '<option value="">Select a room</option>' + options.join('');
      updateRoomName();
    }

    function updateRoomName() {
      const selected = App.byId('roomNumber').selectedOptions[0];
      App.byId('roomName').value = selected?.dataset.roomName || '';
    }

    function populateActivityTypes() {
      const values = activeRows(App.state.settings)
        .filter(row => String(row.Category || '').trim().toLowerCase() === 'activity type')
        .map(row => String(row.Value || '').trim())
        .filter(Boolean);
      App.byId('activityType').innerHTML = '<option value="">Select an activity</option>' +
        [...new Set(values)].map(item => `<option value="${App.escapeAttribute(item)}">${App.escapeHtml(item)}</option>`).join('');
    }

    function populateFacultyAndDepartments() {
      const names = [...new Set(App.state.allLogs.map(row => String(row['Faculty Name'] || '').trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b));
      const departments = [...new Set(App.state.allLogs.map(row => String(row.Department || '').trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b));
      App.byId('facultyList').innerHTML = names.map(name => `<option value="${App.escapeAttribute(name)}"></option>`).join('');
      App.byId('departmentList').innerHTML = departments.map(name => `<option value="${App.escapeAttribute(name)}"></option>`).join('');

      const currentFilter = App.byId('deptFilter').value;
      App.byId('deptFilter').innerHTML = '<option value="">All departments</option>' + departments.map(name => `<option value="${App.escapeAttribute(name)}">${App.escapeHtml(name)}</option>`).join('');
      if (departments.includes(currentFilter)) App.byId('deptFilter').value = currentFilter;
    }

    function validateDateFilters() {
      const start = App.value('startDate');
      const end = App.value('endDate');
      const field = App.byId('endDate');
      const error = App.byId('dateFilterError');
      const invalid = Boolean(start && end && end < start);
      const message = invalid ? 'End date cannot be earlier than start date.' : '';
      field.setAttribute('aria-invalid', invalid ? 'true' : 'false');
      error.textContent = message;
      return !invalid;
    }

    function clearTimeOutError() {
      const field = App.byId('timeOut');
      field.setCustomValidity('');
      field.setAttribute('aria-invalid', 'false');
      App.byId('timeError').textContent = '';
    }

    function getFilters() {
      return {
        search: App.value('search'),
        department: App.value('deptFilter'),
        startDate: App.value('startDate'),
        endDate: App.value('endDate')
      };
    }

    function hasFilters(filters) {
      return Object.values(filters).some(Boolean);
    }

    function setLogLoading() {
      App.byId('logTableWrap').innerHTML = '<div class="table-loading" role="status"><div><div class="inline-loader"></div>Loading logbook records...</div></div>';
    }

    async function loadLogs(showLoading) {
      if (!filtersReady || !validateDateFilters()) return;
      const requestId = ++App.state.logRequestId;
      const filters = getFilters();
      if (showLoading) setLogLoading();
      App.byId('refreshLogs').disabled = true;
      try {
        const rows = await App.runServer('getLogbook', filters);
        if (requestId !== App.state.logRequestId) return;
        App.state.displayedLogs = Array.isArray(rows) ? rows : [];
        if (!hasFilters(filters)) {
          App.state.allLogs = App.state.displayedLogs;
          populateFacultyAndDepartments();
        }
        renderLogs(App.state.displayedLogs);
        if (showLoading) App.showToast('Logbook records are up to date.', 'success', 'Refreshed');
      } catch (error) {
        if (requestId !== App.state.logRequestId) return;
        App.logTechnicalError('Logbook load failed.', error);
        renderLogs(App.state.displayedLogs);
        App.showToast('Unable to load logbook records. Please try again.', 'error');
      } finally {
        if (requestId === App.state.logRequestId) App.byId('refreshLogs').disabled = false;
      }
    }

    function clearFilters() {
      App.byId('search').value = '';
      App.byId('deptFilter').value = '';
      App.byId('startDate').value = '';
      App.byId('endDate').value = '';
      validateDateFilters();
      loadLogs(false);
    }

    function formatRecordCount(count) {
      return count + (count === 1 ? ' record' : ' records');
    }

    function formatTimestamp(value) {
      if (value == null) return '';
      const raw = String(value).trim();
      if (!raw) return '';
      if (value instanceof Date && !isNaN(value)) {
        return formatTimestampViaManila(value);
      }
      const slashMatch = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?(?:\s*([AaPp][Mm]))?)?$/);
      if (slashMatch) {
        const month = Number(slashMatch[1]);
        const day = Number(slashMatch[2]);
        const year = Number(slashMatch[3]);
        if (slashMatch[4] !== undefined) {
          let hour = Number(slashMatch[4]);
          const minute = Number(slashMatch[5]);
          const second = slashMatch[6] !== undefined ? Number(slashMatch[6]) : 0;
          const ampmRaw = slashMatch[7];
          if (ampmRaw) {
            const isPM = ampmRaw.toLowerCase() === 'pm';
            if (hour === 12) hour = isPM ? 12 : 0;
            else hour = isPM ? hour + 12 : hour;
          }
          return formatTimestampFromParts(month, day, year, hour, minute, second);
        }
        return `${month}/${day}/${year}`;
      }
      const hasTimezone = /[Zz]$/.test(raw) || /[+-]\d{2}:?\d{2}$/.test(raw) || /\bGMT\b/i.test(raw);
      if (!hasTimezone) {
        const isoLike = raw.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})(?:[T\s](\d{1,2}):(\d{2})(?::(\d{2}))?(?:\.\d+)?)?$/);
        if (isoLike) {
          const year = Number(isoLike[1]);
          const month = Number(isoLike[2]);
          const day = Number(isoLike[3]);
          if (isoLike[4] !== undefined) {
            const hour = Number(isoLike[4]);
            const minute = Number(isoLike[5]);
            const second = isoLike[6] !== undefined ? Number(isoLike[6]) : 0;
            return formatTimestampFromParts(month, day, year, hour, minute, second);
          }
          return `${month}/${day}/${year}`;
        }
      }
      const parsed = new Date(raw);
      if (!isNaN(parsed)) {
        return formatTimestampViaManila(parsed);
      }
      return raw;
    }

    function formatTimestampViaManila(date) {
      const parts = new Intl.DateTimeFormat('en-US', {
        timeZone: 'Asia/Manila',
        year: 'numeric',
        month: 'numeric',
        day: 'numeric',
        hour: 'numeric',
        minute: '2-digit',
        second: '2-digit',
        hour12: true
      }).formatToParts(date);
      const get = type => parts.find(p => p.type === type)?.value || '';
      return `${get('month')}/${get('day')}/${get('year')} ${get('hour')}:${get('minute')}:${get('second')} ${get('dayPeriod').toUpperCase()}`;
    }

    function formatTimestampFromParts(month, day, year, hour24, minute, second) {
      const pad2 = n => String(n).padStart(2, '0');
      const ampm = hour24 >= 12 ? 'PM' : 'AM';
      let hour12 = hour24 % 12;
      if (hour12 === 0) hour12 = 12;
      return `${month}/${day}/${year} ${hour12}:${pad2(minute)}:${pad2(second)} ${ampm}`;
    }

    function renderLogs(rows) {
      const container = App.byId('logTableWrap');
      if (!container.querySelector('table')) container.innerHTML = '<table id="logTable"><thead></thead><tbody></tbody></table>';
      const table = container.querySelector('table');
      const headers = ['Timestamp', 'Name', 'Department', 'Room', 'Room #', 'Equipment', 'Date', 'Time In', 'Time Out', 'Activity', 'Purpose', 'Remarks'];
      table.querySelector('thead').innerHTML = '<tr>' + headers.map(header => `<th scope="col">${App.escapeHtml(header)}</th>`).join('') + '</tr>';

      const safeRows = Array.isArray(rows) ? rows : [];
      App.byId('logCount').textContent = formatRecordCount(safeRows.length);
      if (!safeRows.length) {
        table.querySelector('tbody').innerHTML = '<tr class="empty-row"><td colspan="12"><div class="empty-state"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 3h14v18H5V3Zm2 2v14h10V5H7Zm2 3h6v2H9V8Zm0 4h6v2H9v-2Z"/></svg><strong>No logbook records found</strong><span>Try changing or clearing the current filters.</span></div></td></tr>';
        return;
      }

      table.querySelector('tbody').innerHTML = safeRows.slice().reverse().map(row => `<tr>
        <td>${App.escapeHtml(formatTimestamp(row.Timestamp))}</td>
        <td><strong>${App.escapeHtml(row['Faculty Name'])}</strong></td>
        <td>${App.escapeHtml(row.Department)}</td>
        <td>${App.escapeHtml(row['Room Name'])}</td>
        <td>${App.escapeHtml(row['Room Number'])}</td>
        <td class="wrap-cell">${App.escapeHtml(row['Equipment Used'])}</td>
        <td>${App.escapeHtml(row.Date)}</td>
        <td>${App.escapeHtml(row['Time In'])}</td>
        <td>${App.escapeHtml(row['Time Out'])}</td>
        <td><span class="status-badge">${App.escapeHtml(row['Activity Type'])}</span></td>
        <td class="wrap-cell">${App.escapeHtml(row.Purpose)}</td>
        <td class="wrap-cell">${App.escapeHtml(row.Remarks)}</td>
      </tr>`).join('');
    }

    function uppercase(value) {
      return String(value || '').trim().toUpperCase();
    }

    function buildEntry() {
      return {
        facultyName: uppercase(App.value('facultyName')),
        department: uppercase(App.value('department')),
        roomNumber: App.value('roomNumber'),
        roomName: App.value('roomName'),
        equipmentUsed: uppercase(App.value('equipmentUsed')),
        timeOut: App.value('timeOut'),
        activityType: App.value('activityType'),
        purpose: uppercase(App.value('purpose')),
        remarks: uppercase(App.value('remarks'))
      };
    }

    function setSubmitting(submitting) {
      App.state.submitPending = submitting;
      const button = App.byId('submitButton');
      button.disabled = submitting;
      button.classList.toggle('is-loading', submitting);
      button.querySelector('.button-label').textContent = submitting ? 'Submitting...' : 'Submit log';
      App.byId('logForm').setAttribute('aria-busy', String(submitting));
    }

    async function handleSubmit(event) {
      event.preventDefault();
      if (App.state.submitPending) return;
      const form = App.byId('logForm');
      if (!form.checkValidity()) {
        form.reportValidity();
        return;
      }

      setSubmitting(true);
      try {
        const response = await App.runServer('submitLog', buildEntry());
        form.reset();
        App.byId('roomName').value = '';
        App.showToast(response?.message || 'Your laboratory session was recorded.', 'success', 'Log submitted');
        await loadLogs(false);
      } catch (error) {
        App.logTechnicalError('Logbook submission failed.', error);
        const message = error?.message === 'Time Out must be later than the server-recorded Time In.'
          ? error.message
          : 'Unable to save your log. Please review the details and try again.';
        if (message === error?.message) {
          const field = App.byId('timeOut');
          field.setCustomValidity(message);
          field.setAttribute('aria-invalid', 'true');
          App.byId('timeError').textContent = message;
          field.focus();
        }
        App.showToast(message, 'error');
      } finally {
        setSubmitting(false);
      }
    }

    return { initialize, hydrate, loadLogs };
  })();
