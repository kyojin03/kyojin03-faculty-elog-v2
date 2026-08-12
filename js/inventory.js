  const Inventory = (() => {
    const LOCK_ICON = '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M7 10V7a5 5 0 0 1 10 0v3h2v12H5V10h2Zm2 0h6V7a3 3 0 0 0-6 0v3Zm-2 2v8h10v-8H7Zm4 2h2v4h-2v-4Z"/></svg>';
    const SEARCH_ICON = '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M10 4a6 6 0 1 0 3.5 10.9l4.8 4.8 1.4-1.4-4.8-4.8A6 6 0 0 0 10 4Zm0 2a4 4 0 1 1 0 8 4 4 0 0 1 0-8Z"/></svg>';
    const REFRESH_ICON = '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M17.7 6.3A8 8 0 1 0 20 12h-2a6 6 0 1 1-1.8-4.3L13 11h7V4l-2.3 2.3Z"/></svg>';

    const modules = {
      consumables: {
        dataKey: 'consumables',
        title: 'Laboratory Supplies',
        subtitle: 'Consumable inventory maintained by the laboratory office.',
        columns: ['Item Name', 'Quantity', 'Unit', 'Availability', 'Expiration Date', 'Remarks'],
        statusColumns: ['Availability'],
        searchPlaceholder: 'Search laboratory supplies'
      },
      glassware: {
        dataKey: 'glassware',
        title: 'Glassware',
        subtitle: 'Current glassware inventory from the official spreadsheet.',
        columns: ['Glassware Name', 'Quantity', 'Unit', 'Remarks'],
        statusColumns: [],
        searchPlaceholder: 'Search glassware'
      },
      equipment: {
        dataKey: 'equipment',
        title: 'Equipment',
        subtitle: 'General laboratory equipment and its recorded condition.',
        columns: ['Equipment Name', 'Quantity', 'Unit', 'Condition', 'Remarks'],
        statusColumns: ['Condition'],
        searchPlaceholder: 'Search equipment'
      },
      specialized: {
        dataKey: 'specializedEquipment',
        title: 'Specialized Equipment',
        subtitle: 'Specialized laboratory assets, locations, and conditions.',
        columns: ['Equipment Name', 'Quantity', 'Unit', 'Location', 'Condition', 'Remarks'],
        statusColumns: ['Condition'],
        emphasizeColumns: ['Location'],
        searchPlaceholder: 'Search specialized equipment'
      },
      forms: {
        dataKey: 'laboratoryForms',
        title: 'Laboratory Forms',
        subtitle: 'Official forms and resources maintained by the laboratory office.',
        columns: ['Form Name', 'Description', 'Status', 'File / Link'],
        statusColumns: ['Status'],
        linkColumn: 'File / Link',
        searchPlaceholder: 'Search laboratory forms'
      }
    };

    function initialize() {
      Object.entries(modules).forEach(([id, config]) => {
        renderShell(id, config);
        App.byId(id + 'Search').addEventListener('input', App.debounce(() => renderTable(id), 180));
        App.byId(id + 'Refresh').addEventListener('click', () => refresh(id));
      });
    }

    function renderShell(id, config) {
      App.byId('view-' + id).innerHTML = `
        <div class="page-heading inventory-heading">
          <div class="heading-copy">
            <div class="heading-icon"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 3h14v18H5V3Zm2 2v14h10V5H7Zm2 3h6v2H9V8Zm0 4h6v2H9v-2Z"/></svg></div>
            <div><span class="section-kicker">Laboratory reference</span><h2 id="${id}Heading">${App.escapeHtml(config.title)}</h2><p>${App.escapeHtml(config.subtitle)}</p><span class="read-only-pill">${LOCK_ICON} Read-only</span></div>
          </div>
          <div class="inventory-toolbar">
            <div class="inventory-search">${SEARCH_ICON}<label class="sr-only" for="${id}Search">${App.escapeHtml(config.searchPlaceholder)}</label><input id="${id}Search" type="search" placeholder="${App.escapeAttribute(config.searchPlaceholder)}" autocomplete="off"></div>
            <button class="button secondary-button" id="${id}Refresh" type="button">${REFRESH_ICON}Refresh</button>
          </div>
        </div>
        <div class="card table-card">
          <div class="card-heading table-heading"><div><h3>${App.escapeHtml(config.title)} directory</h3><p>Information shown here can only be updated in Google Sheets.</p></div><span class="count-badge" id="${id}Count">0 items</span></div>
          <div class="table-wrap" id="${id}Table" tabindex="0" aria-label="Scrollable ${App.escapeAttribute(config.title)} table"><div class="table-loading"><div><div class="inline-loader"></div>Loading inventory...</div></div></div>
        </div>`;
    }

    function hydrate() {
      Object.keys(modules).forEach(renderTable);
    }

    function getRows(id) {
      const config = modules[id];
      const rows = App.state.data?.[config.dataKey];
      return Array.isArray(rows) ? rows : [];
    }

    function filteredRows(id) {
      const config = modules[id];
      const query = (App.byId(id + 'Search')?.value || '').trim().toLowerCase();
      const rows = getRows(id);
      if (!query) return rows;
      return rows.filter(row => config.columns.some(column => String(row[column] || '').toLowerCase().includes(query)));
    }

    function badgeClass(value) {
      const normalized = String(value || '').trim().toLowerCase();
      if (/(available|good|excellent|working|operational|active|complete|ready)/.test(normalized)) return 'good';
      if (/(unavailable|damaged|broken|defective|inactive|out of stock|not available)/.test(normalized)) return 'bad';
      if (/(limited|repair|maintenance|fair|low|pending)/.test(normalized)) return 'warn';
      return '';
    }

    function isValidWebUrl(value) {
      try {
        const url = new URL(String(value || '').trim());
        return url.protocol === 'http:' || url.protocol === 'https:';
      } catch (_error) {
        return false;
      }
    }

    function renderCell(config, column, rawValue) {
      const value = String(rawValue == null ? '' : rawValue);
      if (column === 'Expiration Date' && !value.trim()) return '<td>&mdash;</td>';
      if (config.linkColumn === column) {
        if (!isValidWebUrl(value)) return `<td>${value ? App.escapeHtml(value) : '<span class="muted">Not available</span>'}</td>`;
        return `<td><a class="link-button" href="${App.escapeAttribute(value)}" target="_blank" rel="noopener noreferrer">Open form<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M14 3h7v7h-2V6.4l-8.3 8.3-1.4-1.4L17.6 5H14V3ZM5 5h6v2H7v10h10v-4h2v6H5V5Z"/></svg></a></td>`;
      }
      if ((config.statusColumns || []).includes(column)) return `<td><span class="status-badge ${badgeClass(value)}">${App.escapeHtml(value || 'Not specified')}</span></td>`;
      if ((config.emphasizeColumns || []).includes(column)) return `<td><strong>${App.escapeHtml(value)}</strong></td>`;
      const wrap = column === 'Remarks' || column === 'Description' ? ' class="wrap-cell"' : '';
      return `<td${wrap}>${App.escapeHtml(value)}</td>`;
    }

    function renderTable(id) {
      const config = modules[id];
      const rows = filteredRows(id);
      App.byId(id + 'Count').textContent = rows.length + (rows.length === 1 ? ' item' : ' items');
      const container = App.byId(id + 'Table');
      const heading = `<thead><tr>${config.columns.map(column => `<th scope="col">${App.escapeHtml(column)}</th>`).join('')}</tr></thead>`;
      const body = rows.length
        ? `<tbody>${rows.map(row => `<tr>${config.columns.map(column => renderCell(config, column, row[column])).join('')}</tr>`).join('')}</tbody>`
        : `<tbody><tr class="empty-row"><td colspan="${config.columns.length}"><div class="empty-state"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 3h14v18H5V3Zm2 2v14h10V5H7Zm2 3h6v2H9V8Zm0 4h6v2H9v-2Z"/></svg><strong>No ${App.escapeHtml(config.title.toLowerCase())} found</strong><span>${App.value(id + 'Search') ? 'Try a different search term.' : 'No information has been added to the spreadsheet yet.'}</span></div></td></tr></tbody>`;
      container.innerHTML = `<table>${heading}${body}</table>`;
    }

    async function refresh(id, showConfirmation = true) {
      const button = App.byId(id + 'Refresh');
      button.disabled = true;
      button.innerHTML = '<span class="inline-loader" style="width:14px;height:14px;margin:0;border-width:2px"></span>Refreshing';
      try {
        const data = await App.runServer('getReadOnlyData');
        const normalized = App.normalizeData(data);
        Object.keys(modules).forEach(moduleId => {
          const key = modules[moduleId].dataKey;
          App.state.data[key] = normalized[key];
        });
        App.state.rooms = normalized.rooms;
        App.state.settings = normalized.settings;
        hydrate();
        if (showConfirmation) App.showToast(`${modules[id].title} data is up to date.`, 'success', 'Refreshed');
      } catch (error) {
        App.logTechnicalError(`${modules[id].title} refresh failed.`, error);
        App.showToast('Unable to load laboratory inventory. Please try again.', 'error');
      } finally {
        button.disabled = false;
        button.innerHTML = REFRESH_ICON + 'Refresh';
      }
    }

    function hasModule(id) {
      return Object.prototype.hasOwnProperty.call(modules, id);
    }

    function refreshOnOpen(id) {
      if (hasModule(id) && App.state.data) refresh(id, false);
    }

    return { initialize, hydrate, hasModule, refreshOnOpen };
  })();
