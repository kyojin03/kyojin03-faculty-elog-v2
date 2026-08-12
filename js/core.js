  'use strict';

  const App = (() => {
    const NAV_ITEMS = [
      { id: 'logbook', label: 'Logbook', title: 'Logbook', icon: '<path d="M5 3h14v18H5V3Zm2 2v14h10V5H7Zm2 3h6v2H9V8Zm0 4h6v2H9v-2Zm0 4h4v2H9v-2Z"/>' },
      { id: 'reports', label: 'Reports', title: 'Reports', icon: '<path d="M4 19h16v2H4v-2Zm2-2V9h3v8H6Zm5 0V3h3v14h-3Zm5 0v-6h3v6h-3Z"/>' },
      { id: 'consumables', label: 'Laboratory Supplies', title: 'Laboratory Supplies', icon: '<path d="M5 3h14v18H5V3Zm2 2v4h10V5H7Zm0 6v8h10v-8H7Zm2 2h3v2H9v-2Z"/>' },
      { id: 'glassware', label: 'Glassware', title: 'Glassware', icon: '<path d="M8 2h8v2h-1v5.3l3.3 7.6A3.6 3.6 0 0 1 15 22H9a3.6 3.6 0 0 1-3.3-5.1L9 9.3V4H8V2Zm3 8-3.5 7.7A1.6 1.6 0 0 0 9 20h6a1.6 1.6 0 0 0 1.5-2.3L13 10V4h-2v6Z"/>' },
      { id: 'equipment', label: 'Equipment', title: 'Equipment', icon: '<path d="M7 3h10v5h3v13H4V8h3V3Zm2 2v5H6v9h12v-9h-3V5H9Zm2 0v3h2V5h-2Z"/>' },
      { id: 'specialized', label: 'Specialized Equipment', title: 'Specialized Equipment', icon: '<path d="M12 2a4 4 0 0 1 4 4c0 1-.4 2-1 2.7l5.4 9.4-1.8 1-5.3-9.3c-.4.1-.9.2-1.3.2s-.9-.1-1.3-.2l-5.3 9.3-1.8-1L9 8.7A4 4 0 0 1 12 2Zm0 2a2 2 0 1 0 0 4 2 2 0 0 0 0-4Zm-1 10h2v8h-2v-8Z"/>' },
      { id: 'forms', label: 'Laboratory Forms', title: 'Laboratory Forms', icon: '<path d="M6 2h9l5 5v15H6V2Zm2 2v16h10V8h-4V4H8Zm2 7h6v2h-6v-2Zm0 4h6v2h-6v-2Z"/>' }
    ];

    const state = {
      activeView: 'logbook',
      data: null,
      allLogs: [],
      displayedLogs: [],
      rooms: [],
      settings: [],
      reportsMonth: '',
      reportData: null,
      submitPending: false,
      logRequestId: 0,
      reportRequestId: 0
    };

    const byId = id => document.getElementById(id);
    const value = id => (byId(id)?.value || '').trim();

    function escapeHtml(input) {
      return String(input == null ? '' : input).replace(/[&<>"']/g, char => ({
        '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
      })[char]);
    }

    function escapeAttribute(input) {
      return escapeHtml(input).replace(/`/g, '&#96;');
    }

    function debounce(callback, delay) {
      let timer;
      return (...args) => {
        window.clearTimeout(timer);
        timer = window.setTimeout(() => callback(...args), delay);
      };
    }

    function runServer(functionName, ...args) {
      return window.FacultyElogApi.request(functionName, ...args);
    }

    function logTechnicalError(context, error) {
      console.error('[Faculty eLog] ' + context, error);
    }

    function showToast(message, type = 'info', title) {
      const region = byId('toastRegion');
      const toast = document.createElement('div');
      toast.className = 'toast ' + type;
      toast.setAttribute('role', type === 'error' ? 'alert' : 'status');
      const labels = { success: 'Success', error: 'Something went wrong', info: 'Faculty eLog' };
      const paths = {
        success: '<path d="m9 16.2-3.5-3.5L4.1 14.1 9 19 20.3 7.7l-1.4-1.4L9 16.2Z"/>',
        error: '<path d="M11 7h2v6h-2V7Zm0 8h2v2h-2v-2Zm1-13a10 10 0 1 0 0 20 10 10 0 0 0 0-20Zm0 18a8 8 0 1 1 0-16 8 8 0 0 1 0 16Z"/>',
        info: '<path d="M11 10h2v7h-2v-7Zm0-3h2v2h-2V7Zm1-5a10 10 0 1 0 0 20 10 10 0 0 0 0-20Zm0 18a8 8 0 1 1 0-16 8 8 0 0 1 0 16Z"/>'
      };
      toast.innerHTML = `<svg viewBox="0 0 24 24" aria-hidden="true">${paths[type] || paths.info}</svg><div><strong>${escapeHtml(title || labels[type] || labels.info)}</strong><span>${escapeHtml(message)}</span></div><button type="button" aria-label="Dismiss notification">&times;</button>`;
      toast.querySelector('button').addEventListener('click', () => toast.remove());
      region.appendChild(toast);
      window.setTimeout(() => toast.remove(), type === 'error' ? 6500 : 4200);
    }

    function setStartupState(mode) {
      const overlay = byId('startupOverlay');
      if (mode === 'ready') {
        overlay.classList.add('hidden');
        overlay.setAttribute('aria-hidden', 'true');
        return;
      }
      overlay.classList.remove('hidden');
      overlay.removeAttribute('aria-hidden');
      byId('startupLoading').hidden = mode !== 'loading';
      byId('startupError').hidden = mode !== 'error';
    }

    function buildNavigation() {
      byId('primaryNav').innerHTML = NAV_ITEMS.map((item, index) => `
        <button class="nav-button${index === 0 ? ' active' : ''}" type="button" data-view="${item.id}" aria-current="${index === 0 ? 'page' : 'false'}">
          <svg viewBox="0 0 24 24" aria-hidden="true">${item.icon}</svg><span>${escapeHtml(item.label)}</span>
        </button>`).join('');

      byId('primaryNav').addEventListener('click', event => {
        const button = event.target.closest('[data-view]');
        if (button) showView(button.dataset.view);
      });
    }

    function showView(id) {
      const item = NAV_ITEMS.find(navItem => navItem.id === id);
      if (!item) return;
      state.activeView = id;
      document.querySelectorAll('.nav-button').forEach(button => {
        const active = button.dataset.view === id;
        button.classList.toggle('active', active);
        button.setAttribute('aria-current', active ? 'page' : 'false');
      });
      document.querySelectorAll('.view').forEach(view => view.classList.toggle('active', view.id === 'view-' + id));
      byId('pageTitle').textContent = item.title;
      closeNavigation();
      window.scrollTo({ top: 0, behavior: 'smooth' });
      if (id === 'reports') Reports.ensureLoaded();
      else if (Inventory.hasModule(id)) Inventory.refreshOnOpen(id);
    }

    function openNavigation() {
      document.body.classList.add('nav-open');
      byId('menuButton').setAttribute('aria-expanded', 'true');
      window.setTimeout(() => byId('primaryNav').querySelector('.active')?.focus(), 180);
    }

    function closeNavigation() {
      document.body.classList.remove('nav-open');
      byId('menuButton').setAttribute('aria-expanded', 'false');
    }

    function bindCoreEvents() {
      byId('menuButton').addEventListener('click', () => document.body.classList.contains('nav-open') ? closeNavigation() : openNavigation());
      byId('navScrim').addEventListener('click', closeNavigation);
      document.addEventListener('keydown', event => { if (event.key === 'Escape') closeNavigation(); });
      byId('retryInitialLoad').addEventListener('click', loadInitialData);
    }

    function normalizeData(data) {
      const safe = data && typeof data === 'object' ? data : {};
      return {
        rooms: Array.isArray(safe.rooms) ? safe.rooms : [],
        consumables: Array.isArray(safe.consumables) ? safe.consumables : [],
        glassware: Array.isArray(safe.glassware) ? safe.glassware : [],
        equipment: Array.isArray(safe.equipment) ? safe.equipment : [],
        specializedEquipment: Array.isArray(safe.specializedEquipment) ? safe.specializedEquipment : [],
        laboratoryForms: Array.isArray(safe.laboratoryForms) ? safe.laboratoryForms : [],
        settings: Array.isArray(safe.settings) ? safe.settings : [],
        logbook: Array.isArray(safe.logbook) ? safe.logbook : []
      };
    }

    async function loadInitialData() {
      setStartupState('loading');
      try {
        const data = normalizeData(await runServer('getInitialData'));
        state.data = data;
        state.rooms = data.rooms;
        state.settings = data.settings;
        state.allLogs = data.logbook;
        state.displayedLogs = data.logbook;
        Logbook.hydrate();
        Inventory.hydrate();
        setStartupState('ready');
      } catch (error) {
        logTechnicalError('Initial data load failed.', error);
        setStartupState('error');
      }
    }

    function initialize() {
      buildNavigation();
      bindCoreEvents();
      Logbook.initialize();
      Reports.initialize();
      Inventory.initialize();
      loadInitialData();
    }

    document.addEventListener('DOMContentLoaded', initialize);

    return {
      state, byId, value, escapeHtml, escapeAttribute, debounce, runServer,
      logTechnicalError, showToast, showView, normalizeData
    };
  })();
