/**
 * DashSheet - Application Logic (app.js)
 * Implements client-side filtering, interactive charts, dynamic custom multiselects,
 * sortable tables, and Excel file loading from server using SheetJS.
 */

document.addEventListener('DOMContentLoaded', () => {
  // Global Application State
  const state = {
    // Original imported data
    rawStakeholders: [],
    rawReports: [],
    // Currently filtered datasets
    filteredStakeholders: [],
    filteredReports: [],
    
    // Filter settings
    filters: {
      search: '',
      years: [],
      regions: [],
      countriesReq: [],
      countriesObj: [],
      institutions: [],
      themes: [] // Active thematic dimensions
    },
    
    // Table configurations
    sorting: {
      reports: { column: 'Ano', direction: 'desc' },
      stakeholders: { column: 'Instituição', direction: 'asc' }
    },
    pagination: {
      reports: { page: 1, pageSize: 20 },
      stakeholders: { page: 1, pageSize: 20 }
    },
    
    // Chart instances
    charts: {
      timeline: null,
      regions: null,
      institutions: null,
      thematics: null,
      timelineType: 'line' // 'line' or 'bar'
    }
  };

  // Thematic keys mapping (column name in Excel -> icon / label)
  const THEMATIC_COLUMNS = [
    { key: 'Sementes', label: 'Sementes', icon: 'fa-seedling' },
    { key: 'Político/geopolítico', label: 'Político/Geopolítico', icon: 'fa-globe' },
    { key: 'Econômico', label: 'Econômico', icon: 'fa-coins' },
    { key: 'Social', label: 'Social', icon: 'fa-users' },
    { key: 'Tecnológico', label: 'Tecnológico', icon: 'fa-laptop-code' },
    { key: 'Ambiental', label: 'Ambiental', icon: 'fa-leaf' },
    { key: 'Legal', label: 'Legal', icon: 'fa-scale-balanced' },
    { key: 'Militar', label: 'Militar', icon: 'fa-shield-halved' },
    { key: 'Marítimo', label: 'Marítimo', icon: 'fa-anchor' },
    { key: 'Terrestre', label: 'Terrestre', icon: 'fa-mountain' },
    { key: 'Aeroespacial', label: 'Aeroespacial', icon: 'fa-rocket' }
  ];

  // Initialize application
  async function init() {
    // 1. Setup event listeners
    setupEventListeners();

    // 2. Try to load updated Excel from GitHub Pages server dynamically
    const loaded = await loadExcelFromServer();
    
    if (!loaded) {
      console.log('Server fetch failed or running locally. Falling back to local data.js snapshot.');
      const statusText = document.getElementById('loading-status');
      if (statusText) statusText.innerText = 'Carregando banco de dados local (data.js)...';
      
      // Fallback to data.js database
      if (typeof baseData !== 'undefined') {
        state.rawStakeholders = baseData.stakeholders || [];
        state.rawReports = baseData.reports || [];
      }
    }
    
    // 3. Setup dynamic filter options in dropdowns
    buildFilterDropdowns();
    
    // 4. Run initial filter and render dashboard
    applyFiltersAndRender();

    // 5. Hide loading screen
    hideLoadingOverlay();
  }

  // Helpers to control the loading screen
  function hideLoadingOverlay() {
    const overlay = document.getElementById('loading-overlay');
    if (overlay) {
      overlay.style.opacity = '0';
      overlay.style.pointerEvents = 'none';
      setTimeout(() => {
        overlay.style.visibility = 'hidden';
      }, 500);
    }
  }

  function updateLoadingStatus(text) {
    const statusText = document.getElementById('loading-status');
    if (statusText) statusText.innerText = text;
  }

  // Build filter dropdowns dynamically based on database unique values
  function buildFilterDropdowns() {
    // Helper to get sorted unique values
    const getUniqueValues = (list, key) => {
      const vals = list
        .map(item => item[key])
        .filter(val => val !== null && val !== undefined && val !== '');
      return [...new Set(vals)].sort((a, b) => {
        if (typeof a === 'number' && typeof b === 'number') return a - b;
        return String(a).localeCompare(String(b), 'pt-BR');
      });
    };

    // Get unique values for each filter
    const years = getUniqueValues(state.rawReports, 'Ano').map(y => Math.round(y));
    const regions = getUniqueValues(state.rawReports, 'Região');
    const countriesReq = getUniqueValues(state.rawReports, 'País Solicitante');
    const countriesObj = getUniqueValues(state.rawReports, 'País Objeto');
    
    // We get institutions from stakeholders so it matches the directory
    const institutions = getUniqueValues(state.rawStakeholders, 'Instituição');

    // Setup custom multiselects
    initMultiselect('filter-year', years, 'years');
    initMultiselect('filter-region', regions, 'regions');
    initMultiselect('filter-country-req', countriesReq, 'countriesReq');
    initMultiselect('filter-country-obj', countriesObj, 'countriesObj');
    initMultiselect('filter-institution', institutions, 'institutions');
  }

  // Helper to initialize custom multiselect dropdown logic
  function initMultiselect(containerId, options, filterKey) {
    const container = document.getElementById(containerId);
    if (!container) return;

    const selectedDiv = container.querySelector('.multiselect-selected');
    const optionsDiv = container.querySelector('.multiselect-options');
    const placeholder = selectedDiv.getAttribute('data-placeholder');
    
    // Clear old options
    optionsDiv.innerHTML = '';
    
    // Render option list
    options.forEach(opt => {
      const optionEl = document.createElement('div');
      optionEl.className = 'multiselect-option';
      optionEl.innerHTML = `<i class="fa-regular fa-square"></i> <span>${opt}</span>`;
      optionEl.setAttribute('data-value', opt);
      
      optionEl.addEventListener('click', (e) => {
        e.stopPropagation();
        
        const isSelected = state.filters[filterKey].includes(opt);
        if (isSelected) {
          state.filters[filterKey] = state.filters[filterKey].filter(item => item !== opt);
          optionEl.classList.remove('selected');
          optionEl.querySelector('i').className = 'fa-regular fa-square';
        } else {
          state.filters[filterKey].push(opt);
          optionEl.classList.add('selected');
          optionEl.querySelector('i').className = 'fa-solid fa-square-check';
        }
        
        updateMultiselectHeader(selectedDiv, state.filters[filterKey], placeholder);
        applyFiltersAndRender();
      });

      optionsDiv.appendChild(optionEl);
    });

    // Toggle dropdown open/close
    selectedDiv.onclick = (e) => {
      e.stopPropagation();
      
      // Close all other dropdowns first
      document.querySelectorAll('.custom-multiselect').forEach(el => {
        if (el !== container) el.classList.remove('open');
      });
      
      container.classList.toggle('open');
    };
  }

  // Update header text/tags inside a multiselect
  function updateMultiselectHeader(selectedDiv, selectedList, placeholder) {
    if (selectedList.length === 0) {
      selectedDiv.innerText = placeholder;
      return;
    }

    if (selectedList.length <= 2) {
      selectedDiv.innerHTML = selectedList.map(item => `
        <span class="multiselect-tag">
          ${item}
        </span>
      `).join('');
    } else {
      selectedDiv.innerText = `${selectedList.length} selecionados`;
    }
  }

  // Setup UI Event Listeners
  function setupEventListeners() {
    // 1. Close dropdowns on clicking outside
    document.addEventListener('click', () => {
      document.querySelectorAll('.custom-multiselect').forEach(el => el.classList.remove('open'));
    });

    // 2. Global search input
    const searchInput = document.getElementById('search-input');
    if (searchInput) {
      searchInput.addEventListener('input', (e) => {
        state.filters.search = e.target.value.toLowerCase().trim();
        applyFiltersAndRender();
      });
    }

    // 3. Thematic Filter Toggles (Checkboxes)
    const themeToggles = document.querySelectorAll('#thematic-toggles .theme-checkbox');
    themeToggles.forEach(toggle => {
      toggle.addEventListener('click', () => {
        const theme = toggle.getAttribute('data-theme');
        toggle.classList.toggle('active');
        
        if (state.filters.themes.includes(theme)) {
          state.filters.themes = state.filters.themes.filter(t => t !== theme);
        } else {
          state.filters.themes.push(theme);
        }
        applyFiltersAndRender();
      });
    });

    // 4. Reset Filters Button
    const resetBtn = document.getElementById('btn-reset');
    if (resetBtn) {
      resetBtn.addEventListener('click', () => {
        // Reset state filters
        state.filters.search = '';
        state.filters.years = [];
        state.filters.regions = [];
        state.filters.countriesReq = [];
        state.filters.countriesObj = [];
        state.filters.institutions = [];
        state.filters.themes = [];
        
        // Reset Search Input
        if (searchInput) searchInput.value = '';
        
        // Reset Theme Toggle UI classes
        themeToggles.forEach(t => t.classList.remove('active'));
        
        // Reset Multiselect Dropdowns UI headers and checkmarks
        document.querySelectorAll('.custom-multiselect').forEach(container => {
          const selectedDiv = container.querySelector('.multiselect-selected');
          const placeholder = selectedDiv.getAttribute('data-placeholder');
          selectedDiv.innerText = placeholder;
          
          container.querySelectorAll('.multiselect-option').forEach(opt => {
            opt.classList.remove('selected');
            opt.querySelector('i').className = 'fa-regular fa-square';
          });
        });

        applyFiltersAndRender();
      });
    }

    // 5. Tab switcher buttons
    const tabBtns = document.querySelectorAll('.tab-btn');
    tabBtns.forEach(btn => {
      btn.addEventListener('click', () => {
        const targetTab = btn.getAttribute('data-tab');
        
        // Update button active state
        tabBtns.forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        
        // Show target pane
        document.querySelectorAll('.tab-pane').forEach(pane => {
          pane.classList.remove('active');
          if (pane.id === targetTab) {
            pane.classList.add('active');
          }
        });
      });
    });

    // 6. Chart Actions (Timeline Line/Bar Toggle)
    const timelineToggleBtns = document.querySelectorAll('[data-chart-type]');
    timelineToggleBtns.forEach(btn => {
      btn.addEventListener('click', () => {
        timelineToggleBtns.forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        state.charts.timelineType = btn.getAttribute('data-chart-type');
        renderTimelineChart();
      });
    });

    // 7. Reports Table Sorting
    const reportHeaders = document.querySelectorAll('#table-reports-list th[data-sort]');
    reportHeaders.forEach(th => {
      th.addEventListener('click', () => {
        const col = th.getAttribute('data-sort');
        const dir = state.sorting.reports.column === col && state.sorting.reports.direction === 'asc' ? 'desc' : 'asc';
        state.sorting.reports = { column: col, direction: dir };
        
        // Update header sort icons
        reportHeaders.forEach(h => {
          const icon = h.querySelector('i');
          icon.className = 'fa-solid fa-sort';
        });
        th.querySelector('i').className = `fa-solid fa-sort-${dir === 'asc' ? 'up' : 'down'}`;
        
        renderReportsTable();
      });
    });

    // 8. Stakeholders Table Sorting
    const stakeholderHeaders = document.querySelectorAll('#table-stakeholders-list th[data-sort]');
    stakeholderHeaders.forEach(th => {
      th.addEventListener('click', () => {
        const col = th.getAttribute('data-sort');
        const dir = state.sorting.stakeholders.column === col && state.sorting.stakeholders.direction === 'asc' ? 'desc' : 'asc';
        state.sorting.stakeholders = { column: col, direction: dir };
        
        // Update header sort icons
        stakeholderHeaders.forEach(h => {
          const icon = h.querySelector('i');
          icon.className = 'fa-solid fa-sort';
        });
        th.querySelector('i').className = `fa-solid fa-sort-${dir === 'asc' ? 'up' : 'down'}`;
        
        renderStakeholdersTable();
      });
    });

    // 9. Pagination controls
    document.getElementById('btn-reports-prev').addEventListener('click', () => {
      if (state.pagination.reports.page > 1) {
        state.pagination.reports.page--;
        renderReportsTable();
      }
    });
    document.getElementById('btn-reports-next').addEventListener('click', () => {
      const maxPage = Math.ceil(state.filteredReports.length / state.pagination.reports.pageSize);
      if (state.pagination.reports.page < maxPage) {
        state.pagination.reports.page++;
        renderReportsTable();
      }
    });
    
    document.getElementById('btn-stakeholders-prev').addEventListener('click', () => {
      if (state.pagination.stakeholders.page > 1) {
        state.pagination.stakeholders.page--;
        renderStakeholdersTable();
      }
    });
    document.getElementById('btn-stakeholders-next').addEventListener('click', () => {
      const maxPage = Math.ceil(state.filteredStakeholders.length / state.pagination.stakeholders.pageSize);
      if (state.pagination.stakeholders.page < maxPage) {
        state.pagination.stakeholders.page++;
        renderStakeholdersTable();
      }
    });

    // 10. Page Size Select controls
    const repPageSizeSelect = document.getElementById('reports-page-size');
    if (repPageSizeSelect) {
      repPageSizeSelect.addEventListener('change', (e) => {
        state.pagination.reports.pageSize = parseInt(e.target.value);
        state.pagination.reports.page = 1;
        renderReportsTable();
      });
    }

    const shPageSizeSelect = document.getElementById('stakeholders-page-size');
    if (shPageSizeSelect) {
      shPageSizeSelect.addEventListener('change', (e) => {
        state.pagination.stakeholders.pageSize = parseInt(e.target.value);
        state.pagination.stakeholders.page = 1;
        renderStakeholdersTable();
      });
    }

    // 11. Theme Toggle Button
    const themeToggleBtn = document.getElementById('btn-theme-toggle');
    if (themeToggleBtn) {
      themeToggleBtn.addEventListener('click', () => {
        const isLight = document.body.classList.toggle('light-theme');
        const icon = themeToggleBtn.querySelector('i');
        if (isLight) {
          icon.className = 'fa-solid fa-sun';
        } else {
          icon.className = 'fa-solid fa-moon';
        }
        // Re-render charts to adapt colors
        renderCharts();
      });
    }
  }

  // Core Filtering Logic
  function applyFiltersAndRender() {
    // 1. Filter reports
    state.filteredReports = state.rawReports.filter(rep => {
      // A. Text Search
      if (state.filters.search) {
        const query = state.filters.search;
        const match = 
          (rep['Nome do Estudo'] && rep['Nome do Estudo'].toLowerCase().includes(query)) ||
          (rep['Instituição'] && rep['Instituição'].toLowerCase().includes(query)) ||
          (rep['País Objeto'] && rep['País Objeto'].toLowerCase().includes(query)) ||
          (rep['País Solicitante'] && rep['País Solicitante'].toLowerCase().includes(query)) ||
          (rep['Macrotema?'] && rep['Macrotema?'].toLowerCase().includes(query)) ||
          (rep['Palavra Chave 1'] && rep['Palavra Chave 1'].toLowerCase().includes(query)) ||
          (rep['Palavra Chave 2'] && rep['Palavra Chave 2'].toLowerCase().includes(query)) ||
          (rep['Palavra Chave 3'] && rep['Palavra Chave 3'].toLowerCase().includes(query));
        
        if (!match) return false;
      }
      
      // B. Year Filter
      if (state.filters.years.length > 0) {
        const repYear = rep['Ano'] ? Math.round(rep['Ano']) : null;
        if (!state.filters.years.includes(repYear)) return false;
      }
      
      // C. Region Filter
      if (state.filters.regions.length > 0) {
        if (!state.filters.regions.includes(rep['Região'])) return false;
      }
      
      // D. Requesting Country Filter
      if (state.filters.countriesReq.length > 0) {
        if (!state.filters.countriesReq.includes(rep['País Solicitante'])) return false;
      }
      
      // E. Target Country Filter
      if (state.filters.countriesObj.length > 0) {
        if (!state.filters.countriesObj.includes(rep['País Objeto'])) return false;
      }
      
      // F. Institution Filter
      if (state.filters.institutions.length > 0) {
        if (!state.filters.institutions.includes(rep['Instituição'])) return false;
      }
      
      // G. Thematic Indicators (AND match: must satisfy all checked themes)
      if (state.filters.themes.length > 0) {
        for (let theme of state.filters.themes) {
          const val = rep[theme];
          if (val !== 'sim' && val !== 'Sim' && val !== true && val !== 1) {
            return false;
          }
        }
      }
      
      return true;
    });

    // 2. Filter stakeholders based on active reports/filters
    state.filteredStakeholders = state.rawStakeholders.filter(sh => {
      const instName = sh['Instituição'];
      
      // Text Search matches stakeholder metadata
      if (state.filters.search) {
        const query = state.filters.search;
        const match = 
          (instName && instName.toLowerCase().includes(query)) ||
          (sh['Região'] && sh['Região'].toLowerCase().includes(query)) ||
          (sh['País'] && sh['País'].toLowerCase().includes(query)) ||
          (sh['Órgão Associado'] && sh['Órgão Associado'].toLowerCase().includes(query));
        
        if (!match) return false;
      }
      
      // Region Filter
      if (state.filters.regions.length > 0) {
        if (!state.filters.regions.includes(sh['Região'])) return false;
      }
      
      // Target Country Filter (map Pais to sh.País)
      if (state.filters.countriesObj.length > 0) {
        if (!state.filters.countriesObj.includes(sh['País'])) return false;
      }

      // Institution Filter
      if (state.filters.institutions.length > 0) {
        if (!state.filters.institutions.includes(instName)) return false;
      }
      
      return true;
    });

    // Dynamic Report counts mapping inside the filtered stakeholders list
    state.filteredStakeholders.forEach(sh => {
      const instName = sh['Instituição'];
      sh.dynamicReportCount = state.filteredReports.filter(rep => rep['Instituição'] === instName).length;
    });

    // Reset pagination to page 1
    state.pagination.reports.page = 1;
    state.pagination.stakeholders.page = 1;

    // Render components
    renderKPIs();
    renderCharts();
    renderReportsTable();
    renderStakeholdersTable();
    
    // Update live counts text
    const statusText = document.getElementById('database-status-text');
    if (statusText) {
      statusText.innerHTML = `Exibindo <strong style="color: var(--color-primary);">${state.filteredReports.length}</strong> de <strong>${state.rawReports.length}</strong> relatórios | <strong style="color: var(--color-secondary);">${state.filteredStakeholders.length}</strong> de <strong>${state.rawStakeholders.length}</strong> stakeholders`;
    }
  }

  // Render KPI Cards
  function renderKPIs() {
    // 1. Total Reports
    document.getElementById('kpi-total-reports').innerText = state.filteredReports.length.toLocaleString('pt-BR');
    
    // 2. Active Stakeholders (count distinct institutions present in the filtered reports)
    const activeInsts = new Set(state.filteredReports.map(r => r['Instituição']).filter(Boolean));
    document.getElementById('kpi-active-stakeholders').innerText = activeInsts.size.toLocaleString('pt-BR');
    
    // 3. Unique countries targeted
    const activeCountries = new Set(state.filteredReports.map(r => r['País Objeto']).filter(Boolean));
    document.getElementById('kpi-total-countries').innerText = activeCountries.size.toLocaleString('pt-BR');
    
    // 4. Defense and security percentage
    const defReports = state.filteredReports.filter(r => {
      const val = r['Trata de def/seg'];
      return val === 'sim' || val === 'Sim' || val === true || val === 1;
    });
    const percentage = state.filteredReports.length > 0 
      ? Math.round((defReports.length / state.filteredReports.length) * 100) 
      : 0;
    document.getElementById('kpi-defense-percentage').innerText = `${percentage}%`;
  }

  // Master Charts Controller
  function renderCharts() {
    renderTimelineChart();
    renderRegionsChart();
    renderInstitutionsChart();
    renderThematicsChart();
  }

  // Chart 1: Timeline
  function renderTimelineChart() {
    const ctx = document.getElementById('chart-timeline');
    if (!ctx) return;
    
    if (state.charts.timeline) state.charts.timeline.destroy();

    // Group reports by year
    const countsByYear = {};
    state.filteredReports.forEach(rep => {
      const year = rep['Ano'] ? Math.round(rep['Ano']) : 'Sem Ano';
      countsByYear[year] = (countsByYear[year] || 0) + 1;
    });

    const years = Object.keys(countsByYear).sort((a, b) => {
      if (a === 'Sem Ano') return 1;
      if (b === 'Sem Ano') return -1;
      return parseInt(a) - parseInt(b);
    });
    const counts = years.map(y => countsByYear[y]);

    const isLine = state.charts.timelineType === 'line';
    const isLight = document.body.classList.contains('light-theme');
    const tickColor = isLight ? '#475569' : 'hsl(215, 20%, 65%)';
    const gridColor = isLight ? 'rgba(0, 0, 0, 0.06)' : 'rgba(255, 255, 255, 0.03)';

    state.charts.timeline = new Chart(ctx, {
      type: isLine ? 'line' : 'bar',
      data: {
        labels: years,
        datasets: [{
          label: 'Relatórios Publicados',
          data: counts,
          borderColor: 'hsl(204, 100%, 48%)',
          backgroundColor: isLine ? 'rgba(0, 150, 255, 0.08)' : 'rgba(0, 150, 255, 0.65)',
          fill: true,
          tension: 0.35,
          borderWidth: 2,
          pointBackgroundColor: 'hsl(204, 100%, 55%)',
          pointBorderColor: isLight ? '#475569' : '#fff',
          pointHoverRadius: 6
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            backgroundColor: isLight ? 'rgba(255, 255, 255, 0.95)' : 'rgba(8, 12, 26, 0.95)',
            titleColor: isLight ? '#0f172a' : '#fff',
            bodyColor: isLight ? '#475569' : 'hsl(215, 20%, 85%)',
            borderColor: isLight ? 'rgba(0,0,0,0.08)' : 'rgba(255, 255, 255, 0.08)',
            borderWidth: 1
          }
        },
        scales: {
          x: {
            grid: { color: gridColor },
            ticks: { color: tickColor, font: { family: 'Inter', size: 10 } }
          },
          y: {
            grid: { color: gridColor },
            ticks: { color: tickColor, font: { family: 'Inter', size: 10 }, precision: 0 }
          }
        }
      }
    });
  }

  // Chart 2: Regional Distribution
  function renderRegionsChart() {
    const ctx = document.getElementById('chart-regions');
    if (!ctx) return;

    if (state.charts.regions) state.charts.regions.destroy();

    // Group reports by region
    const countsByRegion = {};
    state.filteredReports.forEach(rep => {
      const reg = rep['Região'] || 'Não especificada';
      countsByRegion[reg] = (countsByRegion[reg] || 0) + 1;
    });

    const regions = Object.keys(countsByRegion).sort((a, b) => countsByRegion[b] - countsByRegion[a]);
    const counts = regions.map(r => countsByRegion[r]);

    // Color gradient palette
    const colors = [
      'hsl(204, 100%, 48%)',  // Blue
      'hsl(268, 96%, 66%)',   // Purple
      'hsl(142, 70%, 45%)',   // Green
      'hsl(38, 92%, 50%)',    // Amber
      'hsl(350, 89%, 60%)',   // Red
      'hsl(190, 90%, 50%)',   // Cyan
      'hsl(320, 80%, 60%)',   // Pink
      'hsl(220, 20%, 40%)'    // Slate
    ];

    const isLight = document.body.classList.contains('light-theme');

    state.charts.regions = new Chart(ctx, {
      type: 'doughnut',
      data: {
        labels: regions,
        datasets: [{
          data: counts,
          backgroundColor: colors.slice(0, regions.length),
          borderColor: isLight ? '#fff' : 'rgba(13, 20, 38, 0.8)',
          borderWidth: 2,
          hoverOffset: 12
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: 'right',
            labels: {
              color: isLight ? '#475569' : 'hsl(215, 20%, 75%)',
              font: { family: 'Inter', size: 10 }
            }
          },
          tooltip: {
            backgroundColor: isLight ? 'rgba(255, 255, 255, 0.95)' : 'rgba(8, 12, 26, 0.95)',
            titleColor: isLight ? '#0f172a' : '#fff',
            bodyColor: isLight ? '#475569' : 'hsl(215, 20%, 85%)',
            borderColor: isLight ? 'rgba(0,0,0,0.08)' : 'rgba(255, 255, 255, 0.08)',
            borderWidth: 1
          }
        },
        cutout: '60%'
      }
    });
  }

  // Chart 3: Top Institutions
  function renderInstitutionsChart() {
    const ctx = document.getElementById('chart-institutions');
    if (!ctx) return;

    if (state.charts.institutions) state.charts.institutions.destroy();

    // Count reports per institution
    const countsByInst = {};
    state.filteredReports.forEach(rep => {
      const inst = rep['Instituição'] || 'Desconhecida';
      countsByInst[inst] = (countsByInst[inst] || 0) + 1;
    });

    // Sort and get Top 10
    const topInsts = Object.keys(countsByInst)
      .sort((a, b) => countsByInst[b] - countsByInst[a])
      .slice(0, 10);
    const counts = topInsts.map(inst => countsByInst[inst]);

    const isLight = document.body.classList.contains('light-theme');
    const tickColor = isLight ? '#475569' : 'hsl(215, 20%, 65%)';
    const gridColor = isLight ? 'rgba(0, 0, 0, 0.06)' : 'rgba(255, 255, 255, 0.03)';

    state.charts.institutions = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: topInsts,
        datasets: [{
          label: 'Nº de Relatórios',
          data: counts,
          backgroundColor: 'rgba(124, 58, 237, 0.7)', // Violet
          borderColor: 'hsl(268, 96%, 66%)',
          borderWidth: 1,
          borderRadius: 4
        }]
      },
      options: {
        indexAxis: 'y', // Horizontal bars
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            backgroundColor: isLight ? 'rgba(255, 255, 255, 0.95)' : 'rgba(8, 12, 26, 0.95)',
            titleColor: isLight ? '#0f172a' : '#fff',
            bodyColor: isLight ? '#475569' : 'hsl(215, 20%, 85%)',
            borderColor: isLight ? 'rgba(0,0,0,0.08)' : 'rgba(255, 255, 255, 0.08)',
            borderWidth: 1
          }
        },
        scales: {
          x: {
            grid: { color: gridColor },
            ticks: { color: tickColor, font: { family: 'Inter', size: 9 }, precision: 0 }
          },
          y: {
            grid: { display: false },
            ticks: {
              color: isLight ? '#0f172a' : '#fff',
              font: { family: 'Inter', size: 9 },
              callback: function(value) {
                const label = this.getLabelForValue(value);
                return label.length > 25 ? label.substring(0, 22) + '...' : label;
              }
            }
          }
        }
      }
    });
  }

  // Chart 4: Thematics
  function renderThematicsChart() {
    const ctx = document.getElementById('chart-thematics');
    if (!ctx) return;

    if (state.charts.thematics) state.charts.thematics.destroy();

    // Calculate occurrences for each thematic key
    const labels = THEMATIC_COLUMNS.map(col => col.label);
    const counts = THEMATIC_COLUMNS.map(col => {
      return state.filteredReports.filter(rep => {
        const val = rep[col.key];
        return val === 'sim' || val === 'Sim' || val === true || val === 1;
      }).length;
    });

    const isLight = document.body.classList.contains('light-theme');

    state.charts.thematics = new Chart(ctx, {
      type: 'radar',
      data: {
        labels: labels,
        datasets: [{
          label: 'Relatórios por Tema',
          data: counts,
          backgroundColor: 'rgba(16, 185, 129, 0.15)', // Green transparency
          borderColor: 'hsl(142, 70%, 45%)',
          borderWidth: 2,
          pointBackgroundColor: 'hsl(142, 70%, 50%)',
          pointBorderColor: isLight ? '#475569' : '#fff',
          pointHoverRadius: 5
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            backgroundColor: isLight ? 'rgba(255, 255, 255, 0.95)' : 'rgba(8, 12, 26, 0.95)',
            titleColor: isLight ? '#0f172a' : '#fff',
            bodyColor: isLight ? '#475569' : 'hsl(215, 20%, 85%)',
            borderColor: isLight ? 'rgba(0,0,0,0.08)' : 'rgba(255, 255, 255, 0.08)',
            borderWidth: 1
          }
        },
        scales: {
          r: {
            grid: { color: isLight ? 'rgba(0,0,0,0.08)' : 'rgba(255, 255, 255, 0.08)' },
            angleLines: { color: isLight ? 'rgba(0,0,0,0.08)' : 'rgba(255, 255, 255, 0.08)' },
            pointLabels: {
              color: isLight ? '#475569' : 'hsl(215, 20%, 75%)',
              font: { family: 'Inter', size: 9, weight: '500' }
            },
            ticks: {
              display: false,
              maxTicksLimit: 4
            }
          }
        }
      }
    });
  }

  // Render Reports Data Table
  function renderReportsTable() {
    const tbody = document.getElementById('reports-table-body');
    if (!tbody) return;

    // Apply Sorting
    const { column, direction } = state.sorting.reports;
    const sorted = [...state.filteredReports].sort((a, b) => {
      let valA = a[column];
      let valB = b[column];
      
      // Handle numeric sorting for Year/Chave
      if (column === 'Ano' || column === 'Chave Relatório') {
        valA = valA ? parseFloat(valA) : 0;
        valB = valB ? parseFloat(valB) : 0;
        return direction === 'asc' ? valA - valB : valB - valA;
      }
      
      // Handle string sorting
      valA = valA ? String(valA).toLowerCase() : '';
      valB = valB ? String(valB).toLowerCase() : '';
      return direction === 'asc' 
        ? valA.localeCompare(valB, 'pt-BR') 
        : valB.localeCompare(valA, 'pt-BR');
    });

    // Apply Pagination
    const { page, pageSize } = state.pagination.reports;
    const totalItems = sorted.length;
    const startIdx = (page - 1) * pageSize;
    const paginated = sorted.slice(startIdx, startIdx + pageSize);

    // Render Table Rows
    tbody.innerHTML = '';
    
    if (paginated.length === 0) {
      tbody.innerHTML = `<tr><td colspan="8" style="text-align: center; color: var(--text-muted); padding: 2rem;">Nenhum relatório encontrado correspondente aos filtros.</td></tr>`;
      document.getElementById('reports-pagination-info').innerText = 'Exibindo 0 de 0 relatórios';
      document.getElementById('btn-reports-prev').disabled = true;
      document.getElementById('btn-reports-next').disabled = true;
      return;
    }

    paginated.forEach(rep => {
      const year = rep['Ano'] ? Math.round(rep['Ano']) : 'Sem Ano';
      const defSeg = rep['Trata de def/seg'] === 'sim' || rep['Trata de def/seg'] === 'Sim' || rep['Trata de def/seg'] === true;
      
      // Build active theme badges
      let activeThemes = [];
      THEMATIC_COLUMNS.forEach(col => {
        const val = rep[col.key];
        if (val === 'sim' || val === 'Sim' || val === true || val === 1) {
          activeThemes.push(`<span class="badge-theme active" title="${col.label}">${col.label}</span>`);
        }
      });
      
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${year}</td>
        <td>${rep['Região'] || '-'}</td>
        <td title="${rep['País Objeto'] || ''}">${rep['País Objeto'] || '-'}</td>
        <td title="${rep['Instituição'] || ''}">${rep['Instituição'] || '-'}</td>
        <td style="white-space: normal; min-width: 250px; max-width: 320px;" title="${rep['Nome do Estudo']}">
          ${rep['Link'] 
            ? `<a href="${rep['Link']}" target="_blank" class="table-link">${rep['Nome do Estudo']} <i class="fa-solid fa-up-right-from-square" style="font-size: 0.7rem;"></i></a>`
            : rep['Nome do Estudo']}
        </td>
        <td>
          <span class="tag ${defSeg ? 'tag-sim' : 'tag-nao'}">${defSeg ? 'Sim' : 'Não'}</span>
        </td>
        <td>
          <div class="thematic-badges">
            ${activeThemes.length > 0 ? activeThemes.slice(0, 3).join('') : '<span class="badge-theme">-</span>'}
            ${activeThemes.length > 3 ? `<span class="badge-theme">+${activeThemes.length - 3}</span>` : ''}
          </div>
        </td>
        <td>
          <button class="chart-btn details-btn" style="padding: 4px 8px; font-size: 0.7rem;" data-id="${rep['Chave Relatório'] || ''}">
            <i class="fa-solid fa-circle-info"></i> Detalhes
          </button>
        </td>
      `;

      tr.querySelector('.details-btn').addEventListener('click', () => {
        showReportDetails(rep);
      });

      tbody.appendChild(tr);
    });

    // Update Pagination UI
    const endIdx = Math.min(startIdx + pageSize, totalItems);
    document.getElementById('reports-pagination-info').innerHTML = `Exibindo <strong>${startIdx + 1}-${endIdx}</strong> de <strong>${totalItems}</strong> relatórios`;
    
    document.getElementById('btn-reports-prev').disabled = page === 1;
    document.getElementById('btn-reports-next').disabled = endIdx >= totalItems;
  }

  // Render Stakeholders Directory
  function renderStakeholdersTable() {
    const tbody = document.getElementById('stakeholders-table-body');
    if (!tbody) return;

    // Apply Sorting
    const { column, direction } = state.sorting.stakeholders;
    const sorted = [...state.filteredStakeholders].sort((a, b) => {
      let valA = a[column];
      let valB = b[column];
      
      // Numeric fields
      if (column === 'Relatórios por instituição' || column === 'dynamicReportCount' || column === 'Chave') {
        valA = valA ? parseFloat(valA) : 0;
        valB = valB ? parseFloat(valB) : 0;
        return direction === 'asc' ? valA - valB : valB - valA;
      }
      
      // String fields
      valA = valA ? String(valA).toLowerCase() : '';
      valB = valB ? String(valB).toLowerCase() : '';
      return direction === 'asc' 
        ? valA.localeCompare(valB, 'pt-BR') 
        : valB.localeCompare(valA, 'pt-BR');
    });

    // Apply Pagination
    const { page, pageSize } = state.pagination.stakeholders;
    const totalItems = sorted.length;
    const startIdx = (page - 1) * pageSize;
    const paginated = sorted.slice(startIdx, startIdx + pageSize);

    tbody.innerHTML = '';
    
    if (paginated.length === 0) {
      tbody.innerHTML = `<tr><td colspan="6" style="text-align: center; color: var(--text-muted); padding: 2rem;">Nenhum stakeholder encontrado.</td></tr>`;
      document.getElementById('stakeholders-pagination-info').innerText = 'Exibindo 0 de 0 stakeholders';
      document.getElementById('btn-stakeholders-prev').disabled = true;
      document.getElementById('btn-stakeholders-next').disabled = true;
      return;
    }

    paginated.forEach(sh => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${sh['Região'] || '-'}</td>
        <td>${sh['País'] || '-'}</td>
        <td style="font-weight: 600; color: var(--text-main);" title="${sh['Instituição'] || ''}">${sh['Instituição'] || '-'}</td>
        <td>${sh['Órgão Associado'] || '-'}</td>
        <td style="text-align: center; font-weight: bold; color: var(--color-primary);">
          ${sh.dynamicReportCount}
        </td>
        <td>
          ${sh['Link'] 
            ? `<a href="${sh['Link']}" target="_blank" class="table-link" title="${sh['Link']}">Acessar Site <i class="fa-solid fa-arrow-up-right-from-square" style="font-size: 0.75rem;"></i></a>`
            : '-'}
        </td>
      `;
      tbody.appendChild(tr);
    });

    // Update Pagination UI
    const endIdx = Math.min(startIdx + pageSize, totalItems);
    document.getElementById('stakeholders-pagination-info').innerHTML = `Exibindo <strong>${startIdx + 1}-${endIdx}</strong> de <strong>${totalItems}</strong> stakeholders`;
    
    document.getElementById('btn-stakeholders-prev').disabled = page === 1;
    document.getElementById('btn-stakeholders-next').disabled = endIdx >= totalItems;
  }

  // Expanded report detail modal window (built dynamically)
  function showReportDetails(rep) {
    // Check if modal already exists
    let modal = document.getElementById('report-details-modal');
    if (!modal) {
      modal = document.createElement('div');
      modal.id = 'report-details-modal';
      modal.className = 'report-modal-overlay';
      document.body.appendChild(modal);
    }
    
    // Build active themes tags list
    let themesHtml = '';
    THEMATIC_COLUMNS.forEach(col => {
      const val = rep[col.key];
      const isSim = val === 'sim' || val === 'Sim' || val === true || val === 1;
      if (isSim) {
        themesHtml += `<span class="badge-theme active" style="padding: 4px 10px; font-size: 0.8rem; margin: 4px;"><i class="fa-solid ${col.icon}"></i> ${col.label}</span>`;
      }
    });

    modal.innerHTML = `
      <div class="report-modal-container">
        <div style="padding: 1.5rem; border-bottom: 1px solid var(--border-glass); display: flex; justify-content: space-between; align-items: center;">
          <h3 style="color: var(--text-main); font-size: 1.25rem; font-weight: 800;"><i class="fa-solid fa-file-invoice" style="color: var(--color-primary);"></i> Detalhes do Estudo</h3>
          <button id="modal-close-btn" style="background: none; border: none; color: var(--text-muted); font-size: 1.25rem; cursor: pointer;"><i class="fa-solid fa-xmark"></i></button>
        </div>
        <div style="padding: 2rem; overflow-y: auto; max-height: 70vh; display: flex; flex-direction: column; gap: 1.5rem;">
          <div>
            <h4 style="font-size: 1.1rem; color: var(--text-main); font-weight: 700; margin-bottom: 0.5rem; line-height: 1.4;">${rep['Nome do Estudo']}</h4>
            ${rep['Link'] ? `<a href="${rep['Link']}" target="_blank" style="color: var(--color-primary); font-size: 0.85rem; text-decoration: none;"><i class="fa-solid fa-link"></i> Abrir Link da Publicação</a>` : ''}
          </div>
          
          <div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 1rem; background: rgba(255,255,255,0.02); border: 1px solid var(--border-glass); padding: 1rem; border-radius: var(--border-radius-md);">
            <div><span style="font-size: 0.75rem; color: var(--text-muted); font-weight: bold; text-transform: uppercase;">Instituição</span><div style="color: var(--text-main); font-size: 0.9rem; font-weight: 500; margin-top: 2px;">${rep['Instituição'] || '-'}</div></div>
            <div><span style="font-size: 0.75rem; color: var(--text-muted); font-weight: bold; text-transform: uppercase;">Ano de Publicação</span><div style="color: var(--text-main); font-size: 0.9rem; font-weight: 500; margin-top: 2px;">${rep['Ano'] ? Math.round(rep['Ano']) : '-'}</div></div>
            <div><span style="font-size: 0.75rem; color: var(--text-muted); font-weight: bold; text-transform: uppercase;">Região</span><div style="color: var(--text-main); font-size: 0.9rem; font-weight: 500; margin-top: 2px;">${rep['Região'] || '-'}</div></div>
            <div><span style="font-size: 0.75rem; color: var(--text-muted); font-weight: bold; text-transform: uppercase;">País Solicitante</span><div style="color: var(--text-main); font-size: 0.9rem; font-weight: 500; margin-top: 2px;">${rep['País Solicitante'] || '-'}</div></div>
            <div><span style="font-size: 0.75rem; color: var(--text-muted); font-weight: bold; text-transform: uppercase;">País Objeto (Alvo)</span><div style="color: var(--text-main); font-size: 0.9rem; font-weight: 500; margin-top: 2px;">${rep['País Objeto'] || '-'}</div></div>
            <div><span style="font-size: 0.75rem; color: var(--text-muted); font-weight: bold; text-transform: uppercase;">Defesa & Segurança</span><div style="margin-top: 4px;"><span class="tag ${rep['Trata de def/seg'] === 'sim' || rep['Trata de def/seg'] === 'Sim' || rep['Trata de def/seg'] === true ? 'tag-sim' : 'tag-nao'}">${rep['Trata de def/seg'] === 'sim' || rep['Trata de def/seg'] === 'Sim' ? 'Sim' : 'Não'}</span></div></div>
          </div>
          
          <div>
            <span style="font-size: 0.75rem; color: var(--text-muted); font-weight: bold; text-transform: uppercase; display: block; margin-bottom: 0.5rem;">Temas Identificados</span>
            <div style="display: flex; flex-wrap: wrap;">
              ${themesHtml || '<span style="color: var(--text-muted); font-size: 0.85rem;">Nenhum tema demarcado na base.</span>'}
            </div>
          </div>

          ${rep['Macrotema?'] || rep['Palavra Chave 1'] || rep['Palavra Chave 2'] || rep['Palavra Chave 3'] ? `
            <div style="background: rgba(0, 150, 255, 0.03); border: 1px dashed rgba(0, 150, 255, 0.15); padding: 1rem; border-radius: var(--border-radius-md);">
              ${rep['Macrotema?'] ? `<div style="margin-bottom: 0.5rem;"><strong style="font-size: 0.75rem; color: var(--text-muted); text-transform: uppercase;">Macrotema:</strong> <span style="font-size: 0.85rem; color: var(--text-main); margin-left: 6px;">${rep['Macrotema?']}</span></div>` : ''}
              <div style="display: flex; gap: 8px; flex-wrap: wrap; align-items: center;">
                <strong style="font-size: 0.75rem; color: var(--text-muted); text-transform: uppercase;">Palavras-Chave:</strong>
                ${rep['Palavra Chave 1'] ? `<span class="badge-theme">${rep['Palavra Chave 1']}</span>` : ''}
                ${rep['Palavra Chave 2'] ? `<span class="badge-theme">${rep['Palavra Chave 2']}</span>` : ''}
                ${rep['Palavra Chave 3'] ? `<span class="badge-theme">${rep['Palavra Chave 3']}</span>` : ''}
              </div>
            </div>
          ` : ''}
        </div>
        <div style="padding: 1rem 1.5rem; border-top: 1px solid var(--border-glass); background: rgba(0,0,0,0.05); display: flex; justify-content: flex-end;">
          <button id="modal-ok-btn" class="tab-btn active" style="padding: 8px 20px;">Fechar</button>
        </div>
      </div>
    `;

    modal.style.display = 'flex';

    // Close logic
    const closeModal = () => {
      modal.style.display = 'none';
    };
    modal.querySelector('#modal-close-btn').onclick = closeModal;
    modal.querySelector('#modal-ok-btn').onclick = closeModal;
    
    // Close on overlay click
    modal.onclick = (e) => {
      if (e.target === modal) closeModal();
    };
  }

  // Core sheet parser & visual normalizer
  function processExcelWorkbook(workbook) {
    // Target Sheet Names (with correct case-sensitive names)
    const sheetStakeholders = workbook.Sheets['Stakeholders'];
    const sheetReports = workbook.Sheets['Relatórios'];
    
    if (!sheetStakeholders || !sheetReports) {
      throw new Error('A planilha deve conter as abas "Stakeholders" e "Relatórios"!');
    }

    // Convert to raw JSON lists
    const rawSh = XLSX.utils.sheet_to_json(sheetStakeholders, { defval: null });
    const rawRep = XLSX.utils.sheet_to_json(sheetReports, { defval: null });

    // Clean & Normalize Stakeholders
    state.rawStakeholders = rawSh.map(row => {
      const cleanRow = {};
      Object.keys(row).forEach(key => {
        let cleanKey = key.trim();
        // Filter out formula headers or empty columns
        if (!cleanKey.startsWith('=')) {
          let val = row[key];
          if (typeof val === 'string') val = val.trim();
          cleanRow[cleanKey] = val;
        }
      });
      return cleanRow;
    }).filter(item => item['Instituição']); // Require Institution name

    // Clean & Normalize Reports
    state.rawReports = rawRep.map((row, idx) => {
      const cleanRow = {};
      Object.keys(row).forEach(key => {
        let cleanKey = key.trim();
        if (!cleanKey.startsWith('=')) {
          let val = row[key];
          if (typeof val === 'string') val = val.trim();
          
          // Standard Date conversion if object
          if (val instanceof Date) {
            // Convert date to YYYY-MM-DD
            const y = val.getFullYear();
            const m = String(val.getMonth() + 1).padStart(2, '0');
            const d = String(val.getDate()).padStart(2, '0');
            val = `${y}-${m}-${d}`;
          }
          
          cleanRow[cleanKey] = val;
        }
      });
      
      // Generate unique key if missing
      if (!cleanRow['Chave Relatório']) {
        cleanRow['Chave Relatório'] = idx + 1;
      }
      return cleanRow;
    }).filter(item => item['Nome do Estudo']); // Require Study name

    console.log('Successfully parsed Excel sheet data!');
    console.log('Stakeholders:', state.rawStakeholders.length);
    console.log('Reports:', state.rawReports.length);
  }

  // Fetch Excel sheet from server paths (useful for GitHub Pages direct sync)
  async function loadExcelFromServer() {
    // List of relative URLs to search for the Excel sheet (looking in current directory first)
    const urls = [
      './base.xlsx',
      './database.xlsx',
      '../BasePPDef/base.xlsx',
      '../BasePPDef/2 - Stakeholders e Relatórios_BasePainel_Atualizacao 2025.xlsx',
      './BasePPDef/2 - Stakeholders e Relatórios_BasePainel_Atualizacao 2025.xlsx'
    ];

    for (const url of urls) {
      try {
        const encodedUrl = encodeURI(url);
        updateLoadingStatus(`Buscando base online: ${url.replace('../', '')}...`);
        
        const response = await fetch(encodedUrl);
        if (!response.ok) {
          continue; // Try next URL
        }
        
        updateLoadingStatus('Planilha encontrada! Carregando dados...');
        const arrayBuffer = await response.arrayBuffer();
        const data = new Uint8Array(arrayBuffer);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true, cellNF: false, cellText: false });
        
        // Parse and clean workbook
        processExcelWorkbook(workbook);
        return true; // Successfully loaded!
      } catch (e) {
        console.warn(`Could not load Excel from ${url}:`, e);
      }
    }
    return false; // All options failed
  }

  // Run the initializer!
  init();
});
