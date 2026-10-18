/**
 * ==========================================================================
 * SFCSED — Plataforma de Sementes de Futuro | Ensino Superior 2040
 * Application Logic (app.js) — Modo Leitor Prospectivo (Estilo Email Client)
 * ==========================================================================
 */

document.addEventListener('DOMContentLoaded', () => {

  // Global State
  const state = {
    // Raw Datasets
    rawSementes: [],
    rawArticles: [],
    articlesMap: new Map(),

    // Filtered Datasets
    filteredSementes: [],
    filteredArticles: [],

    // Currently Selected Seed (for Reading Pane)
    selectedSfId: null,

    // Filters for Sementes (Multi-select Checkbox System)
    filters: {
      search: '',
      eixos: [],
      classificacoes: [],
      pestels: [],
      topicos: [],
      status: []
    },

    // Sort order for Sementes List
    listSort: 'id-asc',

    // Filters for Scopus Articles
    scopusFilters: {
      search: '',
      relation: 'linked', // '', 'linked', 'unlinked'
      year: ''
    },

    // View & Layout
    activeTab: 'sf', // 'sf' or 'scopus'
    analyticsExpanded: false,
    theme: 'dark',

    // Scopus Sorting & Pagination
    sortingScopus: { column: 'id', direction: 'asc' },
    paginationScopus: { page: 1, pageSize: 30 },

    // Chart instances
    charts: {
      classificacao: null,
      pestel: null,
      eixos: null
    }
  };

  // Color Palettes for Sober Executive Presentation
  const PALETTE = {
    dark: {
      text: '#94a3b8',
      textStrong: '#f8fafc',
      grid: 'rgba(255, 255, 255, 0.07)',
      typology: {
        'Ruptura': '#ef4444',
        'Tendência': '#3b82f6',
        'Sinal fraco': '#a855f7',
        'Sinal Fraco': '#a855f7',
        'Incerteza': '#f59e0b',
        'Fato predeterminado': '#94a3b8',
        'Surpresa inevitável': '#f97316',
        'Achado Empírico': '#14b8a6'
      },
      pestel: {
        'Político': '#f43f5e',
        'Econômico': '#f59e0b',
        'Social': '#3b82f6',
        'Tecnológico': '#8b5cf6',
        'Ambiental': '#10b981',
        'Legal': '#64748b'
      },
      eixos: ['#2563eb', '#0284c7', '#4f46e5', '#7c3aed', '#059669']
    },
    light: {
      text: '#64748b',
      textStrong: '#0f172a',
      grid: '#e2e8f0',
      typology: {
        'Ruptura': '#dc2626',
        'Tendência': '#2563eb',
        'Sinal fraco': '#9333ea',
        'Sinal Fraco': '#9333ea',
        'Incerteza': '#d97706',
        'Fato predeterminado': '#475569',
        'Surpresa inevitável': '#ea580c',
        'Achado Empírico': '#0d9488'
      },
      pestel: {
        'Político': '#e11d48',
        'Econômico': '#d97706',
        'Social': '#2563eb',
        'Tecnológico': '#7c3aed',
        'Ambiental': '#059669',
        'Legal': '#475569'
      },
      eixos: ['#1d4ed8', '#0284c7', '#4338ca', '#6d28d9', '#047857']
    }
  };

  // ========================================================================
  // INITIALIZATION
  // ========================================================================
  async function init() {
    setupEventListeners();
    setupTheme();
    setupFontScale();

    // Load data from Excel (via SheetJS) or fallback to data.js
    const loaded = await loadDataPipeline();
    if (!loaded) {
      console.error('Falha crítica no carregamento de dados.');
      updateLoadingStatus('Erro ao carregar dados. Verifique o arquivo SF CSED.xlsx ou data.js.');
      return;
    }

    // Populate Checkbox Filters & Collapsible Sections
    buildFilterCheckboxLists();
    setupCollapsibleFilterSections();

    // Initial Filter & Render
    applySfFilters();
    applyScopusFilters();
    renderKPIs();
    renderCharts();

    hideLoadingOverlay();
  }

  // ========================================================================
  // DATA PIPELINE: SHEETJS FETCH WITH DATA.JS FALLBACK
  // ========================================================================
  function updateDataSourceStatus(source, info) {
    const pill = document.getElementById('data-source-pill');
    if (!pill) return;

    if (source === 'excel') {
      pill.innerHTML = `<i class="fa-solid fa-file-excel source-status-live"></i> <span id="data-source-text">Excel Dinâmico</span>`;
      pill.title = 'Base conectada diretamente ao arquivo Excel SF CSED.xlsx (leitura via SheetJS em tempo real).';
    } else {
      const ts = (typeof sfData !== 'undefined' && sfData.metadata && sfData.metadata.generatedAt) ? sfData.metadata.generatedAt : '';
      const label = ts ? `Base: ${ts}` : 'Base: data.js';
      pill.innerHTML = `<i class="fa-solid fa-database source-status-file"></i> <span id="data-source-text">${label}</span>`;
      pill.title = `Carregado do snapshot consolidado local (data.js).\nÚltima sincronização: ${ts || 'Não informada'}\nPara sincronizar edições do Excel: execute atualizar_dados.command e recarregue a página.`;
    }
  }

  async function loadDataPipeline() {
    updateLoadingStatus('Iniciando carregamento da base prospectiva...');

    const isHttp = window.location.protocol.startsWith('http');

    // Sob protocolo file://, chamadas fetch() locais são bloqueadas por CORS na maioria dos navegadores.
    // Usamos diretamente o snapshot consolidado (data.js) para carregamento instantâneo (~5ms) sem erros.
    if (!isHttp && typeof sfData !== 'undefined' && sfData.sementes && sfData.artigos) {
      state.rawSementes = sfData.sementes;
      state.rawArticles = sfData.artigos;
      state.articlesMap = new Map(sfData.artigos.map(a => [a.id, a]));
      console.log('Dados carregados instantaneamente do snapshot local (data.js) via protocolo file://');
      updateDataSourceStatus('datajs');
      return true;
    }

    let loadedFromExcel = false;
    const bust = isHttp ? ('?v=' + Date.now()) : '';
    const excelUrls = [
      `./SF CSED.xlsx${bust}`,
      `./SF%20CSED.xlsx${bust}`,
      `../sfcsed/SF CSED.xlsx${bust}`
    ];

    if (isHttp && typeof XLSX !== 'undefined') {
      for (const url of excelUrls) {
        try {
          const resp = await fetch(url);
          if (resp.ok) {
            updateLoadingStatus('Processando planilhas Excel...');
            const arrayBuffer = await resp.arrayBuffer();
            const workbook = XLSX.read(arrayBuffer, { type: 'array' });
            const sfSheetName = findSheetName(workbook.SheetNames, ['Sementes', 'SF (Scopus)', 'SF']);
            const refSheetName = findSheetName(workbook.SheetNames, ['Referências', 'Referencias', 'Scopus 3 (Limpo)', 'Scopus 3', 'Scopus']);
            if (sfSheetName && refSheetName) {
              parseExcelWorkbook(workbook, sfSheetName, refSheetName);
              loadedFromExcel = true;
              console.log('Dados carregados dinamicamente de:', url);
              updateDataSourceStatus('excel', url);
              break;
            }
          }
        } catch (err) {
          // fallback will be used
        }
      }
    }

    if (!loadedFromExcel) {
      updateLoadingStatus('Carregando base consolidada local (data.js)...');
      if (typeof sfData !== 'undefined' && sfData.sementes && sfData.artigos) {
        state.rawSementes = sfData.sementes;
        state.rawArticles = sfData.artigos;
        state.articlesMap = new Map(sfData.artigos.map(a => [a.id, a]));
        console.log('Dados carregados do snapshot data.js');
        updateDataSourceStatus('datajs');
        return true;
      } else {
        return false;
      }
    }

    return true;
  }

  function findSheetName(sheetNames, candidates) {
    if (!sheetNames || !Array.isArray(sheetNames)) return null;
    for (const cand of candidates) {
      const candNorm = removeAccents(cand.trim().toLowerCase());
      const found = sheetNames.find(s => removeAccents(s.trim().toLowerCase()) === candNorm);
      if (found) return found;
    }
    return null;
  }

  function getRowVal(row, candidates, defaultVal = '') {
    if (!row) return defaultVal;
    for (const cand of candidates) {
      if (row[cand] !== undefined && row[cand] !== null && String(row[cand]).trim() !== '') {
        return row[cand];
      }
    }
    const rowKeys = Object.keys(row);
    for (const cand of candidates) {
      const candNorm = removeAccents(cand.trim().toLowerCase());
      const matchKey = rowKeys.find(k => removeAccents(k.trim().toLowerCase()) === candNorm);
      if (matchKey && row[matchKey] !== undefined && row[matchKey] !== null && String(row[matchKey]).trim() !== '') {
        return row[matchKey];
      }
    }
    return defaultVal;
  }

  function parseExcelWorkbook(wb, sfSheetName, refSheetName) {
    // 1. Localizar e processar Referências / Artigos Scopus
    const actualRefSheet = refSheetName || findSheetName(wb.SheetNames, ['Referências', 'Referencias', 'Scopus 3 (Limpo)', 'Scopus 3', 'Scopus']);
    const wsScopus = actualRefSheet ? wb.Sheets[actualRefSheet] : null;
    const scopusJson = wsScopus ? XLSX.utils.sheet_to_json(wsScopus, { defval: '' }) : [];
    
    state.rawArticles = [];
    state.articlesMap.clear();

    scopusJson.forEach(row => {
      const rawId = getRowVal(row, ['ID', 'ID Scopus']);
      const artId = parseInt(rawId, 10);
      if (!isNaN(artId)) {
        let doiLink = getRowVal(row, ['DOILINK', 'DOI Link', 'Link']);
        const doi = getRowVal(row, ['DOI', 'Doi']);
        if (!doiLink && doi) {
          doiLink = String(doi).startsWith('http') ? String(doi) : `https://doi.org/${doi}`;
        }

        const artObj = {
          id: artId,
          fonte: String(getRowVal(row, ['Fonte', 'Origem'], '')).trim(),
          authors: getRowVal(row, ['Authors', 'Autores'], 'Autores não informados'),
          title: getRowVal(row, ['Title', 'Título', 'Titulo', 'Título do Artigo'], 'Título não informado'),
          year: getRowVal(row, ['Year', 'Ano'], ''),
          sourceTitle: getRowVal(row, ['Source title', 'Periódico / Fonte', 'Periodico / Fonte', 'Source'], ''),
          doi: doi,
          doiLink: doiLink,
          abstract: getRowVal(row, ['Abstract', 'Resumo', 'Resumo (Abstract)'], ''),
          authorKeywords: getRowVal(row, ['Author Keywords', 'Palavras-chave do Autor', 'Keywords'], ''),
          indexKeywords: getRowVal(row, ['Index Keywords', 'Palavras-chave Indexadas'], ''),
          documentType: getRowVal(row, ['Document Type', 'Tipo de Documento'], ''),
          associatedSFs: []
        };
        state.rawArticles.push(artObj);
        state.articlesMap.set(artId, artObj);
      }
    });

    // 2. Localizar e processar Sementes de Futuro
    const actualSfSheet = sfSheetName || findSheetName(wb.SheetNames, ['Sementes', 'SF (Scopus)', 'SF']);
    const wsSF = actualSfSheet ? wb.Sheets[actualSfSheet] : null;
    const sfJson = wsSF ? XLSX.utils.sheet_to_json(wsSF, { defval: '' }) : [];

    state.rawSementes = [];
    sfJson.forEach(row => {
      const sfId = String(getRowVal(row, ['ID da SF', 'ID SF', 'ID']) || '').trim();
      if (!sfId) return;

      // Parse Associated Article IDs
      const rawArtIds = String(getRowVal(row, ['Artigos Scopus Associados (IDs)', 'Artigos Associados (IDs)', 'Artigos Associados', 'Artigos Scopus', 'Artigos (IDs)']) || '');
      const artIds = [];
      rawArtIds.replace(/;/g, ',').split(',').forEach(p => {
        const num = parseInt(p.trim(), 10);
        if (!isNaN(num)) {
          artIds.push(num);
          const art = state.articlesMap.get(num);
          if (art && !art.associatedSFs.includes(sfId)) {
            art.associatedSFs.push(sfId);
          }
        }
      });

      // Parse PESTEL
      const rawPestel = String(getRowVal(row, ['PESTEL', 'Dimensões PESTEL', 'Dimensoes PESTEL', 'Dimensões']) || '');
      const pestelList = rawPestel.replace(/;/g, ',').split(',').map(d => d.trim()).filter(Boolean);

      // Reviewed & status parsing
      const rawRev = getRowVal(row, ['Revisada', 'Status Revisão', 'Revisado', 'Status'], 0);
      const rawRevStr = String(rawRev || '').trim();
      const statusLower = rawRevStr.toLowerCase();

      // Vínculo (Nova coluna dedicada)
      const rawVinculo = String(getRowVal(row, ['Vínculo', 'Vinculo', 'Vínculo SF', 'Vinculo SF', 'SFs Relacionadas', 'SF Relacionada', 'Vínculos'], '')).trim();
      const vinculoList = (rawVinculo && rawVinculo.toUpperCase() !== 'N/A' && rawVinculo.toUpperCase() !== 'NAN')
        ? rawVinculo.split(/[;,]+/).map(v => v.trim()).filter(Boolean)
        : [];

      // Justificativa
      const rawJust = String(getRowVal(row, ['Justificativa', 'Justificativa Metodológica', 'Motivo'], '')).trim();
      const justVal = (rawJust.toUpperCase() === 'N/A' || rawJust.toUpperCase() === 'NAN') ? '' : rawJust;

      let statusTipo = 'pendente';
      let fusaoId = '';
      let divisaoIds = [];
      let isRev = false;

      if (statusLower.includes('divid') || statusLower.includes('desmembr') || statusLower.includes('separ') || statusLower.includes('divis')) {
        statusTipo = 'dividida';
        isRev = true;
        if (vinculoList.length > 0) {
          divisaoIds = vinculoList.filter(d => d.toUpperCase() !== sfId.toUpperCase());
        } else {
          const matchParen = rawRevStr.match(/\((.*?)\)/);
          const searchTarget = matchParen ? matchParen[1] : rawRevStr;
          const extracted = searchTarget.match(/(SF[A-Za-z0-9]+)/gi) || [];
          divisaoIds = extracted.filter(d => d.toUpperCase() !== sfId.toUpperCase());
        }
      } else if (statusLower.includes('mescl') || statusLower.includes('fus') || statusLower.includes('fundid')) {
        statusTipo = 'mesclada';
        isRev = true;
        if (vinculoList.length > 0) {
          fusaoId = vinculoList[0];
        } else {
          const matchParen = rawRevStr.match(/\((.*?)\)/);
          if (matchParen) {
            fusaoId = matchParen[1].trim();
          } else {
            const matchSf = rawRevStr.match(/(SF[A-Z0-9]+)/i);
            if (matchSf) fusaoId = matchSf[1].trim();
          }
        }
      } else if (statusLower.includes('descart')) {
        statusTipo = 'descartada';
        isRev = true;
      } else if (['ok', '1', 'sim', 'true', 's', 'revisada', 'revisado', 'aprovada', 'aprovado'].includes(statusLower) || statusLower.includes('ok')) {
        statusTipo = 'ok';
        isRev = true;
      } else {
        statusTipo = 'pendente';
        isRev = false;
      }

      let statusLabel = 'Em revisão';
      if (statusTipo === 'ok') {
        statusLabel = 'Revisada (OK)';
      } else if (statusTipo === 'mesclada') {
        statusLabel = fusaoId ? `Revisada (Mesclada: ${fusaoId})` : 'Revisada (Mesclada)';
      } else if (statusTipo === 'descartada') {
        statusLabel = 'Revisada (Descartada)';
      } else if (statusTipo === 'dividida') {
        statusLabel = divisaoIds.length ? `Revisada (Dividida: ${divisaoIds.join(', ')})` : 'Revisada (Dividida)';
      }

      // Indicadores
      const rawInd = String(getRowVal(row, ['Indicadores', 'Indicadores Prospectivos'], '')).trim();
      const indVal = (rawInd.toUpperCase() === 'N/A' || rawInd.toUpperCase() === 'NAN') ? '' : rawInd;

      // Qtd Artigos
      const rawQtd = getRowVal(row, ['Qtd Artigos', 'Quantidade de Artigos'], 0);
      const parsedQtd = parseInt(rawQtd, 10);

      state.rawSementes.push({
        id: sfId,
        fonte: String(getRowVal(row, ['Fonte', 'Origem'], '')).trim(),
        topico: getRowVal(row, ['Tópico', 'Tópico Mais Amplo Padronizado', 'Topico'], 'Geral'),
        titulo: getRowVal(row, ['Título SF', 'Título da Semente de Futuro', 'Titulo SF', 'Título', 'Titulo'], 'Sem título'),
        descricao: getRowVal(row, ['Descrição (Ref N)', 'Descrição Global do Fenômeno (com Citações Numéricas)', 'Descricao (Ref N)', 'Descrição', 'Descricao'], ''),
        descricaoApa: getRowVal(row, ['Descrição (Ref APA)', 'Descricao (Ref APA)'], ''),
        classificacao: getRowVal(row, ['Classificação', 'Classificação Prospectiva', 'Classificacao'], 'Não classificada'),
        pestel: pestelList,
        eixo: getRowVal(row, ['Eixo', 'Eixo CSED', 'Eixo Temático'], 'Não categorizado'),
        qtdArtigos: artIds.length || (isNaN(parsedQtd) ? 0 : parsedQtd),
        artigosIds: artIds,
        referenciasResumidas: getRowVal(row, ['Referências', 'Referências Bibliográficas Resumidas (Autores, Ano, Título)', 'Referencias'], ''),
        implicacoes: getRowVal(row, ['Implicações', 'Conexão / Implicações Estratégicas para o Ensino Superior Privado até 2040', 'Implicacoes'], ''),
        indicadores: indVal,
        revisada: isRev,
        statusTipo: statusTipo,
        statusRevisao: (rawRevStr && rawRevStr !== '0') ? rawRevStr : 'Em revisão',
        statusLabel: statusLabel,
        justificativa: justVal,
        fusaoId: fusaoId,
        divisaoIds: divisaoIds,
        origemDivisao: '',
        vinculo: vinculoList,
        incorporadas: []
      });
    });

    // Mapeamento reverso de incorporações (ex: SFS02 incorpora SF22) e divisões (ex: SF31a/b originam de SF31)
    const sfMap = new Map();
    state.rawSementes.forEach(s => sfMap.set(s.id.toUpperCase(), s));

    state.rawSementes.forEach(s => {
      if (s.fusaoId && sfMap.has(s.fusaoId.toUpperCase())) {
        const target = sfMap.get(s.fusaoId.toUpperCase());
        if (!target.incorporadas.includes(s.id)) {
          target.incorporadas.push(s.id);
        }
      }
      if (s.divisaoIds && s.divisaoIds.length > 0) {
        s.divisaoIds = s.divisaoIds.map(d => sfMap.has(d.toUpperCase()) ? sfMap.get(d.toUpperCase()).id : d);
        s.divisaoIds.forEach(childId => {
          if (sfMap.has(childId.toUpperCase())) {
            const child = sfMap.get(childId.toUpperCase());
            child.origemDivisao = s.id;
          }
        });
      }
      if (s.vinculo && s.vinculo.length > 0) {
        s.vinculo.forEach(v => {
          if (sfMap.has(v.toUpperCase())) {
            const related = sfMap.get(v.toUpperCase());
            if (related.statusTipo === 'dividida' && !s.origemDivisao) {
              s.origemDivisao = related.id;
              if (!related.divisaoIds.includes(s.id)) {
                related.divisaoIds.push(s.id);
              }
            } else if (related.statusTipo === 'mesclada' && !s.incorporadas.includes(related.id)) {
              s.incorporadas.push(related.id);
            }
          }
        });
      }
    });

    // Sort SFs numerically & alphanumeric (ex: SF01..SF31, SF31a, SF31b..SF62, SFQ, SFS)
    state.rawSementes.sort((a, b) => {
      const catA = a.id.startsWith('SF') && !a.id.startsWith('SFQ') && !a.id.startsWith('SFS') ? 0 : (a.id.startsWith('SFQ') ? 1 : 2);
      const catB = b.id.startsWith('SF') && !b.id.startsWith('SFQ') && !b.id.startsWith('SFS') ? 0 : (b.id.startsWith('SFQ') ? 1 : 2);
      if (catA !== catB) return catA - catB;
      const numA = parseInt(a.id.replace(/\D/g, ''), 10) || 999;
      const numB = parseInt(b.id.replace(/\D/g, ''), 10) || 999;
      if (numA !== numB) return numA - numB;
      return a.id.localeCompare(b.id);
    });
  }

  // ========================================================================
  // ========================================================================
  // CHECKBOX FILTERS & COLLAPSIBLE SECTIONS GENERATION
  // ========================================================================
  function buildFilterCheckboxLists() {
    // 1. Status de Revisão
    const containerStatus = document.getElementById('filter-status-options');
    if (containerStatus) {
      containerStatus.innerHTML = '';
      const statusOptions = [
        { val: 'pendente', label: 'Em revisão', dotClass: 'filter-dot-status-pendente', icon: 'fa-regular fa-clock' },
        { val: 'ok', label: 'Revisada (OK)', dotClass: 'filter-dot-status-ok' },
        { val: 'dividida', label: 'Revisada (Dividida)', dotClass: 'filter-dot-status-dividida' },
        { val: 'mesclada', label: 'Revisada (Mesclada)', dotClass: 'filter-dot-status-mesclada' },
        { val: 'descartada', label: 'Revisada (Descartada)', dotClass: 'filter-dot-status-descartada' }
      ];
      statusOptions.forEach(opt => {
        const count = state.rawSementes.filter(s => s.statusTipo === opt.val).length;
        const item = createCheckboxItem('filter-status', opt.val, opt.label, count, opt.dotClass, opt.icon);
        containerStatus.appendChild(item);
      });
    }

    // 2. Eixos Estratégicos
    const containerEixo = document.getElementById('filter-eixo-options');
    if (containerEixo) {
      containerEixo.innerHTML = '';
      const eixos = [...new Set(state.rawSementes.map(s => s.eixo).filter(Boolean))].sort();
      eixos.forEach((e, idx) => {
        const count = state.rawSementes.filter(s => s.eixo === e).length;
        const palette = PALETTE[state.theme] || PALETTE.dark;
        const color = palette.eixos[idx % palette.eixos.length];
        const item = createCheckboxItem('filter-eixo', e, e, count, null, null, color);
        containerEixo.appendChild(item);
      });
    }

    // 3. Tipologia Prospectiva (Classificação)
    const containerClass = document.getElementById('filter-classificacao-options');
    if (containerClass) {
      containerClass.innerHTML = '';
      const classificacoes = [...new Set(state.rawSementes.map(s => s.classificacao).filter(Boolean))].sort();
      classificacoes.forEach(c => {
        const count = state.rawSementes.filter(s => s.classificacao === c).length;
        const dotClass = 'filter-dot-' + cleanBadgeClass(c);
        const item = createCheckboxItem('filter-classificacao', c, c, count, dotClass);
        containerClass.appendChild(item);
      });
    }

    // 4. Dimensões PESTEL
    const containerPestel = document.getElementById('filter-pestel-options');
    if (containerPestel) {
      containerPestel.innerHTML = '';
      const allPestel = ['Político', 'Econômico', 'Social', 'Tecnológico', 'Ambiental', 'Legal'];
      allPestel.forEach(p => {
        const count = state.rawSementes.filter(s => s.pestel.includes(p)).length;
        const dotClass = 'filter-dot-' + cleanBadgeClass(p);
        const item = createCheckboxItem('filter-pestel', p, p, count, dotClass);
        containerPestel.appendChild(item);
      });
    }

    // 5. Tópicos Padronizados
    const containerTopico = document.getElementById('filter-topico-options');
    if (containerTopico) {
      containerTopico.innerHTML = '';
      const topicos = [...new Set(state.rawSementes.map(s => s.topico).filter(Boolean))].sort();
      topicos.forEach(t => {
        const count = state.rawSementes.filter(s => s.topico === t).length;
        const item = createCheckboxItem('filter-topico', t, t, count);
        containerTopico.appendChild(item);
      });
    }

    // 6. Scopus Years
    const years = [...new Set(state.rawArticles.map(a => a.year).filter(Boolean))].sort((a, b) => b - a);
    const selectYear = document.getElementById('scopus-filter-year');
    if (selectYear && selectYear.options.length <= 1) {
      years.forEach(y => {
        const opt = document.createElement('option');
        opt.value = y;
        opt.textContent = y;
        selectYear.appendChild(opt);
      });
    }

    // Scopus Relation Options
    const selectRel = document.getElementById('scopus-filter-relation');
    if (selectRel) {
      const linkedCount = state.rawArticles.filter(a => a.associatedSFs.length > 0).length;
      const unlinkedCount = state.rawArticles.length - linkedCount;
      selectRel.innerHTML = `
        <option value="">Todos os Artigos (${state.rawArticles.length})</option>
        <option value="linked" selected>Apenas com SFs Vinculadas (${linkedCount})</option>
        <option value="unlinked">Sem SFs Vinculadas (${unlinkedCount})</option>
      `;
    }
  }

  function createCheckboxItem(groupName, value, labelText, count, dotClass, iconClass, inlineColor) {
    const label = document.createElement('label');
    label.className = 'filter-checkbox-label';
    label.title = `${labelText} (${count} sementes)`;

    let visual = '';
    if (dotClass) {
      visual = `<span class="filter-dot-indicator ${dotClass}"></span>`;
    } else if (inlineColor) {
      visual = `<span class="filter-dot-indicator" style="background-color: ${inlineColor};"></span>`;
    } else if (iconClass) {
      visual = `<i class="${iconClass}" style="font-size: 0.72rem; width: 14px; text-align: center; color: var(--text-muted);"></i>`;
    }

    label.innerHTML = `
      <input type="checkbox" class="filter-checkbox" name="${groupName}" value="${escapeHtml(value)}">
      ${visual}
      <span class="filter-label-text">${escapeHtml(labelText)}</span>
      <span class="filter-item-count">${count}</span>
    `;

    return label;
  }

  function syncFilterStateFromDOM() {
    function getChecked(name) {
      return Array.from(document.querySelectorAll(`input[name="${name}"]:checked`)).map(cb => cb.value);
    }

    state.filters.status = getChecked('filter-status');
    state.filters.eixos = getChecked('filter-eixo');
    state.filters.classificacoes = getChecked('filter-classificacao');
    state.filters.pestels = getChecked('filter-pestel');
    state.filters.topicos = getChecked('filter-topico');

    // Update active count badges on headers
    updateActiveCountBadge('status', state.filters.status.length);
    updateActiveCountBadge('eixo', state.filters.eixos.length);
    updateActiveCountBadge('classificacao', state.filters.classificacoes.length);
    updateActiveCountBadge('pestel', state.filters.pestels.length);
    updateActiveCountBadge('topico', state.filters.topicos.length);

    // Update .is-checked class on labels
    document.querySelectorAll('.filter-checkbox').forEach(cb => {
      const parentLabel = cb.closest('.filter-checkbox-label');
      if (parentLabel) {
        if (cb.checked) {
          parentLabel.classList.add('is-checked');
        } else {
          parentLabel.classList.remove('is-checked');
        }
      }
    });
  }

  function updateActiveCountBadge(group, count) {
    const badge = document.getElementById(`active-count-${group}`);
    if (badge) {
      if (count > 0) {
        badge.textContent = count;
        badge.style.display = 'inline-flex';
      } else {
        badge.textContent = '0';
        badge.style.display = 'none';
      }
    }
  }

  function setupCollapsibleFilterSections() {
    document.querySelectorAll('.filter-section-header').forEach(header => {
      header.addEventListener('click', (e) => {
        e.preventDefault();
        const section = header.closest('.sidebar-filter-section');
        if (section) {
          const isCollapsed = section.classList.toggle('collapsed');
          header.setAttribute('aria-expanded', !isCollapsed);
        }
      });
    });
  }

  // ========================================================================
  // FILTERING LOGIC: SEMENTES DE FUTURO (MULTI-SELECT)
  // ========================================================================
  function applySfFilters() {
    const q = removeAccents(state.filters.search.trim());
    const fStatus = state.filters.status;
    const fEixos = state.filters.eixos;
    const fClass = state.filters.classificacoes;
    const fPestels = state.filters.pestels;
    const fTopicos = state.filters.topicos;

    state.filteredSementes = state.rawSementes.filter(sf => {
      // 1. Text Search (ID, Title, Description, Implications, Topic, Indicators, Justification, Lineage, or author mentions)
      if (q) {
        const matches = 
          removeAccents(sf.id).includes(q) ||
          removeAccents(sf.titulo).includes(q) ||
          removeAccents(sf.descricao).includes(q) ||
          removeAccents(sf.implicacoes).includes(q) ||
          removeAccents(sf.topico).includes(q) ||
          removeAccents(sf.indicadores || '').includes(q) ||
          removeAccents(sf.justificativa || '').includes(q) ||
          removeAccents(sf.origemDivisao || '').includes(q) ||
          (sf.divisaoIds && sf.divisaoIds.some(d => removeAccents(d).includes(q))) ||
          removeAccents(sf.referenciasResumidas).includes(q);
        if (!matches) return false;
      }

      // 2. Status Filter (Multi-select OR: sf matches if in selected status list)
      if (fStatus.length > 0 && !fStatus.includes(sf.statusTipo)) return false;

      // 3. Eixo Filter (Multi-select OR: sf matches if its eixo is in selected eixos)
      if (fEixos.length > 0 && !fEixos.includes(sf.eixo)) return false;

      // 4. Classificação Filter (Multi-select OR: sf matches if in selected classificacoes)
      if (fClass.length > 0 && !fClass.includes(sf.classificacao)) return false;

      // 5. PESTEL Filter (Multi-select: sf contains AT LEAST ONE selected dimension)
      if (fPestels.length > 0 && !sf.pestel.some(p => fPestels.includes(p))) return false;

      // 6. Tópico Filter (Multi-select OR: sf matches if in selected topicos)
      if (fTopicos.length > 0 && !fTopicos.includes(sf.topico)) return false;

      return true;
    });

    // Sort Filtered Sementes
    sortFilteredSementes();

    // Check if currently selected SF is still in filtered list
    if (state.selectedSfId) {
      const stillExists = state.filteredSementes.some(s => s.id === state.selectedSfId);
      if (!stillExists) {
        deselectSf();
      }
    }

    // Render active tags, counts and list
    renderActiveFilterTags();
    renderSfInboxList();
    updateChartsData();
  }

  function sortFilteredSementes() {
    const sortVal = state.listSort;

    state.filteredSementes.sort((a, b) => {
      if (sortVal === 'id-asc') {
        const catA = a.id.startsWith('SF') && !a.id.startsWith('SFQ') && !a.id.startsWith('SFS') ? 0 : (a.id.startsWith('SFQ') ? 1 : 2);
        const catB = b.id.startsWith('SF') && !b.id.startsWith('SFQ') && !b.id.startsWith('SFS') ? 0 : (b.id.startsWith('SFQ') ? 1 : 2);
        if (catA !== catB) return catA - catB;
        const nA = parseInt(a.id.replace(/\D/g, ''), 10) || 0;
        const nB = parseInt(b.id.replace(/\D/g, ''), 10) || 0;
        if (nA !== nB) return nA - nB;
        return a.id.localeCompare(b.id);
      }
      if (sortVal === 'id-desc') {
        const catA = a.id.startsWith('SF') && !a.id.startsWith('SFQ') && !a.id.startsWith('SFS') ? 0 : (a.id.startsWith('SFQ') ? 1 : 2);
        const catB = b.id.startsWith('SF') && !b.id.startsWith('SFQ') && !b.id.startsWith('SFS') ? 0 : (b.id.startsWith('SFQ') ? 1 : 2);
        if (catA !== catB) return catB - catA;
        const nA = parseInt(a.id.replace(/\D/g, ''), 10) || 0;
        const nB = parseInt(b.id.replace(/\D/g, ''), 10) || 0;
        if (nA !== nB) return nB - nA;
        return b.id.localeCompare(a.id);
      }
      if (sortVal === 'artigos-desc') {
        return b.qtdArtigos - a.qtdArtigos;
      }
      if (sortVal === 'titulo-asc') {
        return a.titulo.localeCompare(b.titulo, 'pt-BR');
      }
      return 0;
    });
  }

  // ========================================================================
  // RENDERING: COLUNA 2 — LISTA DE SEMENTES SELECIONÁVEIS (INBOX)
  // ========================================================================
  function renderSfInboxList() {
    const container = document.getElementById('sf-inbox-list');
    container.innerHTML = '';

    const total = state.rawSementes.length;
    const filteredCount = state.filteredSementes.length;

    // Update Counter Badges
    document.getElementById('sf-filtered-count').textContent = filteredCount;
    document.getElementById('sf-total-count').textContent = total;
    document.getElementById('list-count-badge').textContent = filteredCount;

    if (filteredCount === 0) {
      container.innerHTML = `
        <div style="padding: 36px 20px; text-align: center; color: var(--text-muted);">
          <i class="fa-solid fa-seedling" style="font-size: 1.8rem; margin-bottom: 10px; opacity: 0.4;"></i>
          <p style="font-size: 0.88rem; font-weight: 600; color: var(--text-primary);">Nenhuma semente encontrada</p>
          <p style="font-size: 0.78rem; margin-top: 4px;">Tente ajustar ou limpar os filtros na coluna à esquerda.</p>
        </div>
      `;
      deselectSf();
      return;
    }

    state.filteredSementes.forEach(sf => {
      const item = document.createElement('div');
      item.className = 'inbox-item' + (sf.statusTipo ? ' item-' + sf.statusTipo : '');
      item.setAttribute('data-sf-id', sf.id);
      if (state.selectedSfId === sf.id) {
        item.classList.add('active');
      }

      const typClass = cleanBadgeClass(sf.classificacao);

      let statusBadge = '';
      if (sf.statusTipo === 'mesclada') {
        const justTitle = sf.justificativa ? `Justificativa: ${sf.justificativa}` : `Mesclada com ${sf.fusaoId || ''}`;
        statusBadge = `<span class="badge-status-subtle badge-status-fundida" title="${escapeHtml(justTitle)}"><i class="fa-solid fa-code-merge"></i> Mesclada</span>`;
      } else if (sf.statusTipo === 'dividida') {
        const justTitle = sf.justificativa ? `Justificativa: ${sf.justificativa}` : `Dividida em ${sf.divisaoIds ? sf.divisaoIds.join(', ') : ''}`;
        statusBadge = `<span class="badge-status-subtle badge-status-dividida" title="${escapeHtml(justTitle)}"><i class="fa-solid fa-scissors"></i> Dividida</span>`;
      } else if (sf.statusTipo === 'descartada') {
        const justTitle = sf.justificativa ? `Justificativa: ${sf.justificativa}` : 'Descartada na revisão';
        statusBadge = `<span class="badge-status-subtle badge-status-descartada" title="${escapeHtml(justTitle)}"><i class="fa-solid fa-ban"></i> Descartada</span>`;
      } else if (sf.statusTipo === 'ok') {
        const justTitle = sf.justificativa ? `Justificativa: ${sf.justificativa}` : 'Revisada (OK)';
        const childBadge = sf.origemDivisao ? `<span class="badge-status-subtle badge-status-filha" title="Desmembrada a partir de ${escapeHtml(sf.origemDivisao)}"><i class="fa-solid fa-code-branch"></i> da ${escapeHtml(sf.origemDivisao)}</span> ` : '';
        statusBadge = `${childBadge}<span class="icon-reviewed-subtle" title="${escapeHtml(justTitle)}"><i class="fa-solid fa-circle-check"></i></span>`;
      } else {
        statusBadge = `<span class="icon-pending-subtle" title="Em revisão"><i class="fa-regular fa-clock"></i></span>`;
      }

      const pestelChips = sf.pestel.map(p => {
        return `<span class="pestel-chip pestel-chip-neutral">${escapeHtml(p)}</span>`;
      }).join('');

      item.innerHTML = `
        <div class="inbox-item-title">${escapeHtml(sf.titulo)}</div>

        <div class="inbox-item-topic" title="${escapeHtml(sf.topico)}">
          <i class="fa-solid fa-tag"></i> <span>${escapeHtml(sf.topico)}</span>
        </div>

        <div class="inbox-item-meta-row">
          <span class="badge-id-neutral">${sf.id}</span>
          <span class="badge-typology ${typClass}">${escapeHtml(sf.classificacao)}</span>
          ${statusBadge}
          <span class="inbox-art-count" title="${sf.qtdArtigos} referências vinculadas">
            <i class="fa-solid fa-book-open"></i> ${sf.qtdArtigos} ref.
          </span>
        </div>

        <div class="inbox-item-pestel-row">
          ${pestelChips}
        </div>
      `;

      item.addEventListener('click', () => {
        selectSf(sf.id);
      });

      container.appendChild(item);
    });
  }

  // ========================================================================
  // SELEÇÃO & ÁREA DE LEITURA (COLUNA 3 — READING PANE)
  // ========================================================================
  function selectSf(sfId) {
    // Se a SF selecionada não estiver visível na filtragem atual, limpa os filtros para torná-la acessível
    const inFiltered = state.filteredSementes.some(s => s.id === sfId);
    if (!inFiltered && state.rawSementes.some(s => s.id === sfId)) {
      resetSfFilters();
    }

    state.selectedSfId = sfId;

    // Atualiza classe ativa na lista de itens
    document.querySelectorAll('.inbox-item').forEach(el => {
      if (el.getAttribute('data-sf-id') === sfId) {
        el.classList.add('active');
        el.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
      } else {
        el.classList.remove('active');
      }
    });

    const sf = state.rawSementes.find(s => s.id === sfId);
    if (!sf) return;

    // Oculta empty state e exibe o dossiê de leitura
    const emptyState = document.getElementById('reading-empty-state');
    const dossier = document.getElementById('reading-dossier');
    emptyState.style.display = 'none';
    dossier.style.display = 'flex';

    // Rola para o topo do dossiê
    const readingPane = document.getElementById('sf-reading-pane');
    readingPane.scrollTop = 0;

    const typClass = cleanBadgeClass(sf.classificacao);

    let reviewedBadge = '';
    if (sf.statusTipo === 'mesclada') {
      reviewedBadge = `<span class="badge-status-dossier badge-dossier-fundida"><i class="fa-solid fa-code-merge"></i> Revisada (Mesclada)${sf.fusaoId ? ' → ' + escapeHtml(sf.fusaoId) : ''}</span>`;
    } else if (sf.statusTipo === 'dividida') {
      reviewedBadge = `<span class="badge-status-dossier badge-dossier-dividida"><i class="fa-solid fa-scissors"></i> Revisada (Dividida)${sf.divisaoIds && sf.divisaoIds.length ? ' → ' + escapeHtml(sf.divisaoIds.join(', ')) : ''}</span>`;
    } else if (sf.statusTipo === 'descartada') {
      reviewedBadge = `<span class="badge-status-dossier badge-dossier-descartada"><i class="fa-solid fa-ban"></i> Revisada (Descartada)</span>`;
    } else if (sf.statusTipo === 'ok') {
      const childTag = sf.origemDivisao ? ` <span class="badge-child-tag" title="Originada da divisão de ${escapeHtml(sf.origemDivisao)}">(da ${escapeHtml(sf.origemDivisao)})</span>` : '';
      reviewedBadge = `<span class="badge-reviewed" style="font-size: 0.74rem; padding: 3px 8px;"><i class="fa-solid fa-check-circle"></i> Revisada (OK)${childTag}</span>`;
    } else {
      reviewedBadge = `<span class="badge-pending" style="font-size: 0.74rem; padding: 3px 8px;"><i class="fa-regular fa-clock"></i> Em revisão</span>`;
    }

    const pestelChips = sf.pestel.map(p => {
      const pClass = cleanBadgeClass(p);
      return `<span class="pestel-chip ${pClass}" style="font-size: 0.74rem; padding: 2px 8px;">${p}</span>`;
    }).join('');

    const descFormatted = highlightNumericCitations(sf.descricao);

    // Banner contextual de status (Mesclada, Dividida, Filha de Divisão, Descartada, Incorporadora)
    let statusBannerHtml = '';
    if (sf.statusTipo === 'mesclada') {
      const targetSf = sf.fusaoId ? state.rawSementes.find(s => s.id === sf.fusaoId) : null;
      statusBannerHtml = `
        <div class="dossier-status-banner banner-fundida">
          <div class="status-banner-icon"><i class="fa-solid fa-code-merge"></i></div>
          <div class="status-banner-body">
            <div class="status-banner-title">Semente Revisada (Mesclada)</div>
            ${sf.justificativa ? `
              <div class="status-banner-justificativa">
                <i class="fa-solid fa-circle-info"></i> <span><strong>Justificativa metodológica:</strong> ${escapeHtml(sf.justificativa)}</span>
              </div>
            ` : ''}
            <div class="status-banner-text">
              Esta semente foi integrada e sintetizada na semente 
              <a href="javascript:void(0)" class="banner-sf-link" data-jump-sf="${escapeHtml(sf.fusaoId)}">
                <strong>${escapeHtml(sf.fusaoId)}</strong>${targetSf ? ' — ' + escapeHtml(targetSf.titulo) : ''}
              </a>. Seus dados e referências originais permanecem arquivados aqui para fins de rastreabilidade.
            </div>
          </div>
          ${sf.fusaoId ? `<button class="btn-banner-jump" data-jump-sf="${escapeHtml(sf.fusaoId)}"><i class="fa-solid fa-arrow-right"></i> Ver ${escapeHtml(sf.fusaoId)}</button>` : ''}
        </div>
      `;
    } else if (sf.statusTipo === 'dividida') {
      const childrenLinks = (sf.divisaoIds || []).map(childId => {
        const childSf = state.rawSementes.find(s => s.id === childId);
        return `<a href="javascript:void(0)" class="banner-sf-link" data-jump-sf="${escapeHtml(childId)}"><strong>${escapeHtml(childId)}</strong>${childSf ? ' — ' + escapeHtml(childSf.titulo) : ''}</a>`;
      }).join(' e ');
      const jumpButtons = (sf.divisaoIds || []).map(childId => {
        return `<button class="btn-banner-jump" data-jump-sf="${escapeHtml(childId)}"><i class="fa-solid fa-arrow-right"></i> Ver ${escapeHtml(childId)}</button>`;
      }).join(' ');

      statusBannerHtml = `
        <div class="dossier-status-banner banner-dividida">
          <div class="status-banner-icon"><i class="fa-solid fa-scissors"></i></div>
          <div class="status-banner-body">
            <div class="status-banner-title">Semente Revisada (Dividida)</div>
            ${sf.justificativa ? `
              <div class="status-banner-justificativa">
                <i class="fa-solid fa-circle-info"></i> <span><strong>Justificativa metodológica:</strong> ${escapeHtml(sf.justificativa)}</span>
              </div>
            ` : ''}
            <div class="status-banner-text">
              Esta semente foi desmembrada em novas sementes autônomas na revisão prospectiva: ${childrenLinks}. Seus dados originais e referências associadas permanecem arquivados nesta plataforma para histórico científico e rastreabilidade metodológica.
            </div>
          </div>
          ${jumpButtons ? `<div class="banner-jump-group">${jumpButtons}</div>` : ''}
        </div>
      `;
    } else if (sf.origemDivisao) {
      const parentSf = state.rawSementes.find(s => s.id === sf.origemDivisao);
      statusBannerHtml = `
        <div class="dossier-status-banner banner-origem-divisao">
          <div class="status-banner-icon"><i class="fa-solid fa-code-branch"></i></div>
          <div class="status-banner-body">
            <div class="status-banner-title">Semente Resultante de Divisão Metodológica</div>
            <div class="status-banner-text">
              Esta semente resultou do desmembramento e aprofundamento analítico da semente 
              <a href="javascript:void(0)" class="banner-sf-link" data-jump-sf="${escapeHtml(sf.origemDivisao)}">
                <strong>${escapeHtml(sf.origemDivisao)}</strong>${parentSf ? ' — ' + escapeHtml(parentSf.titulo) : ''}
              </a>.
            </div>
          </div>
          <button class="btn-banner-jump" data-jump-sf="${escapeHtml(sf.origemDivisao)}"><i class="fa-solid fa-arrow-left"></i> Ver semente original (${escapeHtml(sf.origemDivisao)})</button>
        </div>
      `;
    } else if (sf.statusTipo === 'descartada') {
      statusBannerHtml = `
        <div class="dossier-status-banner banner-descartada">
          <div class="status-banner-icon"><i class="fa-solid fa-ban"></i></div>
          <div class="status-banner-body">
            <div class="status-banner-title">Semente Revisada (Descartada)</div>
            ${sf.justificativa ? `
              <div class="status-banner-justificativa">
                <i class="fa-solid fa-circle-exclamation"></i> <span><strong>Justificativa metodológica:</strong> ${escapeHtml(sf.justificativa)}</span>
              </div>
            ` : ''}
            <div class="status-banner-text">
              Esta semente foi descontinuada na revisão metodológica de cenários. Seus artigos e metadados permanecem preservados nesta base para histórico e auditoria científica.
            </div>
          </div>
        </div>
      `;
    } else if (sf.incorporadas && sf.incorporadas.length > 0) {
      statusBannerHtml = `
        <div class="dossier-status-banner banner-incorporadora">
          <div class="status-banner-icon"><i class="fa-solid fa-object-group"></i></div>
          <div class="status-banner-body">
            <div class="status-banner-title">Síntese de Consolidação Metodológica</div>
            <div class="status-banner-text">
              Esta semente incorpora e expande os achados de: 
              ${sf.incorporadas.map(incId => {
                const incSf = state.rawSementes.find(s => s.id === incId);
                return `<a href="javascript:void(0)" class="banner-sf-link" data-jump-sf="${escapeHtml(incId)}" title="${incSf ? escapeHtml(incSf.titulo) : ''}"><strong>${escapeHtml(incId)}</strong></a>`;
              }).join(', ')}.
            </div>
          </div>
        </div>
      `;
    }

    // Renderiza bloco de Indicadores se houver
    let indicatorsHtml = '';
    const rawIndicadores = (sf.indicadores || '').trim();
    if (rawIndicadores && rawIndicadores.toUpperCase() !== 'N/A') {
      const indItems = rawIndicadores.split(';').map(i => i.trim()).filter(Boolean);
      const indItemsHtml = indItems.length > 0 
        ? indItems.map(item => `<li><i class="fa-solid fa-chart-line"></i> <span>${escapeHtml(item)}</span></li>`).join('')
        : `<li><i class="fa-solid fa-chart-line"></i> <span>${escapeHtml(rawIndicadores)}</span></li>`;

      indicatorsHtml = `
        <div class="dossier-indicators-box">
          <div class="indicators-box-header">
            <i class="fa-solid fa-chart-line"></i>
            <span>Indicadores de Monitoramento</span>
          </div>
          <ul class="indicators-list">
            ${indItemsHtml}
          </ul>
        </div>
      `;
    }

    // Renderiza a lista de artigos fundamentadores com acordeão inline de Abstract
    let articlesHtml = '';
    if (sf.artigosIds.length === 0) {
      articlesHtml = `
        <div style="padding: 16px; color: var(--text-muted); font-size: 0.85rem; text-align: center; background: var(--bg-surface-elevated); border-radius: var(--radius-md);">
          Nenhum artigo científico vinculado diretamente nesta semente.
        </div>
      `;
    } else {
      articlesHtml = sf.artigosIds.map(artId => {
        const art = state.articlesMap.get(artId);
        if (!art) {
          return `
            <div class="dossier-article-card">
              <span class="art-card-id">#ID ${artId}</span>
            </div>
          `;
        }

        const doiButton = art.doiLink 
          ? `<a href="${art.doiLink}" target="_blank" rel="noopener noreferrer" class="btn-dossier-action" title="Abrir DOI do artigo"><i class="fa-solid fa-arrow-up-right-from-square"></i> DOI</a>`
          : '';

        return `
          <div class="dossier-article-card" id="art-card-${art.id}">
            <div class="art-card-top">
              <div class="art-card-meta-left">
                <span class="art-card-id">#ID ${art.id}</span>
                <span>•</span>
                <span>${art.year || 's.d.'}</span>
                <span>•</span>
                <span class="art-card-source">${escapeHtml(art.sourceTitle)}</span>
              </div>
              <div class="art-card-actions">
                <button class="btn-art-inline-toggle" data-toggle-abstract="${art.id}">
                  <i class="fa-solid fa-file-lines"></i> <span>Resumo</span>
                </button>
                ${doiButton}
              </div>
            </div>

            <div class="art-card-title">${escapeHtml(art.title)}</div>
            <div class="art-card-authors">${escapeHtml(art.authors)}</div>

            <div class="art-inline-abstract" id="abstract-drawer-${art.id}">
              <strong>Resumo:</strong>
              <p style="margin-top: 6px;">${escapeHtml(art.abstract || 'Resumo não disponível na base bibliográfica.')}</p>
            </div>
          </div>
        `;
      }).join('');
    }

    dossier.innerHTML = `
      <!-- Barra Superior do Dossiê -->
      <div class="dossier-top-bar">
        <div class="dossier-meta-group">
          <span class="badge-id" style="font-size: 0.85rem; padding: 3px 8px;">${sf.id}</span>
          <span class="badge-eixo-lg" title="${escapeHtml(sf.eixo)}">${escapeHtml(sf.eixo)}</span>
          <span class="badge-typology ${typClass}">${sf.classificacao}</span>
          ${reviewedBadge}
        </div>
        <div class="dossier-actions">
          <button class="btn-dossier-action" id="btn-copy-dossier" title="Copiar resumo da semente">
            <i class="fa-regular fa-copy"></i> Copiar
          </button>
          <button class="btn-dossier-action" id="btn-close-reading" title="Fechar leitura">
            <i class="fa-solid fa-xmark"></i> Fechar
          </button>
        </div>
      </div>

      ${statusBannerHtml}

      <!-- Título Executivo -->
      <h2 class="dossier-title">${escapeHtml(sf.titulo)}</h2>

      <!-- Tópico Temático Padronizado -->
      <div class="dossier-topic-row">
        <i class="fa-solid fa-tag"></i> ${escapeHtml(sf.topico)}
      </div>

      <!-- Dimensões -->
      <div class="dossier-pestel-row">
        <span class="dossier-pestel-label">Dimensões:</span>
        <div class="dossier-pestel-chips">
          ${pestelChips}
        </div>
      </div>

      <!-- Seção 1: Descrição -->
      <div class="dossier-section">
        <div class="dossier-section-header">
          <div class="dossier-section-header-title">
            <i class="fa-solid fa-align-left"></i>
            <span>Descrição</span>
          </div>
        </div>
        <div class="text-narrative-box">
          ${descFormatted}
        </div>
        ${indicatorsHtml}
      </div>

      <!-- Seção 2: Implicações Estratégicas 2040 -->
      <div class="dossier-section">
        <div class="dossier-implications-box">
          <div class="implications-box-header">
            <i class="fa-solid fa-compass"></i>
            <span>Conexão e Implicações Estratégicas para o Ensino Superior Privado até 2040</span>
          </div>
          <div class="implications-box-content">
            ${escapeHtml(sf.implicacoes || 'Implicações estratégicas em consolidação pelo grupo de pesquisa.')}
          </div>
        </div>
      </div>

      <!-- Seção 3: Referências -->
      <div class="dossier-section">
        <div class="dossier-section-header">
          <div class="dossier-section-header-title">
            <i class="fa-solid fa-book-open"></i>
            <span>Referências</span>
          </div>
        </div>
        <div class="dossier-articles-list">
          ${articlesHtml}
        </div>
      </div>
    `;

    // Event listeners dos botões internos do dossiê
    document.getElementById('btn-close-reading').addEventListener('click', () => {
      deselectSf();
    });

    document.getElementById('btn-copy-dossier').addEventListener('click', () => {
      let copyText = `[${sf.id}] ${sf.titulo}\n\nStatus: ${sf.statusLabel || sf.statusRevisao} | Tipologia: ${sf.classificacao} | Eixo: ${sf.eixo}`;
      if (sf.justificativa) {
        copyText += `\nJustificativa: ${sf.justificativa}`;
      }
      if (sf.origemDivisao) {
        copyText += `\nOrigem: Desmembrada a partir de ${sf.origemDivisao}`;
      }
      if (sf.divisaoIds && sf.divisaoIds.length > 0) {
        copyText += `\nDivisões resultantes: ${sf.divisaoIds.join(', ')}`;
      }
      copyText += `\n\nDescrição:\n${sf.descricao}`;
      if (rawIndicadores && rawIndicadores.toUpperCase() !== 'N/A') {
        copyText += `\n\nIndicadores:\n${rawIndicadores}`;
      }
      copyText += `\n\nImplicações 2040:\n${sf.implicacoes}`;
      const btn = document.getElementById('btn-copy-dossier');
      copyTextToClipboard(copyText, () => {
        btn.innerHTML = `<i class="fa-solid fa-check"></i> Copiado!`;
        setTimeout(() => {
          btn.innerHTML = `<i class="fa-regular fa-copy"></i> Copiar`;
        }, 1800);
      });
    });

    // Navegação interativa para sementes fundidas/incorporadas
    dossier.querySelectorAll('[data-jump-sf]').forEach(el => {
      el.addEventListener('click', (e) => {
        e.preventDefault();
        const targetId = el.getAttribute('data-jump-sf');
        if (targetId) {
          selectSf(targetId);
        }
      });
    });

    // Navegação interativa nas citações numéricas do texto
    dossier.querySelectorAll('.citation-badge-link').forEach(link => {
      link.addEventListener('click', (e) => {
        e.preventDefault();
        const refId = link.getAttribute('data-art-ref');
        const targetCard = document.getElementById(`art-card-${refId}`);
        if (targetCard) {
          targetCard.scrollIntoView({ behavior: 'smooth', block: 'center' });
          targetCard.classList.remove('highlight-pulse');
          void targetCard.offsetWidth; // force reflow
          targetCard.classList.add('highlight-pulse');

          // Abre o abstract automaticamente caso esteja oculto
          const drawer = document.getElementById(`abstract-drawer-${refId}`);
          const btn = targetCard.querySelector(`[data-toggle-abstract="${refId}"]`);
          if (drawer && !drawer.classList.contains('visible')) {
            drawer.classList.add('visible');
            if (btn) btn.innerHTML = `<i class="fa-solid fa-chevron-up"></i> <span>Ocultar</span>`;
          }
        } else {
          // Se o artigo citado não estiver na lista direta desta SF, abre no modal Scopus
          const artNum = parseInt(refId, 10);
          if (!isNaN(artNum)) {
            openArticleModal(artNum);
          }
        }
      });
    });

    // Toggle inline abstract de cada artigo no dossiê
    dossier.querySelectorAll('[data-toggle-abstract]').forEach(btn => {
      btn.addEventListener('click', () => {
        const artId = btn.getAttribute('data-toggle-abstract');
        const drawer = document.getElementById(`abstract-drawer-${artId}`);
        if (drawer) {
          const isVisible = drawer.classList.contains('visible');
          if (isVisible) {
            drawer.classList.remove('visible');
            btn.innerHTML = `<i class="fa-solid fa-file-lines"></i> <span>Resumo</span>`;
          } else {
            drawer.classList.add('visible');
            btn.innerHTML = `<i class="fa-solid fa-chevron-up"></i> <span>Ocultar</span>`;
          }
        }
      });
    });
  }

  function deselectSf() {
    state.selectedSfId = null;
    document.querySelectorAll('.inbox-item').forEach(el => el.classList.remove('active'));
    document.getElementById('reading-empty-state').style.display = 'flex';
    document.getElementById('reading-dossier').style.display = 'none';
  }

  // ========================================================================
  // ACTIVE TAGS BAR
  // ========================================================================
  function uncheckFilter(groupName, value) {
    const checkboxes = document.querySelectorAll(`input[name="${groupName}"]`);
    checkboxes.forEach(cb => {
      if (cb.value === value) {
        cb.checked = false;
        const parentLabel = cb.closest('.filter-checkbox-label');
        if (parentLabel) parentLabel.classList.remove('is-checked');
      }
    });
    syncFilterStateFromDOM();
    applySfFilters();
  }

  function renderActiveFilterTags() {
    const container = document.getElementById('active-tags-container');
    if (!container) return;
    container.innerHTML = '';

    const addTag = (label, removeCallback) => {
      const tag = document.createElement('span');
      tag.className = 'filter-tag';
      tag.innerHTML = `
        <span>${escapeHtml(label)}</span>
        <i class="fa-solid fa-xmark filter-tag-remove"></i>
      `;
      tag.querySelector('.filter-tag-remove').addEventListener('click', removeCallback);
      container.appendChild(tag);
    };

    if (state.filters.search) {
      addTag(`"${state.filters.search}"`, () => {
        state.filters.search = '';
        const searchInput = document.getElementById('sf-search-input');
        if (searchInput) searchInput.value = '';
        const searchClear = document.getElementById('sf-search-clear');
        if (searchClear) searchClear.style.display = 'none';
        applySfFilters();
      });
    }

    // Status Tags
    const statusLabelMap = {
      'pendente': 'Status: Em revisão',
      'ok': 'Status: Revisada (OK)',
      'dividida': 'Status: Revisada (Dividida)',
      'mesclada': 'Status: Revisada (Mesclada)',
      'descartada': 'Status: Revisada (Descartada)'
    };
    (state.filters.status || []).forEach(val => {
      const label = statusLabelMap[val] || `Status: ${val}`;
      addTag(label, () => uncheckFilter('filter-status', val));
    });

    // Eixos Tags
    (state.filters.eixos || []).forEach(e => {
      const shortEixo = e.length > 25 ? e.substring(0, 23) + '...' : e;
      addTag(`Eixo: ${shortEixo}`, () => uncheckFilter('filter-eixo', e));
    });

    // Classificação Tags
    (state.filters.classificacoes || []).forEach(c => {
      addTag(`${c}`, () => uncheckFilter('filter-classificacao', c));
    });

    // PESTEL Tags
    (state.filters.pestels || []).forEach(p => {
      addTag(`PESTEL: ${p}`, () => uncheckFilter('filter-pestel', p));
    });

    // Tópicos Tags
    (state.filters.topicos || []).forEach(t => {
      const shortTopico = t.length > 22 ? t.substring(0, 20) + '...' : t;
      addTag(`Tópico: ${shortTopico}`, () => uncheckFilter('filter-topico', t));
    });

    // Botão Limpar Filtros se houver 2 ou mais critérios
    const totalActiveCount = (state.filters.search ? 1 : 0) +
      (state.filters.status ? state.filters.status.length : 0) +
      (state.filters.eixos ? state.filters.eixos.length : 0) +
      (state.filters.classificacoes ? state.filters.classificacoes.length : 0) +
      (state.filters.pestels ? state.filters.pestels.length : 0) +
      (state.filters.topicos ? state.filters.topicos.length : 0);

    if (totalActiveCount >= 2) {
      const clearBtn = document.createElement('button');
      clearBtn.type = 'button';
      clearBtn.className = 'filter-tag-clear-all';
      clearBtn.innerHTML = '<i class="fa-solid fa-rotate-left"></i> Limpar filtros';
      clearBtn.addEventListener('click', () => {
        resetSfFilters();
      });
      container.appendChild(clearBtn);
    }
  }

  // ========================================================================
  // KPIS RENDERING
  // ========================================================================
  function renderKPIs() {
    const totalSF = state.rawSementes.length;
    const artigosCitados = state.rawArticles.filter(a => a.associatedSFs.length > 0).length;
    const totalEixos = new Set(state.rawSementes.map(s => s.eixo).filter(Boolean)).size;
    const totalRevisadas = state.rawSementes.filter(s => s.revisada).length;
    const totalAprovadas = state.rawSementes.filter(s => s.statusTipo === 'ok').length;
    const totalMescladas = state.rawSementes.filter(s => s.statusTipo === 'mesclada').length;
    const totalDescartadas = state.rawSementes.filter(s => s.statusTipo === 'descartada').length;
    const totalDivididas = state.rawSementes.filter(s => s.statusTipo === 'dividida').length;

    document.getElementById('kpi-total-sf').textContent = totalSF;
    document.getElementById('kpi-artigos-citados').textContent = artigosCitados;
    document.getElementById('kpi-total-eixos').textContent = totalEixos;
    
    const kpiRev = document.getElementById('kpi-revisadas');
    if (kpiRev) {
      kpiRev.textContent = totalRevisadas;
      kpiRev.title = `${totalRevisadas} revisadas (${totalAprovadas} OK, ${totalMescladas} Mesclada, ${totalDivididas} Dividida, ${totalDescartadas} Descartada)`;
    }

    document.getElementById('badge-total-sf').textContent = totalSF;
    document.getElementById('badge-total-scopus').textContent = state.rawArticles.length;
  }

  // ========================================================================
  // EXECUTIVE ANALYTICS: CHART.JS VISUALIZATIONS
  // ========================================================================
  function renderCharts() {
    if (typeof Chart === 'undefined') return;

    const themeColors = PALETTE[state.theme];

    // Chart 1: Tipologia Prospectiva (Doughnut)
    const ctxClass = document.getElementById('chart-classificacao');
    if (ctxClass) {
      const counts = {};
      state.rawSementes.forEach(s => {
        counts[s.classificacao] = (counts[s.classificacao] || 0) + 1;
      });

      const labels = Object.keys(counts);
      const data = Object.values(counts);
      const bgColors = labels.map(l => themeColors.typology[l] || '#64748b');

      if (state.charts.classificacao) state.charts.classificacao.destroy();
      state.charts.classificacao = new Chart(ctxClass, {
        type: 'doughnut',
        data: {
          labels: labels,
          datasets: [{
            data: data,
            backgroundColor: bgColors,
            borderWidth: 1,
            borderColor: state.theme === 'dark' ? '#0e1523' : '#ffffff'
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: {
              position: 'right',
              labels: {
                boxWidth: 12,
                color: themeColors.text,
                font: { family: 'Inter', size: 11 }
              }
            }
          },
          cutout: '65%'
        }
      });
    }

    // Chart 2: Matriz PESTEL (Bar)
    const ctxPestel = document.getElementById('chart-pestel');
    if (ctxPestel) {
      const pestelCounts = {
        'Político': 0, 'Econômico': 0, 'Social': 0,
        'Tecnológico': 0, 'Ambiental': 0, 'Legal': 0
      };
      state.rawSementes.forEach(s => {
        s.pestel.forEach(p => {
          if (pestelCounts[p] !== undefined) pestelCounts[p]++;
        });
      });

      const labels = Object.keys(pestelCounts);
      const data = Object.values(pestelCounts);
      const bgColors = labels.map(l => themeColors.pestel[l] || '#64748b');

      if (state.charts.pestel) state.charts.pestel.destroy();
      state.charts.pestel = new Chart(ctxPestel, {
        type: 'bar',
        data: {
          labels: labels,
          datasets: [{
            label: 'Sementes Impactadas',
            data: data,
            backgroundColor: bgColors,
            borderRadius: 4
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { display: false }
          },
          scales: {
            x: {
              grid: { display: false },
              ticks: { color: themeColors.text, font: { family: 'Inter', size: 10 } }
            },
            y: {
              grid: { color: themeColors.grid },
              ticks: { color: themeColors.text, font: { family: 'JetBrains Mono', size: 10 } }
            }
          }
        }
      });
    }

    // Chart 3: Cobertura por Eixo Estratégico (Horizontal Bar)
    const ctxEixos = document.getElementById('chart-eixos');
    if (ctxEixos) {
      const eixoCounts = {};
      state.rawSementes.forEach(s => {
        const simplified = s.eixo.split('-')[0].trim();
        eixoCounts[simplified] = (eixoCounts[simplified] || 0) + 1;
      });

      const labels = Object.keys(eixoCounts).sort();
      const data = labels.map(l => eixoCounts[l]);

      if (state.charts.eixos) state.charts.eixos.destroy();
      state.charts.eixos = new Chart(ctxEixos, {
        type: 'bar',
        data: {
          labels: labels,
          datasets: [{
            label: 'Total de Sementes',
            data: data,
            backgroundColor: themeColors.eixos,
            borderRadius: 4
          }]
        },
        options: {
          indexAxis: 'y',
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { display: false }
          },
          scales: {
            x: {
              grid: { color: themeColors.grid },
              ticks: { color: themeColors.text, font: { family: 'JetBrains Mono', size: 10 } }
            },
            y: {
              grid: { display: false },
              ticks: { color: themeColors.text, font: { family: 'Inter', size: 10 } }
            }
          }
        }
      });
    }
  }

  function updateChartsData() {
    if (!state.charts.classificacao) return;

    const themeColors = PALETTE[state.theme];
    const sourceData = state.filteredSementes;

    // 1. Tipologia
    const counts = {};
    sourceData.forEach(s => {
      counts[s.classificacao] = (counts[s.classificacao] || 0) + 1;
    });
    state.charts.classificacao.data.labels = Object.keys(counts);
    state.charts.classificacao.data.datasets[0].data = Object.values(counts);
    state.charts.classificacao.data.datasets[0].backgroundColor = Object.keys(counts).map(l => themeColors.typology[l] || '#64748b');
    state.charts.classificacao.update();

    // 2. PESTEL
    const pestelCounts = {
      'Político': 0, 'Econômico': 0, 'Social': 0,
      'Tecnológico': 0, 'Ambiental': 0, 'Legal': 0
    };
    sourceData.forEach(s => {
      s.pestel.forEach(p => {
        if (pestelCounts[p] !== undefined) pestelCounts[p]++;
      });
    });
    state.charts.pestel.data.datasets[0].data = Object.values(pestelCounts);
    state.charts.pestel.update();

    // 3. Eixos
    const eixoCounts = {};
    sourceData.forEach(s => {
      const simplified = s.eixo.split('-')[0].trim();
      eixoCounts[simplified] = (eixoCounts[simplified] || 0) + 1;
    });
    const labels = Object.keys(eixoCounts).sort();
    state.charts.eixos.data.labels = labels;
    state.charts.eixos.data.datasets[0].data = labels.map(l => eixoCounts[l]);
    state.charts.eixos.update();
  }

  // ========================================================================
  // SCOPUS REFERENCES TABLE VIEW (TAB 2)
  // ========================================================================
  function applyScopusFilters() {
    const q = removeAccents(state.scopusFilters.search.trim());
    const rel = state.scopusFilters.relation;
    const yr = state.scopusFilters.year;

    state.filteredArticles = state.rawArticles.filter(art => {
      if (q) {
        const matches = 
          String(art.id).includes(q) ||
          removeAccents(art.title).includes(q) ||
          removeAccents(art.authors).includes(q) ||
          removeAccents(art.sourceTitle).includes(q) ||
          removeAccents(art.authorKeywords).includes(q) ||
          removeAccents(art.abstract).includes(q);
        if (!matches) return false;
      }

      if (rel === 'linked' && art.associatedSFs.length === 0) return false;
      if (rel === 'unlinked' && art.associatedSFs.length > 0) return false;
      if (yr && String(art.year) !== String(yr)) return false;

      return true;
    });

    sortFilteredArticles();
    state.paginationScopus.page = 1;
    renderScopusTable();
  }

  function sortFilteredArticles() {
    const { column, direction } = state.sortingScopus;
    const modifier = direction === 'asc' ? 1 : -1;

    state.filteredArticles.sort((a, b) => {
      if (column === 'id' || column === 'year') {
        const nA = parseInt(a[column], 10) || 0;
        const nB = parseInt(b[column], 10) || 0;
        return (nA - nB) * modifier;
      }
      const valA = String(a[column] || '').toLowerCase();
      const valB = String(b[column] || '').toLowerCase();
      return valA.localeCompare(valB, 'pt-BR') * modifier;
    });
  }

  function renderScopusTable() {
    const tbody = document.getElementById('scopus-table-tbody');
    tbody.innerHTML = '';

    const total = state.rawArticles.length;
    const filteredCount = state.filteredArticles.length;

    document.getElementById('scopus-filtered-count').textContent = filteredCount;
    document.getElementById('scopus-total-count').textContent = total;

    const { page, pageSize } = state.paginationScopus;
    const size = parseInt(pageSize, 10);
    const start = (page - 1) * size;
    const end = Math.min(start + size, filteredCount);
    const pageItems = state.filteredArticles.slice(start, end);

    if (pageItems.length === 0) {
      tbody.innerHTML = `
        <tr>
          <td colspan="6" style="text-align: center; padding: 32px; color: var(--text-muted);">
            Nenhum artigo encontrado para os critérios de busca.
          </td>
        </tr>
      `;
      renderScopusPagination();
      return;
    }

    pageItems.forEach(art => {
      const tr = document.createElement('tr');

      const linkedChips = art.associatedSFs.map(sfId => {
        return `<span class="badge-id" style="cursor: pointer; margin: 2px;" data-sf-id="${sfId}">${sfId}</span>`;
      }).join('') || '<span style="color: var(--text-muted); font-size: 0.74rem;">Sem vínculo direto</span>';

      const doiButton = art.doiLink 
        ? `<a href="${art.doiLink}" target="_blank" rel="noopener noreferrer" class="btn-art-action" style="display: inline-flex;" title="Abrir DOI na Scopus / Editora"><i class="fa-solid fa-arrow-up-right-from-square"></i> DOI</a>`
        : '';

      tr.innerHTML = `
        <td class="td-id">#${art.id}</td>
        <td>
          <div class="td-title-group">
            <span class="td-title-main">${escapeHtml(art.title)}</span>
            <span class="td-title-topic" style="font-style: italic;">${escapeHtml(art.sourceTitle)}</span>
          </div>
        </td>
        <td style="font-size: 0.8rem; color: var(--text-secondary);">${escapeHtml(art.authors)}</td>
        <td style="text-align: center; font-family: var(--font-mono); font-weight: 600;">${art.year}</td>
        <td>${linkedChips}</td>
        <td style="text-align: right;">
          <div style="display: flex; align-items: center; justify-content: flex-end; gap: 6px;">
            <button class="btn-art-action btn-open-abstract" data-art-id="${art.id}">
              <i class="fa-solid fa-file-lines"></i> Resumo
            </button>
            ${doiButton}
          </div>
        </td>
      `;

      tr.querySelector('.btn-open-abstract').addEventListener('click', () => {
        openArticleModal(art.id);
      });

      tr.querySelectorAll('.badge-id[data-sf-id]').forEach(chip => {
        chip.addEventListener('click', (e) => {
          e.stopPropagation();
          const sfId = chip.getAttribute('data-sf-id');
          // Switch to SF tab and select
          document.getElementById('tab-btn-sf').click();
          selectSf(sfId);
        });
      });

      tbody.appendChild(tr);
    });

    renderScopusPagination();
    updateScopusTableSortHeaders();
  }

  function renderScopusPagination() {
    const { page, pageSize } = state.paginationScopus;
    const total = state.filteredArticles.length;
    const size = parseInt(pageSize, 10);
    const totalPages = Math.ceil(total / size) || 1;

    const start = total === 0 ? 0 : (page - 1) * size + 1;
    const end = Math.min(page * size, total);

    document.getElementById('scopus-pagination-info').textContent = 
      `Exibindo ${start} a ${end} de ${total} artigos`;

    document.getElementById('scopus-page-current').textContent = `${page} / ${totalPages}`;
    document.getElementById('scopus-page-prev').disabled = (page <= 1);
    document.getElementById('scopus-page-next').disabled = (page >= totalPages);
  }

  // ========================================================================
  // ARTIGO MODAL (SCOPUS VIEW)
  // ========================================================================
  function openArticleModal(artId) {
    const art = state.articlesMap.get(artId);
    if (!art) return;

    const modal = document.getElementById('modal-article');

    document.getElementById('modal-art-id').textContent = `ID Scopus: ${art.id}`;
    document.getElementById('modal-art-year').textContent = art.year || 'Ano não informado';
    document.getElementById('modal-art-type').textContent = art.documentType || 'Artigo';

    document.getElementById('modal-art-title').textContent = art.title;
    document.getElementById('modal-art-authors').textContent = art.authors;
    document.getElementById('modal-art-source').textContent = art.sourceTitle || '';

    const doiSection = document.getElementById('modal-art-doi-section');
    const doiLink = document.getElementById('modal-art-doi-link');
    if (art.doiLink) {
      doiSection.style.display = 'block';
      doiLink.href = art.doiLink;
      doiLink.textContent = art.doiLink;
    } else {
      doiSection.style.display = 'none';
    }

    const abstractEl = document.getElementById('modal-art-abstract');
    abstractEl.textContent = art.abstract || 'Resumo não disponível na extração bibliográfica.';

    const kwSection = document.getElementById('modal-art-keywords-section');
    const kwCloud = document.getElementById('modal-art-keywords');
    const keywords = (art.authorKeywords || art.indexKeywords || '')
      .split(';')
      .map(k => k.trim())
      .filter(Boolean);

    if (keywords.length > 0) {
      kwSection.style.display = 'block';
      kwCloud.innerHTML = keywords.map(k => `<span class="keyword-chip">${escapeHtml(k)}</span>`).join('');
    } else {
      kwSection.style.display = 'none';
    }

    const linkedChips = document.getElementById('modal-art-linked-sfs');
    linkedChips.innerHTML = '';
    if (art.associatedSFs.length === 0) {
      linkedChips.innerHTML = '<span style="color: var(--text-muted); font-size: 0.82rem;">Este artigo não está vinculado a nenhuma semente de futuro no momento.</span>';
    } else {
      art.associatedSFs.forEach(sfId => {
        const chip = document.createElement('span');
        chip.className = 'sf-link-chip';
        chip.innerHTML = `<i class="fa-solid fa-seedling"></i> ${sfId}`;
        chip.addEventListener('click', () => {
          closeArticleModal();
          document.getElementById('tab-btn-sf').click();
          selectSf(sfId);
        });
        linkedChips.appendChild(chip);
      });
    }

    modal.style.display = 'flex';
  }

  function closeArticleModal() {
    const modal = document.getElementById('modal-article');
    modal.style.display = 'none';
  }

  // ========================================================================
  // EVENT LISTENERS SETUP
  // ========================================================================
  function setupEventListeners() {
    // 1. Tab Navigation
    const tabSf = document.getElementById('tab-btn-sf');
    const tabScopus = document.getElementById('tab-btn-scopus');
    const viewSf = document.getElementById('view-sf');
    const viewScopus = document.getElementById('view-scopus');

    tabSf.addEventListener('click', () => {
      tabSf.classList.add('active');
      tabScopus.classList.remove('active');
      viewSf.style.display = 'flex';
      viewScopus.style.display = 'none';
      state.activeTab = 'sf';
    });

    tabScopus.addEventListener('click', () => {
      tabScopus.classList.add('active');
      tabSf.classList.remove('active');
      viewScopus.style.display = 'flex';
      viewSf.style.display = 'none';
      state.activeTab = 'scopus';
      renderScopusTable();
    });

    // 2. Analytics Collapsible Toggle
    const btnAnalytics = document.getElementById('btn-toggle-analytics');
    const analyticsPanel = document.getElementById('analytics-panel');

    btnAnalytics.addEventListener('click', () => {
      state.analyticsExpanded = !state.analyticsExpanded;
      if (state.analyticsExpanded) {
        analyticsPanel.classList.remove('collapsed');
        btnAnalytics.classList.add('active');
        setTimeout(() => {
          if (state.charts.classificacao) state.charts.classificacao.resize();
          if (state.charts.pestel) state.charts.pestel.resize();
          if (state.charts.eixos) state.charts.eixos.resize();
        }, 150);
      } else {
        analyticsPanel.classList.add('collapsed');
        btnAnalytics.classList.remove('active');
      }
    });

    // 3. SF Filters Handlers
    const sfSearchInput = document.getElementById('sf-search-input');
    const sfSearchClear = document.getElementById('sf-search-clear');

    if (sfSearchInput) {
      sfSearchInput.addEventListener('input', (e) => {
        state.filters.search = e.target.value;
        if (sfSearchClear) sfSearchClear.style.display = state.filters.search ? 'block' : 'none';
        debouncedApplySfFilters();
      });
    }

    if (sfSearchClear) {
      sfSearchClear.addEventListener('click', () => {
        if (sfSearchInput) sfSearchInput.value = '';
        state.filters.search = '';
        sfSearchClear.style.display = 'none';
        applySfFilters();
      });
    }

    // Checkbox delegation for all SF filters in sidebar
    const sidebarBody = document.querySelector('.sidebar-body');
    if (sidebarBody) {
      sidebarBody.addEventListener('change', (e) => {
        if (e.target && e.target.classList.contains('filter-checkbox')) {
          syncFilterStateFromDOM();
          applySfFilters();
        }
      });
    }

    const btnClearFilters = document.getElementById('btn-clear-filters');
    if (btnClearFilters) {
      btnClearFilters.addEventListener('click', () => {
        resetSfFilters();
      });
    }

    // 4. SF List Sorting Selector
    document.getElementById('sf-list-sort').addEventListener('change', (e) => {
      state.listSort = e.target.value;
      sortFilteredSementes();
      renderSfInboxList();
    });

    // 5. Scopus Filters Handlers
    const scopusSearchInput = document.getElementById('scopus-search-input');
    const scopusSearchClear = document.getElementById('scopus-search-clear');

    scopusSearchInput.addEventListener('input', (e) => {
      state.scopusFilters.search = e.target.value;
      scopusSearchClear.style.display = state.scopusFilters.search ? 'block' : 'none';
      debouncedApplyScopusFilters();
    });

    scopusSearchClear.addEventListener('click', () => {
      scopusSearchInput.value = '';
      state.scopusFilters.search = '';
      scopusSearchClear.style.display = 'none';
      applyScopusFilters();
    });

    document.getElementById('scopus-filter-relation').addEventListener('change', (e) => {
      state.scopusFilters.relation = e.target.value;
      applyScopusFilters();
    });

    document.getElementById('scopus-filter-year').addEventListener('change', (e) => {
      state.scopusFilters.year = e.target.value;
      applyScopusFilters();
    });

    document.getElementById('scopus-page-size').addEventListener('change', (e) => {
      state.paginationScopus.pageSize = e.target.value;
      state.paginationScopus.page = 1;
      renderScopusTable();
    });

    document.getElementById('scopus-page-prev').addEventListener('click', () => {
      if (state.paginationScopus.page > 1) {
        state.paginationScopus.page--;
        renderScopusTable();
      }
    });

    document.getElementById('scopus-page-next').addEventListener('click', () => {
      const size = parseInt(state.paginationScopus.pageSize, 10);
      const totalPages = Math.ceil(state.filteredArticles.length / size);
      if (state.paginationScopus.page < totalPages) {
        state.paginationScopus.page++;
        renderScopusTable();
      }
    });

    document.querySelectorAll('#scopus-table th.th-sortable').forEach(th => {
      th.addEventListener('click', () => {
        const sortCol = th.getAttribute('data-sort');
        if (state.sortingScopus.column === sortCol) {
          state.sortingScopus.direction = state.sortingScopus.direction === 'asc' ? 'desc' : 'asc';
        } else {
          state.sortingScopus.column = sortCol;
          state.sortingScopus.direction = 'asc';
        }
        sortFilteredArticles();
        renderScopusTable();
      });
    });

    // 6. Article Modal Close Handlers
    document.getElementById('btn-close-art-modal').addEventListener('click', closeArticleModal);
    document.getElementById('btn-art-modal-close-footer').addEventListener('click', closeArticleModal);
    document.getElementById('modal-article').addEventListener('click', (e) => {
      if (e.target.id === 'modal-article') closeArticleModal();
    });

    // 7. Keyboard Navigation (Arrows Up / Down for SF Inbox reading)
    window.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        closeArticleModal();
        return;
      }

      // Ignore arrows if user is focused inside an input or select
      const activeTag = document.activeElement ? document.activeElement.tagName.toLowerCase() : '';
      if (activeTag === 'input' || activeTag === 'select' || activeTag === 'textarea') {
        return;
      }

      if (state.activeTab === 'sf' && state.filteredSementes.length > 0) {
        if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
          e.preventDefault();
          const currentIndex = state.filteredSementes.findIndex(s => s.id === state.selectedSfId);
          let newIndex = 0;

          if (e.key === 'ArrowDown') {
            if (currentIndex === -1 || currentIndex >= state.filteredSementes.length - 1) {
              newIndex = 0;
            } else {
              newIndex = currentIndex + 1;
            }
          } else if (e.key === 'ArrowUp') {
            if (currentIndex <= 0) {
              newIndex = state.filteredSementes.length - 1;
            } else {
              newIndex = currentIndex - 1;
            }
          }

          selectSf(state.filteredSementes[newIndex].id);
        }
      }
    });
  }

  function resetSfFilters() {
    state.filters = {
      search: '',
      status: [],
      eixos: [],
      classificacoes: [],
      pestels: [],
      topicos: []
    };
    const sfSearchInput = document.getElementById('sf-search-input');
    if (sfSearchInput) sfSearchInput.value = '';
    const sfSearchClear = document.getElementById('sf-search-clear');
    if (sfSearchClear) sfSearchClear.style.display = 'none';

    document.querySelectorAll('.filter-checkbox').forEach(cb => {
      cb.checked = false;
    });
    document.querySelectorAll('.filter-checkbox-label').forEach(lbl => {
      lbl.classList.remove('is-checked');
    });

    ['status', 'eixo', 'classificacao', 'pestel', 'topico'].forEach(group => {
      updateActiveCountBadge(group, 0);
    });

    applySfFilters();
  }

  // ========================================================================
  // THEME MANAGEMENT (DARK / LIGHT SOBER EXECUTIVE)
  // ========================================================================
  function setupTheme() {
    const savedTheme = safeGetStorage('sfcsed_theme', 'dark');
    setTheme(savedTheme);

    const btnTheme = document.getElementById('btn-theme-toggle');
    btnTheme.addEventListener('click', () => {
      const newTheme = state.theme === 'dark' ? 'light' : 'dark';
      setTheme(newTheme);
    });
  }

  function setTheme(theme) {
    state.theme = theme;
    document.documentElement.setAttribute('data-theme', theme);
    safeSetStorage('sfcsed_theme', theme);

    const icon = document.getElementById('theme-icon');
    const text = document.getElementById('theme-text');
    if (theme === 'dark') {
      if (icon) icon.className = 'fa-solid fa-sun';
      if (text) text.textContent = 'Modo Claro';
    } else {
      if (icon) icon.className = 'fa-solid fa-moon';
      if (text) text.textContent = 'Modo Escuro';
    }

    renderCharts();
  }

  // ========================================================================
  // ACCESSIBILITY: FONT SIZE SCALING (A- / A / A+)
  // ========================================================================
  const FONT_SCALES = [
    { label: '90%', value: 90 },
    { label: '100%', value: 100 },
    { label: '110%', value: 110 },
    { label: '120%', value: 120 },
    { label: '130%', value: 130 }
  ];
  const DEFAULT_FONT_SCALE_INDEX = 1; // 100%

  function setupFontScale() {
    const savedScale = safeGetStorage('sfcsed_font_scale');
    let scaleIndex = savedScale !== null ? parseInt(savedScale, 10) : DEFAULT_FONT_SCALE_INDEX;
    if (isNaN(scaleIndex) || scaleIndex < 0 || scaleIndex >= FONT_SCALES.length) {
      scaleIndex = DEFAULT_FONT_SCALE_INDEX;
    }

    applyFontScale(scaleIndex);

    const btnDec = document.getElementById('btn-font-dec');
    const btnInc = document.getElementById('btn-font-inc');
    const btnReset = document.getElementById('btn-font-reset');

    if (btnDec) {
      btnDec.addEventListener('click', () => {
        const current = getStoredFontScaleIndex();
        if (current > 0) {
          applyFontScale(current - 1);
        }
      });
    }

    if (btnInc) {
      btnInc.addEventListener('click', () => {
        const current = getStoredFontScaleIndex();
        if (current < FONT_SCALES.length - 1) {
          applyFontScale(current + 1);
        }
      });
    }

    if (btnReset) {
      btnReset.addEventListener('click', () => {
        applyFontScale(DEFAULT_FONT_SCALE_INDEX);
      });
    }
  }

  function getStoredFontScaleIndex() {
    const val = safeGetStorage('sfcsed_font_scale');
    const idx = val !== null ? parseInt(val, 10) : DEFAULT_FONT_SCALE_INDEX;
    return isNaN(idx) ? DEFAULT_FONT_SCALE_INDEX : Math.max(0, Math.min(FONT_SCALES.length - 1, idx));
  }

  function applyFontScale(index) {
    const scale = FONT_SCALES[index];
    if (scale.value === 100) {
      document.documentElement.style.fontSize = '';
    } else {
      document.documentElement.style.fontSize = `${scale.value}%`;
    }
    safeSetStorage('sfcsed_font_scale', index);

    const btnDec = document.getElementById('btn-font-dec');
    const btnInc = document.getElementById('btn-font-inc');
    const btnReset = document.getElementById('btn-font-reset');
    const label = document.getElementById('font-scale-label');

    if (btnDec) btnDec.disabled = (index === 0);
    if (btnInc) btnInc.disabled = (index === FONT_SCALES.length - 1);

    if (label && btnReset) {
      if (scale.value === 100) {
        label.textContent = 'A';
        btnReset.classList.remove('is-scaled');
        btnReset.title = 'Tamanho padrão (100%). Clique para redefinir.';
      } else {
        label.textContent = `${scale.value}%`;
        btnReset.classList.add('is-scaled');
        btnReset.title = `Tamanho atual: ${scale.value}%. Clique para voltar ao padrão (100%).`;
      }
    }
  }

  // ========================================================================
  // UTILITY HELPERS
  // ========================================================================
  function removeAccents(str) {
    if (!str) return '';
    return String(str)
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .toLowerCase();
  }

  function cleanBadgeClass(str) {
    if (!str) return 'generic';
    return String(str)
      .toLowerCase()
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .replace(/\s+/g, '-');
  }

  function escapeHtml(str) {
    if (!str) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  function highlightNumericCitations(text) {
    if (!text) return '';
    const safe = escapeHtml(text);
    return safe.replace(/\((\d+(?:[\s,]+\d+)*)\)/g, (match, numsStr) => {
      const parts = numsStr.split(/[\s,]+/).filter(Boolean);
      const links = parts.map(num => {
        return `<a href="#art-card-${num}" class="citation-badge-link" data-art-ref="${num}" title="Ver referência #${num}">#${num}</a>`;
      }).join(', ');
      return `<span class="citation-badge-group">(${links})</span>`;
    });
  }

  function copyTextToClipboard(text, onSuccess, onError) {
    if (navigator.clipboard && window.isSecureContext) {
      navigator.clipboard.writeText(text).then(() => {
        if (onSuccess) onSuccess();
      }).catch(() => {
        fallbackCopyText(text, onSuccess, onError);
      });
    } else {
      fallbackCopyText(text, onSuccess, onError);
    }
  }

  function fallbackCopyText(text, onSuccess, onError) {
    try {
      const textArea = document.createElement('textarea');
      textArea.value = text;
      textArea.style.position = 'fixed';
      textArea.style.left = '-999999px';
      textArea.style.top = '-999999px';
      textArea.setAttribute('readonly', '');
      document.body.appendChild(textArea);
      textArea.focus();
      textArea.select();
      const successful = document.execCommand('copy');
      document.body.removeChild(textArea);
      if (successful) {
        if (onSuccess) onSuccess();
      } else if (onError) {
        onError();
      }
    } catch (err) {
      console.warn('Falha na cópia para o clipboard:', err);
      if (onError) onError(err);
    }
  }

  function safeGetStorage(key, fallback = null) {
    try {
      const val = localStorage.getItem(key);
      return val !== null ? val : fallback;
    } catch (e) {
      return fallback;
    }
  }

  function safeSetStorage(key, value) {
    try {
      localStorage.setItem(key, value);
    } catch (e) {
      // Storage indisponível ou cota excedida
    }
  }

  function updateScopusTableSortHeaders() {
    const { column, direction } = state.sortingScopus;
    document.querySelectorAll('#scopus-table th.th-sortable').forEach(th => {
      const col = th.getAttribute('data-sort');
      const icon = th.querySelector('i');
      if (col === column) {
        th.classList.add('sorted-active');
        if (icon) {
          icon.className = direction === 'asc' ? 'fa-solid fa-sort-up' : 'fa-solid fa-sort-down';
        }
      } else {
        th.classList.remove('sorted-active');
        if (icon) {
          icon.className = 'fa-solid fa-sort';
        }
      }
    });
  }

  function hideLoadingOverlay() {
    const overlay = document.getElementById('loading-overlay');
    if (overlay) {
      overlay.style.opacity = '0';
      overlay.style.pointerEvents = 'none';
      setTimeout(() => {
        overlay.style.visibility = 'hidden';
      }, 350);
    }
  }

  function updateLoadingStatus(msg) {
    const el = document.getElementById('loading-status');
    if (el) el.textContent = msg;
  }

  function debounce(func, delay) {
    let timer;
    return function(...args) {
      clearTimeout(timer);
      timer = setTimeout(() => func.apply(this, args), delay);
    };
  }

  const debouncedApplySfFilters = debounce(() => applySfFilters(), 250);
  const debouncedApplyScopusFilters = debounce(() => applyScopusFilters(), 250);

  // Execute initialization
  init();

});
