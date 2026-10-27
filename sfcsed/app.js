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
      topicos: []
    },

    // Sort order for Sementes List
    listSort: 'id-asc',

    // Filters for Referências
    scopusFilters: {
      search: '',
      year: ''
    },

    // View & Layout
    activeTab: 'sf', // 'sf', 'map' or 'scopus'
    analyticsExpanded: false,
    theme: 'dark',

    // Matrix Map State
    mapFilters: {
      search: '',
      classificacao: ''
    },
    selectedMapSfId: null,

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
        'Tendência': '#3b82f6',
        'Fato Portador de Futuro': '#10b981',
        'Incerteza': '#f59e0b',
        'Ruptura': '#ef4444',
        'Surpresa inevitável': '#f97316',
        'Wild Card': '#a855f7',
        'Sinal fraco': '#a855f7',
        'Sinal Fraco': '#a855f7',
        'Fato predeterminado': '#94a3b8',
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
        'Tendência': '#2563eb',
        'Fato Portador de Futuro': '#047857',
        'Incerteza': '#d97706',
        'Ruptura': '#dc2626',
        'Surpresa inevitável': '#ea580c',
        'Wild Card': '#7e22ce',
        'Sinal fraco': '#9333ea',
        'Sinal Fraco': '#9333ea',
        'Fato predeterminado': '#475569',
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
    try {
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
      renderSfMatrixMap();
    } catch (err) {
      console.error('Erro na inicialização da plataforma:', err);
    } finally {
      hideLoadingOverlay();
    }
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

      // Parse PESTEL (normalizado)
      const pestelNormMap = {
        'politico': 'Político', 'político': 'Político',
        'economico': 'Econômico', 'econômico': 'Econômico',
        'social': 'Social',
        'tecnologico': 'Tecnológico', 'tecnológico': 'Tecnológico',
        'ambiental': 'Ambiental',
        'legal': 'Legal', 'lega': 'Legal'
      };
      const rawPestel = String(getRowVal(row, ['PESTEL', 'Dimensões PESTEL', 'Dimensoes PESTEL', 'Dimensões']) || '');
      const pestelList = rawPestel.replace(/;/g, ',').split(',').map(d => {
        const clean = d.trim();
        return pestelNormMap[clean.toLowerCase()] || clean;
      }).filter(Boolean);

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
        topico: String(getRowVal(row, ['Tópico', 'Tópico Mais Amplo Padronizado', 'Topico'], 'Geral')).trim(),
        titulo: String(getRowVal(row, ['Título SF', 'Título da Semente de Futuro', 'Titulo SF', 'Título', 'Titulo'], 'Sem título')).trim(),
        descricao: getRowVal(row, ['Descrição (Ref N)', 'Descrição Global do Fenômeno (com Citações Numéricas)', 'Descricao (Ref N)', 'Descrição', 'Descricao'], ''),
        descricaoApa: getRowVal(row, ['Descrição (Ref APA)', 'Descricao (Ref APA)'], ''),
        classificacao: String(getRowVal(row, ['Classificação', 'Classificação Prospectiva', 'Classificacao'], 'Não classificada')).trim(),
        pestel: pestelList,
        eixo: String(getRowVal(row, ['Eixo', 'Eixo CSED', 'Eixo Temático'], 'Não categorizado')).trim(),
        qtdArtigos: artIds.length || (isNaN(parsedQtd) ? 0 : parsedQtd),
        artigosIds: artIds,
        referenciasResumidas: getRowVal(row, ['Referências', 'Referências Bibliográficas Resumidas (Autores, Ano, Título)', 'Referencias'], ''),
        implicacoes: getRowVal(row, ['Implicações', 'Conexão / Implicações Estratégicas para o Ensino Superior Privado até 2040', 'Implicacoes'], ''),
        indicadores: indVal,
        revisada: isRev,
        statusTipo: statusTipo,
        statusRevisao: (rawRevStr && rawRevStr !== '0') ? rawRevStr : 'Consolidada',
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

    // Filtrar apenas artigos/referências citadas diretamente em SFs
    state.rawArticles = state.rawArticles.filter(a => a.associatedSFs && a.associatedSFs.length > 0);
    state.rawArticles.sort((a, b) => (parseInt(a.id, 10) || 0) - (parseInt(b.id, 10) || 0));

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
    // 1. Eixos Estratégicos
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

    // 3. Classificação
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
        <option value="linked" selected>Apenas Referências Citadas (${linkedCount})</option>
        <option value="unlinked">Não Citadas Diretamente (${unlinkedCount})</option>
        <option value="">Todas as Referências (${state.rawArticles.length})</option>
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

    state.filters.eixos = getChecked('filter-eixo');
    state.filters.classificacoes = getChecked('filter-classificacao');
    state.filters.pestels = getChecked('filter-pestel');
    state.filters.topicos = getChecked('filter-topico');

    // Update active count badges on headers
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

      // 2. Eixo Filter (Multi-select OR: sf matches if its eixo is in selected eixos)
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
      const typClass = cleanBadgeClass(sf.classificacao);
      const item = document.createElement('div');
      item.className = `inbox-item item-typology-${typClass}`;
      item.setAttribute('data-sf-id', sf.id);
      if (state.selectedSfId === sf.id) {
        item.classList.add('active');
      }

      const lineageOrigin = sf.origemDivisao || (sf.vinculo && sf.vinculo.length > 0 ? sf.vinculo[0] : '');
      const lineageBadge = lineageOrigin
        ? `<span class="badge-status-subtle badge-status-filha" title="Linhagem metodológica: derivada de ${escapeHtml(lineageOrigin)}"><i class="fa-solid fa-code-branch"></i> da ${escapeHtml(lineageOrigin)}</span>`
        : '';

      const pestelChips = sf.pestel.map(p => {
        return `<span class="pestel-chip pestel-chip-neutral">${escapeHtml(p)}</span>`;
      }).join('');

      item.innerHTML = `
        <div class="inbox-item-header">
          <div class="inbox-item-title">${escapeHtml(sf.titulo)}</div>
          <span class="inbox-item-active-arrow"><i class="fa-solid fa-chevron-right"></i></span>
        </div>

        <div class="inbox-item-topic" title="${escapeHtml(sf.topico)}">
          <i class="fa-solid fa-tag"></i> <span>${escapeHtml(sf.topico)}</span>
        </div>

        <div class="inbox-item-meta-row">
          <span class="badge-id-neutral">${sf.id}</span>
          <span class="badge-typology ${typClass}">${escapeHtml(sf.classificacao)}</span>
          ${lineageBadge}
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
    if (readingPane) readingPane.scrollTop = 0;

    // Suporte mobile: ativa visualização do dossiê em tela cheia no smartphone
    const layout = document.querySelector('.sf-email-layout');
    if (layout) {
      layout.classList.add('mobile-reading-active');
    }
    if (window.innerWidth <= 768) {
      window.scrollTo({ top: 0, behavior: 'smooth' });
    }

    const typClass = cleanBadgeClass(sf.classificacao);

    const lineageOrigin = sf.origemDivisao || (sf.vinculo && sf.vinculo.length > 0 ? sf.vinculo[0] : '');
    const lineageChip = lineageOrigin
      ? `<span class="badge-eixo-lg" title="Linhagem metodológica: derivada de ${escapeHtml(lineageOrigin)}"><i class="fa-solid fa-code-branch"></i> da ${escapeHtml(lineageOrigin)}</span>`
      : '';

    const pestelChips = sf.pestel.map(p => {
      const pClass = cleanBadgeClass(p);
      return `<span class="pestel-chip ${pClass}" style="font-size: 0.74rem; padding: 2px 8px;">${p}</span>`;
    }).join('');

    const descFormatted = highlightNumericCitations(sf.descricao);

    // Banner de linhagem se houver
    let lineageBannerHtml = '';
    if (lineageOrigin) {
      const parentSf = state.rawSementes.find(s => s.id === lineageOrigin);
      lineageBannerHtml = `
        <div class="dossier-lineage-banner">
          <i class="fa-solid fa-code-branch"></i>
          <div>
            <strong>Nota Metodológica de Linhagem:</strong> Semente derivada/desmembrada de <strong>${escapeHtml(lineageOrigin)}</strong>${parentSf ? ' (' + escapeHtml(parentSf.titulo) + ')' : ''} para maior especificidade analítica no estudo prospectivo.
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
      <!-- Mobile Top Navigation Bar (Visível apenas em smartphones) -->
      <div class="mobile-reading-nav mobile-only">
        <button type="button" class="btn-mobile-back-to-list" id="btn-mobile-back-to-list" title="Voltar à lista de sementes">
          <i class="fa-solid fa-arrow-left"></i> <span>Voltar à Lista</span>
        </button>
        <span class="badge-id" style="font-size: 0.85rem; padding: 4px 10px;">${sf.id}</span>
      </div>

      <!-- Barra Superior do Dossiê -->
      <div class="dossier-top-bar">
        <div class="dossier-meta-group">
          <span class="badge-id" style="font-size: 0.85rem; padding: 3px 8px;">${sf.id}</span>
          <span class="badge-typology ${typClass}">${sf.classificacao}</span>
          <span class="badge-eixo-lg" title="${escapeHtml(sf.eixo)}">${escapeHtml(sf.eixo)}</span>
          ${lineageChip}
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

      ${lineageBannerHtml}

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
    const btnMobileBack = document.getElementById('btn-mobile-back-to-list');
    if (btnMobileBack) {
      btnMobileBack.addEventListener('click', () => {
        deselectSf();
      });
    }

    document.getElementById('btn-close-reading').addEventListener('click', () => {
      deselectSf();
    });

    document.getElementById('btn-copy-dossier').addEventListener('click', () => {
      let copyText = `[${sf.id}] ${sf.titulo}\n\nClassificação: ${sf.classificacao} | Eixo: ${sf.eixo} | Tópico: ${sf.topico}`;
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

    // Suporte mobile: desativa visualização do dossiê e retorna à lista de sementes
    const layout = document.querySelector('.sf-email-layout');
    if (layout) {
      layout.classList.remove('mobile-reading-active');
    }
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
      (state.filters.eixos ? state.filters.eixos.length : 0) +
      (state.filters.classificacoes ? state.filters.classificacoes.length : 0) +
      (state.filters.pestels ? state.filters.pestels.length : 0) +
      (state.filters.topicos ? state.filters.topicos.length : 0);

    const mobileBadge = document.getElementById('mobile-filter-badge');
    if (mobileBadge) {
      if (totalActiveCount > 0) {
        mobileBadge.textContent = totalActiveCount;
        mobileBadge.style.display = 'inline-flex';
      } else {
        mobileBadge.style.display = 'none';
      }
    }

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
    const artigosCitados = state.rawArticles.length;
    const totalEixos = new Set(state.rawSementes.map(s => s.eixo).filter(Boolean)).size;
    const totalTopicos = new Set(state.rawSementes.map(s => s.topico).filter(Boolean)).size;

    const elTotalSf = document.getElementById('kpi-total-sf');
    if (elTotalSf) elTotalSf.textContent = totalSF;

    const kpiArtCitados = document.getElementById('kpi-artigos-citados');
    if (kpiArtCitados) {
      kpiArtCitados.textContent = artigosCitados;
      kpiArtCitados.title = `${artigosCitados} referências científicas citadas diretamente nas Sementes de Futuro`;
    }

    const elTotalEixos = document.getElementById('kpi-total-eixos');
    if (elTotalEixos) elTotalEixos.textContent = totalEixos;
    
    const kpiTopicos = document.getElementById('kpi-total-topicos');
    if (kpiTopicos) {
      kpiTopicos.textContent = totalTopicos;
      kpiTopicos.title = `${totalTopicos} tópicos padronizados estruturando as ${totalSF} sementes consolidadas`;
    }

    const badgeSf = document.getElementById('badge-total-sf');
    if (badgeSf) badgeSf.textContent = totalSF;
    const badgeMap = document.getElementById('badge-total-map');
    if (badgeMap) badgeMap.textContent = totalSF;
    const badgeScopus = document.getElementById('badge-total-scopus');
    if (badgeScopus) badgeScopus.textContent = artigosCitados;
  }

  // ========================================================================
  // EXECUTIVE ANALYTICS: CHART.JS VISUALIZATIONS
  // ========================================================================
  function renderCharts() {
    if (typeof Chart === 'undefined') return;

    const themeColors = PALETTE[state.theme];

    // Chart 1: Classificação (Doughnut)
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

    // 1. Classificação
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
        ? `<a href="${art.doiLink}" target="_blank" rel="noopener noreferrer" class="btn-art-action" style="display: inline-flex;" title="Abrir DOI na Editora / Periódico"><i class="fa-solid fa-arrow-up-right-from-square"></i> DOI</a>`
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

      const btnAbstract = tr.querySelector('.btn-open-abstract');
      if (btnAbstract) {
        btnAbstract.addEventListener('click', () => {
          openArticleModal(art.id);
        });
      }

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
      `Exibindo ${start} a ${end} de ${total} referências`;

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

    document.getElementById('modal-art-id').textContent = `ID: #${art.id}`;
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
  // MATRIX MAP: TÓPICOS PADRONIZADOS × DIMENSÕES PESTEL (TAB 2)
  // ========================================================================
  const PESTEL_ORDER = ['Político', 'Econômico', 'Social', 'Tecnológico', 'Ambiental', 'Legal'];
  const PESTEL_CLASSES = {
    'Político': 'pestel-col-politico',
    'Econômico': 'pestel-col-economico',
    'Social': 'pestel-col-social',
    'Tecnológico': 'pestel-col-tecnologico',
    'Ambiental': 'pestel-col-ambiental',
    'Legal': 'pestel-col-legal'
  };

  function renderSfMatrixMap() {
    const thead = document.getElementById('sf-matrix-thead');
    const tbody = document.getElementById('sf-matrix-tbody');
    const tfoot = document.getElementById('sf-matrix-tfoot');
    if (!thead || !tbody || !tfoot) return;

    thead.innerHTML = '';
    tbody.innerHTML = '';
    tfoot.innerHTML = '';

    // Distinct topics sorted alphabetically
    const topics = [...new Set(state.rawSementes.map(s => s.topico).filter(Boolean))].sort();

    // 1. Build Thead
    const headerTr = document.createElement('tr');
    headerTr.innerHTML = `
      <th class="matrix-corner-cell">
        <div style="display: flex; align-items: center; justify-content: space-between;">
          <span>Tópico Padronizado</span>
          <span class="tab-badge" style="font-size: 0.65rem;">${topics.length} Tópicos</span>
        </div>
      </th>
      ${PESTEL_ORDER.map(p => {
        const count = state.rawSementes.filter(s => s.pestel && s.pestel.includes(p)).length;
        const colClass = PESTEL_CLASSES[p] || '';
        return `
          <th class="pestel-header-cell ${colClass}">
            <div class="pestel-header-inner">
              <span class="pestel-header-title">${p}</span>
              <span class="pestel-header-count">(${count})</span>
            </div>
          </th>
        `;
      }).join('')}
      <th class="matrix-total-header">Total SF</th>
    `;
    thead.appendChild(headerTr);

    // Matrix stats tracking
    let occupiedQuadrants = 0;
    let totalPlacements = 0;
    const colTotals = {
      'Político': 0,
      'Econômico': 0,
      'Social': 0,
      'Tecnológico': 0,
      'Ambiental': 0,
      'Legal': 0
    };

    // 2. Build Tbody
    topics.forEach((topic) => {
      const sfsInTopic = state.rawSementes.filter(s => s.topico === topic);
      const row = document.createElement('tr');

      // Sticky Topic Header Cell
      const thTopic = document.createElement('th');
      thTopic.className = 'matrix-topic-cell';
      thTopic.innerHTML = `
        <div class="matrix-topic-content">
          <i class="fa-solid fa-folder-open matrix-topic-icon"></i>
          <div class="matrix-topic-info">
            <span class="matrix-topic-name">${escapeHtml(topic)}</span>
            <span class="matrix-topic-badge">${sfsInTopic.length} ${sfsInTopic.length === 1 ? 'semente' : 'sementes'}</span>
          </div>
        </div>
      `;
      row.appendChild(thTopic);

      // PESTEL Columns
      PESTEL_ORDER.forEach(p => {
        const matchingSfs = sfsInTopic.filter(s => s.pestel && s.pestel.includes(p));
        const td = document.createElement('td');
        td.className = 'matrix-cell';
        td.setAttribute('data-topic', topic);
        td.setAttribute('data-pestel', p);

        if (matchingSfs.length === 0) {
          td.innerHTML = `<div class="matrix-empty-slot" title="Nenhuma semente no cruzamento ${escapeHtml(topic)} × ${p}">—</div>`;
        } else {
          occupiedQuadrants++;
          totalPlacements += matchingSfs.length;
          colTotals[p] += matchingSfs.length;

          const chipsWrapper = document.createElement('div');
          chipsWrapper.className = 'matrix-chips-wrapper';

          matchingSfs.forEach(sf => {
            const palette = PALETTE[state.theme] || PALETTE.dark;
            const dotColor = (palette.typology && (palette.typology[sf.classificacao] || palette.typology['Tendência'])) || '#64748b';
            const chip = document.createElement('button');
            chip.type = 'button';
            chip.className = 'matrix-sf-chip';
            chip.setAttribute('data-sf-id', sf.id);
            chip.setAttribute('data-typology', sf.classificacao);
            chip.title = `[${sf.id}] ${sf.titulo}\nClassificação: ${sf.classificacao} | Tópico: ${sf.topico}\nClique para ver detalhes`;

            chip.innerHTML = `
              <span class="chip-dot" style="background-color: ${dotColor};"></span>
              <span>${escapeHtml(sf.id)}</span>
            `;

            chip.addEventListener('click', (e) => {
              e.stopPropagation();
              openSfPreviewModal(sf.id);
            });

            chipsWrapper.appendChild(chip);
          });

          td.appendChild(chipsWrapper);
        }

        row.appendChild(td);
      });

      // Row Total Cell
      const tdRowTotal = document.createElement('td');
      tdRowTotal.className = 'matrix-row-total';
      tdRowTotal.textContent = sfsInTopic.length;
      tdRowTotal.title = `${sfsInTopic.length} sementes no tópico: ${topic}`;
      row.appendChild(tdRowTotal);

      tbody.appendChild(row);
    });

    // 3. Build Tfoot (Summary totals row)
    const tfootTr = document.createElement('tr');
    tfootTr.innerHTML = `
      <th class="matrix-footer-corner">
        <div style="display: flex; align-items: center; justify-content: space-between;">
          <span>Totais por Dimensão</span>
          <span style="font-size: 0.68rem; color: var(--text-muted); font-weight: 500;">(Incidências)</span>
        </div>
      </th>
      ${PESTEL_ORDER.map(p => {
        const distinctSfCount = state.rawSementes.filter(s => s.pestel && s.pestel.includes(p)).length;
        return `
          <td class="matrix-footer-total">
            <span title="${distinctSfCount} sementes distintas tocam a dimensão ${p}">${distinctSfCount}</span>
          </td>
        `;
      }).join('')}
      <td class="matrix-footer-grand-total" title="Total de Sementes de Futuro">
        ${state.rawSementes.length}
      </td>
    `;
    tfoot.appendChild(tfootTr);

    // Update Toolbar Metrics
    const totalQuadrants = topics.length * PESTEL_ORDER.length;
    const elTopics = document.getElementById('map-total-topics');
    const elIntersections = document.getElementById('map-total-intersections');
    const elPlacements = document.getElementById('map-total-placements');
    const badgeMap = document.getElementById('badge-total-map');

    if (elTopics) elTopics.textContent = topics.length;
    if (elIntersections) elIntersections.textContent = `${occupiedQuadrants}/${totalQuadrants}`;
    if (elPlacements) elPlacements.textContent = totalPlacements;
    if (badgeMap) badgeMap.textContent = state.rawSementes.length;

    // Apply any active filters
    applyMapFilters();
  }

  function applyMapFilters() {
    const q = removeAccents((state.mapFilters.search || '').trim());
    const fClass = state.mapFilters.classificacao;

    const chips = document.querySelectorAll('.matrix-sf-chip');
    if (!chips.length) return;

    chips.forEach(chip => {
      const sfId = chip.getAttribute('data-sf-id');
      const sf = state.rawSementes.find(s => s.id === sfId);
      if (!sf) return;

      let matches = true;

      // 1. Text Search
      if (q) {
        const matchesQuery = 
          removeAccents(sf.id).includes(q) ||
          removeAccents(sf.titulo).includes(q) ||
          removeAccents(sf.topico).includes(q) ||
          removeAccents(sf.classificacao).includes(q) ||
          removeAccents(sf.descricao).includes(q);
        if (!matchesQuery) matches = false;
      }

      // 2. Classificação Filter
      if (matches && fClass) {
        if (sf.classificacao !== fClass) matches = false;
      }

      if (matches) {
        chip.classList.remove('is-dimmed');
        if (q || fClass) {
          chip.classList.add('is-highlighted');
        } else {
          chip.classList.remove('is-highlighted');
        }
      } else {
        chip.classList.add('is-dimmed');
        chip.classList.remove('is-highlighted');
      }
    });

    // Update legend item active state
    document.querySelectorAll('.map-legend-item').forEach(item => {
      const typ = item.getAttribute('data-legend-typology');
      if (fClass && typ === fClass) {
        item.classList.add('active');
      } else {
        item.classList.remove('active');
      }
    });
  }

  function resetMapFilters() {
    state.mapFilters = {
      search: '',
      classificacao: ''
    };

    const mapSearch = document.getElementById('map-search-input');
    const mapSearchClear = document.getElementById('map-search-clear');
    const mapFilterClass = document.getElementById('map-filter-classificacao');

    if (mapSearch) mapSearch.value = '';
    if (mapSearchClear) mapSearchClear.style.display = 'none';
    if (mapFilterClass) mapFilterClass.value = '';

    document.querySelectorAll('.map-legend-item').forEach(item => item.classList.remove('active'));

    applyMapFilters();
  }

  function openSfPreviewModal(sfId) {
    const sf = state.rawSementes.find(s => s.id === sfId);
    if (!sf) return;

    state.selectedMapSfId = sfId;

    const modal = document.getElementById('modal-sf-preview');
    if (!modal) return;

    // Header info
    document.getElementById('modal-sf-id').textContent = sf.id;
    
    const typEl = document.getElementById('modal-sf-typology');
    if (typEl) {
      typEl.textContent = sf.classificacao;
      typEl.className = `badge-typology ${cleanBadgeClass(sf.classificacao)}`;
    }

    const eixoEl = document.getElementById('modal-sf-eixo');
    if (eixoEl) eixoEl.textContent = sf.eixo;

    document.getElementById('modal-sf-title').textContent = sf.titulo;
    document.getElementById('modal-sf-topic-text').textContent = sf.topico;

    // PESTEL chips
    const pestelContainer = document.getElementById('modal-sf-pestel-chips');
    if (pestelContainer) {
      pestelContainer.innerHTML = (sf.pestel || []).map(p => {
        const cls = cleanBadgeClass(p);
        return `<span class="pestel-chip ${cls}"><span class="chip-dot-pestel"></span> ${escapeHtml(p)}</span>`;
      }).join('');
    }

    // Descrição
    const descEl = document.getElementById('modal-sf-desc');
    if (descEl) descEl.innerHTML = highlightNumericCitations(sf.descricao || 'Descrição em consolidação.');

    // Implicações
    const impEl = document.getElementById('modal-sf-implications');
    if (impEl) impEl.textContent = sf.implicacoes || 'Sem implicações registradas.';

    // Indicadores
    const indSection = document.getElementById('modal-sf-indicators-section');
    const indEl = document.getElementById('modal-sf-indicators');
    if (sf.indicadores && sf.indicadores.toUpperCase() !== 'N/A') {
      indSection.style.display = 'block';
      indEl.textContent = sf.indicadores;
    } else {
      indSection.style.display = 'none';
    }

    // Referências vinculadas
    const refsContainer = document.getElementById('modal-sf-refs-list');
    if (refsContainer) {
      if (sf.artigosIds && sf.artigosIds.length > 0) {
        refsContainer.innerHTML = sf.artigosIds.map(artId => {
          const art = state.articlesMap.get(artId);
          const authorYear = art ? `${art.authors.split(';')[0]} (${art.year || 's.d.'})` : `#${artId}`;
          return `<span class="sf-link-chip" title="${art ? escapeHtml(art.title) : ''}"><i class="fa-solid fa-book-bookmark"></i> #${artId} - ${escapeHtml(authorYear)}</span>`;
        }).join('');
      } else {
        refsContainer.innerHTML = '<span style="color: var(--text-muted); font-size: 0.8rem;">Nenhum artigo científico vinculado diretamente.</span>';
      }
    }

    modal.style.display = 'flex';
  }

  function closeSfPreviewModal() {
    const modal = document.getElementById('modal-sf-preview');
    if (modal) modal.style.display = 'none';
  }

  // ========================================================================
  // EVENT LISTENERS SETUP
  // ========================================================================
  function setupEventListeners() {
    // 1. Tab Navigation (SF, Mapa das SF, Referências)
    const tabSf = document.getElementById('tab-btn-sf');
    const tabMap = document.getElementById('tab-btn-map');
    const tabScopus = document.getElementById('tab-btn-scopus');
    const viewSf = document.getElementById('view-sf');
    const viewMap = document.getElementById('view-map');
    const viewScopus = document.getElementById('view-scopus');

    function switchTab(targetTab) {
      state.activeTab = targetTab;

      if (tabSf) tabSf.classList.toggle('active', targetTab === 'sf');
      if (tabMap) tabMap.classList.toggle('active', targetTab === 'map');
      if (tabScopus) tabScopus.classList.toggle('active', targetTab === 'scopus');

      if (viewSf) viewSf.style.display = targetTab === 'sf' ? 'flex' : 'none';
      if (viewMap) viewMap.style.display = targetTab === 'map' ? 'flex' : 'none';
      if (viewScopus) viewScopus.style.display = targetTab === 'scopus' ? 'flex' : 'none';

      if (targetTab === 'scopus') {
        renderScopusTable();
      } else if (targetTab === 'map') {
        renderSfMatrixMap();
      }
    }

    if (tabSf) {
      tabSf.addEventListener('click', () => switchTab('sf'));
    }

    if (tabMap) {
      tabMap.addEventListener('click', () => switchTab('map'));
    }

    if (tabScopus) {
      tabScopus.addEventListener('click', () => switchTab('scopus'));
    }

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

    // Mobile Drawer de Filtros
    const btnMobileFilters = document.getElementById('btn-mobile-filter-toggle');
    const btnCloseMobileFilters = document.getElementById('btn-close-mobile-filters');
    const btnApplyMobileFilters = document.getElementById('btn-mobile-apply-filters');
    const sidebar = document.getElementById('sf-filters-sidebar');
    const backdrop = document.getElementById('sidebar-backdrop');

    function openMobileFilters() {
      if (sidebar) sidebar.classList.add('mobile-open');
      if (backdrop) backdrop.classList.add('active');
      document.body.classList.add('mobile-drawer-open');
    }

    function closeMobileFilters() {
      if (sidebar) sidebar.classList.remove('mobile-open');
      if (backdrop) backdrop.classList.remove('active');
      document.body.classList.remove('mobile-drawer-open');
    }

    if (btnMobileFilters) btnMobileFilters.addEventListener('click', openMobileFilters);
    if (btnCloseMobileFilters) btnCloseMobileFilters.addEventListener('click', closeMobileFilters);
    if (btnApplyMobileFilters) btnApplyMobileFilters.addEventListener('click', closeMobileFilters);
    if (backdrop) backdrop.addEventListener('click', closeMobileFilters);

    // 4. SF List Sorting Selector
    const sfListSort = document.getElementById('sf-list-sort');
    if (sfListSort) {
      sfListSort.addEventListener('change', (e) => {
        state.listSort = e.target.value;
        sortFilteredSementes();
        renderSfInboxList();
      });
    }

    // 5. Scopus Filters Handlers
    const scopusSearchInput = document.getElementById('scopus-search-input');
    const scopusSearchClear = document.getElementById('scopus-search-clear');

    if (scopusSearchInput) {
      scopusSearchInput.addEventListener('input', (e) => {
        state.scopusFilters.search = e.target.value;
        if (scopusSearchClear) scopusSearchClear.style.display = state.scopusFilters.search ? 'block' : 'none';
        debouncedApplyScopusFilters();
      });
    }

    if (scopusSearchClear) {
      scopusSearchClear.addEventListener('click', () => {
        if (scopusSearchInput) scopusSearchInput.value = '';
        state.scopusFilters.search = '';
        scopusSearchClear.style.display = 'none';
        applyScopusFilters();
      });
    }

    const scopusFilterRel = document.getElementById('scopus-filter-relation');
    if (scopusFilterRel) {
      scopusFilterRel.addEventListener('change', (e) => {
        state.scopusFilters.relation = e.target.value;
        applyScopusFilters();
      });
    }

    const scopusFilterYear = document.getElementById('scopus-filter-year');
    if (scopusFilterYear) {
      scopusFilterYear.addEventListener('change', (e) => {
        state.scopusFilters.year = e.target.value;
        applyScopusFilters();
      });
    }

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

    // 6. Map Toolbar Controls & Filters
    const mapSearchInput = document.getElementById('map-search-input');
    const mapSearchClear = document.getElementById('map-search-clear');
    const mapFilterClass = document.getElementById('map-filter-classificacao');
    const btnResetMap = document.getElementById('map-btn-reset-filters');

    if (mapSearchInput) {
      mapSearchInput.addEventListener('input', (e) => {
        state.mapFilters.search = e.target.value;
        if (mapSearchClear) mapSearchClear.style.display = state.mapFilters.search ? 'block' : 'none';
        debouncedApplyMapFilters();
      });
    }

    if (mapSearchClear) {
      mapSearchClear.addEventListener('click', () => {
        if (mapSearchInput) mapSearchInput.value = '';
        state.mapFilters.search = '';
        mapSearchClear.style.display = 'none';
        applyMapFilters();
      });
    }

    if (mapFilterClass) {
      mapFilterClass.addEventListener('change', (e) => {
        state.mapFilters.classificacao = e.target.value;
        applyMapFilters();
      });
    }

    if (btnResetMap) {
      btnResetMap.addEventListener('click', () => {
        resetMapFilters();
      });
    }

    // Interactive Legend items toggle classificacao filter
    document.querySelectorAll('.map-legend-item').forEach(item => {
      item.addEventListener('click', () => {
        const typ = item.getAttribute('data-legend-typology');
        if (state.mapFilters.classificacao === typ) {
          state.mapFilters.classificacao = '';
          if (mapFilterClass) mapFilterClass.value = '';
        } else {
          state.mapFilters.classificacao = typ;
          if (mapFilterClass) mapFilterClass.value = typ;
        }
        applyMapFilters();
      });
    });

    // 7. Modal SF Preview Handlers
    const btnCloseSfModal = document.getElementById('btn-close-sf-modal');
    const btnSfModalCloseFooter = document.getElementById('btn-sf-modal-close-footer');
    const modalSfPreview = document.getElementById('modal-sf-preview');
    const btnSfModalOpenDossier = document.getElementById('btn-sf-modal-open-dossier');

    if (btnCloseSfModal) btnCloseSfModal.addEventListener('click', closeSfPreviewModal);
    if (btnSfModalCloseFooter) btnSfModalCloseFooter.addEventListener('click', closeSfPreviewModal);
    if (modalSfPreview) {
      modalSfPreview.addEventListener('click', (e) => {
        if (e.target.id === 'modal-sf-preview') closeSfPreviewModal();
      });
    }

    if (btnSfModalOpenDossier) {
      btnSfModalOpenDossier.addEventListener('click', () => {
        const sfId = state.selectedMapSfId;
        closeSfPreviewModal();
        if (sfId) {
          switchTab('sf');
          selectSf(sfId);
          const dossier = document.getElementById('reading-dossier');
          if (dossier) {
            dossier.scrollIntoView({ behavior: 'smooth', block: 'start' });
          }
        }
      });
    }

    // 8. Article Modal Close Handlers
    document.getElementById('btn-close-art-modal').addEventListener('click', closeArticleModal);
    document.getElementById('btn-art-modal-close-footer').addEventListener('click', closeArticleModal);
    document.getElementById('modal-article').addEventListener('click', (e) => {
      if (e.target.id === 'modal-article') closeArticleModal();
    });

    // 9. Keyboard Navigation (Arrows Up / Down for SF Inbox reading & Esc to close modals)
    window.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        closeArticleModal();
        closeSfPreviewModal();
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

    ['eixo', 'classificacao', 'pestel', 'topico'].forEach(group => {
      updateActiveCountBadge(group, 0);
    });

    const mobileBadge = document.getElementById('mobile-filter-badge');
    if (mobileBadge) mobileBadge.style.display = 'none';

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
    if (state.activeTab === 'map') {
      renderSfMatrixMap();
    }
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
        overlay.style.display = 'none';
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
  const debouncedApplyMapFilters = debounce(() => applyMapFilters(), 200);

  // Temporizador de segurança: garante que a tela de carregamento nunca fique travada
  setTimeout(() => {
    hideLoadingOverlay();
  }, 2500);

  // Execute initialization
  init();

});
