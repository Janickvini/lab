# SFCSED — Catálogo Executivo de Sementes de Futuro (v3.0.1)
### Estudo Prospectivo: Cenários da Educação Superior 2040 — Resultados Finais Consolidados

Plataforma executiva para consulta, exploração analítica e tomada de decisão estratégica fundamentada nas **Sementes de Futuro (SF)** e em seu ecossistema de evidências bibliográficas (**Referências**).

> [!NOTE]
> **Versão 3.0.1 — Design Executivo com Arquitetura de Contraste Estratégico:**
> Esta versão introduz um design system aprimorado que resolve a fadiga visual e a desorientação do "branco monocromático". Através de um canvas com tonalidades balanceadas, elevação de planos (zoning), cartões de KPI destacados, abas de navegação de alto contraste e bordas semânticas por classificação, a plataforma guia intuitivamente a atenção do gestor para onde ele deve ir.

---

## 🎨 Inovações Visuais e Estratégia de Cores da Versão 3.0.1

1. **Canvas de Apoio e Profundidade de Planos (*Visual Grounding*):**
   * O fundo geral (`--bg-app`) no Modo Claro utiliza um suave tom ardósia (`#edf2f7`), eliminando a sensação de "tudo branco" e fazendo com que os módulos e painéis saltem aos olhos como superfícies legíveis.
   * Diferenciação cromática clara entre as 3 colunas da interface: Coluna 1 de Filtros em tom frio estruturado (`#f8fafc`), Coluna 2 de Sementes em canal cinza suave (`#eef2f6`), e Coluna 3 do Dossiê como folha executiva pura com sombras balanceadas (`#ffffff`).

2. **Abas de Navegação em Segmented Control de Alto Contraste:**
   * A aba ativa destaca-se com azul real vibrante (`#2563eb`), texto branco e badge luminosa com sombra direcionada, informando com clareza instantânea em qual seção o usuário se encontra.
   * O botão do **Mapa das SF** conta com destaque especial em azul céu com badge contrastante.

3. **Cartões de Indicadores Executivos Elevados (KPI Tiles):**
   * Os 4 indicadores do cabeçalho agora são cartões destacados com bordas semânticas à esquerda de 4px:
     * **Sementes Consolidadas:** Borda azul (`#2563eb`) e ícone em caixa azul suave.
     * **Referências Citadas:** Borda roxa (`#7c3aed`) e ícone em caixa violeta.
     * **Eixos Estratégicos:** Borda esmeralda (`#059669`) e ícone em caixa verde.
     * **Tópicos Padronizados:** Borda âmbar (`#d97706`) e ícone em caixa dourada.

4. **Hierarquia e Estado Ativo na Lista de Sementes (Inbox):**
   * Cada cartão de semente flutua individualmente com borda indicativa de sua classificação prospectiva (Ruptura em vermelho, Tendência em azul, Incerteza em âmbar, Fato Portador em esmeralda, Surpresa em laranja e Wild Card em violeta).
   * **Semente Selecionada (Active State):** O cartão selecionado ganha fundo azul límpido (`#eff6ff`), borda azul intensa de 6px à esquerda, sombra expansiva, título em negrito contrastante e uma seta ativa (`>`) à direita.

5. **Dossiê Analítico com Caixas de Destaque Estruturadas:**
   * **Implicações Estratégicas 2040:** Caixa em degradê suave azul-esmeralda com borda lateral azul de 5px.
   * **Indicadores de Monitoramento:** Caixa esmeralda translúcida com ícones de verificação verdes.
   * **Referências Citadas:** Cartões individuais brancos elevados com selos numéricos de citação de alto contraste.

6. **Mapa Matricial PESTEL com Colunas Temáticas:**
   * Cada uma das 6 dimensões PESTEL (Político, Econômico, Social, Tecnológico, Ambiental e Legal) conta com cabeçalho colorido característico e chips de sementes contrastantes.

---

## 📊 Panorama Geral dos Resultados Consolidados

* **61 Sementes de Futuro Consolidadas:** Todas as sementes validadas e estruturadas em dossiês analíticos completos com implicações estratégicas para o horizonte 2040.
* **247 Referências Citadas:** Base bibliográfica de evidências científicas que fundamentam diretamente as sementes de futuro catalogadas.
* **13 Tópicos Padronizados:** Estruturação temática transversal padronizada para consulta orientada.
* **6 Dimensões PESTEL:** Análise de impacto abrangendo os vetores *Político, Econômico, Social, Tecnológico, Ambiental e Legal*.
* **5 Eixos Estratégicos:** Alinhamento com as diretrizes do Ensino Superior Privado até 2040.
* **6 Classificações Prospectivas Finais:** *Tendência, Fato Portador de Futuro, Incerteza, Ruptura, Surpresa inevitável e Wild Card*.

---

## 📁 Estrutura de Arquivos da Versão 3.0.1

* **`index.html`**: Interface web executiva responsiva com visualização em 3 colunas e abas integradas.
* **`style.css`**: Design system sóbrio, executivo e acessível (com contraste aperfeiçoado e zoneamento cromático).
* **`app.js`**: Mecanismo de alta performance para busca, filtros combinados, geração do Mapa Matricial, gráficos analíticos (Chart.js) e atalhos por teclado.
* **`SF CSED.xlsx`**: Planilha oficial de dados consolidados com as abas `Sementes` e `Referências`.
* **`data.js`**: Base compilada com as 61 SFs e 247 referências citadas para inicialização offline instantânea e total resiliência a restrições de CORS.
* **`atualizar_dados.py`**: Script de sincronização que lê `SF CSED.xlsx` e compila `data.js`.
* **`atualizar_dados.command`**: Executável de duplo clique para Mac para sincronização automática imediata.
* **`robots.txt`**: Diretrizes de privacidade para evitar indexação pública em motores de busca.

---

## 🚀 Como Executar

### 1. Uso Local Imediato (Offline)
* Dê um duplo clique no arquivo **`index.html`** no seu navegador de preferência (Safari, Chrome, Firefox, Edge).
* A plataforma carrega instantaneamente todos os dados consolidados via `data.js`.

### 2. Atualização dos Dados após salvar o Excel
* Caso sejam feitas novas edições na planilha `SF CSED.xlsx`:
  1. Dê um **duplo clique no arquivo `atualizar_dados.command`** (ou execute no terminal `python3 atualizar_dados.py`).
  2. O script atualizará o arquivo `data.js` em menos de 1 segundo.
  3. No navegador, pressione **`Cmd + Shift + R`** (ou `F5`) para recarregar com as alterações.

### 3. Publicação em Servidor Web / GitHub Pages
* Ao hospedar a pasta `3.0.1` em servidor web (`http://` ou `https://`):
  * O aplicativo busca dinamicamente o arquivo `SF CSED.xlsx` com leitura em tempo real via SheetJS.
  * O arquivo `data.js` funciona como fallback de alta velocidade.

---

## 🔍 Funcionalidades de Consulta Executiva

1. **Catálogo de Sementes em 3 Colunas (Interface Tipo Leitor):**
   * **Coluna 1 (Filtros Estruturados):** Busca textual livre e seletores múltiplos colapsáveis por *Eixo Estratégico, Classificação Prospectiva, Dimensão PESTEL e Tópico Temático*.
   * **Coluna 2 (Lista Dinâmica):** Navegação rápida entre as 61 sementes, contendo IDs, títulos, badges temáticos, quantidade de referências e notas metodológicas de linhagem. Suporta ordenação por ID, Tópico e Volume de Referências.
   * **Coluna 3 (Dossiê Analítico Completo):** Leitura detalhada com narrativa prospectiva, citações numéricas vinculadas a artigos, implicações estratégicas para o ensino superior privado até 2040, indicadores de monitoramento e lista interativa de referências com resumos inline.
2. **Mapa Matricial das SF (Tópicos × PESTEL):**
   * Matriz cruzada de dupla entrada com **13 Tópicos Padronizados** (linhas) × **6 Dimensões PESTEL** (colunas).
   * 162 ocorrências distribuídas em 56 quadrantes ocupados (de 78 possíveis).
   * Badges coloridos por classificação prospectiva com clique para modal de pré-visualização rápida e atalho para leitura do dossiê completo.
   * Filtro textual dedicado no mapa e legenda interativa de classificação.
3. **Painel Analítico de Gráficos (Chart.js):**
   * Distribuição por Classificação Prospectiva.
   * Incidência das Dimensões PESTEL.
   * Concentração de Sementes por Eixo Estratégico.
4. **Base de Referências Rastreáveis:**
   * Tabela executiva com busca textual, paginação por ano e ordenação dinâmica, listando exclusivamente as 247 referências citadas no estudo.
   * Modal de visualização completa com abstract, autores, periódico, ano, DOI link e lista de SFs fundamentadas.
5. **Versão Responsiva Otimizada para Smartphones:**
   * **Navegação Master-Detail:** No smartphone, a navegação alterna fluidamente entre a lista de sementes e o dossiê em tela cheia com botão de retorno ("← Voltar à Lista").
   * **Drawer de Filtros Off-Canvas:** Painel lateral de filtros desliza como gaveta nativa de aplicativo, com backdrop escuro e botão de aplicação direta.
   * **Visualização Matricial e Tabela com Toque:** Rolagem horizontal fluida com coluna de tópicos e ID fixos (*sticky*) para leitura contínua em telas pequenas.
   * **KPIs em Grid 2x2:** Indicadores executivos compactados para rápida assimilação em dispositivos móveis.
6. **Ergonomia e Acessibilidade:**
   * **Controle de Escala de Fonte (A- / A / A+):** Ajuste em 5 níveis (90% a 130%) com persistência automática.
   * **Navegação por Teclado:** Setas ↑ e ↓ para alternar sementes na lista e tecla Esc para fechar modais.
   * **Alternância de Tema:** Modos Claro e Escuro com paletas corporativas sóbrias de alto contraste.
