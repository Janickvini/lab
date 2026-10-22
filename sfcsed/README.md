# SFCSED — Plataforma de Sementes de Futuro
### Estudo Prospectivo: Cenários da Educação Superior 2040

Protótipo executivo desenvolvido para visualização analítica, filtragem e curadoria das **Sementes de Futuro (SF)** e sua base de evidências bibliográficas (**Referências**).

---

## 📁 Estrutura de Arquivos

* **`index.html`**: Interface web principal da plataforma (pronta para execução local ou hospedagem no GitHub Pages).
* **`style.css`**: Folha de estilos corporativa e sóbria (suporte nativo a Tema Escuro Executivo e Tema Claro).
* **`app.js`**: Mecanismo de renderização, busca em tempo real, filtros multifacetados, gráficos executivos (Chart.js) e modais de dossiê.
* **`SF CSED.xlsx`**: Planilha fonte com as abas `Sementes` e `Referências`.
* **`data.js`**: Banco de dados compilado para visualização offline local imediata (sem bloqueios de CORS).
* **`atualizar_dados.py`**: Script de sincronização que lê o Excel e regera o `data.js`.
* **`atualizar_dados.command`**: Executável de duplo clique para Mac que atualiza a base automaticamente.
* **`robots.txt`**: Instruções de bloqueio para motores de busca (impede indexação pública no Google).

---

## 🚀 Como Executar

### 1. Uso Local Offline (Duplo Clique)
* Dê um duplo clique no arquivo **`index.html`** no seu Mac.
* A plataforma abrirá diretamente no navegador (Chrome, Safari, Edge) utilizando o snapshot de alta performance `data.js`.

### 2. Atualização dos Dados após editar o Excel (Uso Local no Computador)
* Ao abrir o site localmente via arquivo (`file://`), por segurança padrão dos navegadores, scripts não podem ler arquivos externos do disco diretamente.
* Por isso, toda vez que você alterar e salvar o arquivo **`SF CSED.xlsx`**:
  1. Dê um **duplo clique no arquivo `atualizar_dados.command`** (ou execute `python3 atualizar_dados.py`).
  2. O script atualizará o arquivo `data.js` com a data/hora exata em menos de 1 segundo.
  3. No Firefox/Safari/Chrome, basta recarregar a página com **`Cmd + Shift + R`**. O indicador no topo confirmará o horário exato da sincronização.

### 3. Hospedagem na Nuvem (GitHub Pages / Servidor Web)
* Ao publicar esta pasta em um servidor web ou GitHub Pages (protocolo `http://` / `https://`):
  * O navegador passa a ter permissão para carregar diretamente o arquivo **`SF CSED.xlsx`** via SheetJS com atualização instantânea.
  * O arquivo compilado **`data.js`** continua ativo como garantia de alta performance e redundância imediata caso a planilha demore para baixar.

---

## 🔍 Funcionalidades Disponíveis no Protótipo

1. **Mapa Matricial das SF (Novo na v2.0.0):**
   * Matriz cruzada de dupla entrada com **13 Tópicos Padronizados** (linhas) × **6 Dimensões PESTEL** (colunas: *Político, Econômico, Social, Tecnológico, Ambiental, Legal*).
   * Interseção visual contendo badges interativos das Sementes de Futuro correspondentes a cada cruzamento (208 incidências distribuídas em 58 quadrantes ocupados de 78 no total).
   * Barra de ferramentas com busca em tempo real, filtro rápido por Classificação, filtro por Status de Curadoria e métricas em tempo real.
   * Cabeçalho e coluna lateral fixos (*sticky*) para navegação fluida em telas amplas.
   * Modal de pré-visualização rápida ao clicar em qualquer badge, com metadados, descrição, implicações estratégicas para 2040 e botão com 1 clique para navegar até o dossiê completo na aba principal.
   * Legenda de cores por classificação interativa (ao clicar em uma classificação na legenda, o mapa filtra instantaneamente os badges).
2. **Painel Executivo / KPIs:**
   * Total de sementes (76), referências citadas (315/340), eixos estratégicos e status de revisão concluída.
3. **Painel Analítico Recolhível (Gráficos):**
   * Gráfico de rosca: Distribuição por Classificação (*Ruptura, Tendência, Sinal Fraco, Incerteza, Fato Predeterminado, Surpresa Inevitável, Achado Empírico*).
   * Gráfico de barras: Incidência nas Dimensões PESTEL (*Político, Econômico, Social, Tecnológico, Ambiental, Legal*).
   * Gráfico horizontal: Concentração de sementes por Eixo Estratégico.
4. **Filtros e Busca em Tempo Real:**
   * Busca textual instantânea (pesquisa por ID, termos de título, texto descritivo, implicações ou autores).
   * Filtros combinados por Eixo, Classificação, Dimensão PESTEL, Tópico Temático e Status de Curadoria via checkboxes colapsáveis.
5. **Navegação Integrada em 3 Colunas (Estilo Leitor de E-mail):**
   * **Coluna da Esquerda (Filtros):** Busca textual e seletores por Eixo, Classificação, PESTEL, Tópico e Status de Curadoria.
   * **Coluna Central (Lista de Sementes):** Relação scannável de todas as sementes encontradas com badges de identificação, ordenação dinâmica e indicador de seleção ativa.
   * **Coluna da Direita (Área de Leitura Contínua):** Carrega o dossiê completo imediatamente ao clicar em qualquer semente, contendo narrativa com citações numéricas, implicações para 2040 e lista de referências fundamentadoras com expansão inline de Resumo.
6. **Navegação Rápida por Teclado:**
   * É possível usar as setas **↑** e **↓** do teclado para alternar entre as sementes da lista e realizar a leitura contínua de forma ágil e fluida durante reuniões executivas, além de tecla **Esc** para fechar modais.
7. **Aba de Referências (Artigos):**
   * Navegação independente pelos artigos científicos rastreados, com busca por autor, periódico, ano e identificação das SFs que fundamentam.
8. **Acessibilidade e Conforto Visual (Controle de Fonte A- / A / A+):**
   * Botões no cabeçalho superior que permitem redimensionar todas as fontes em 5 níveis (de 90% a 130%), facilitando a leitura por longos períodos em reuniões ou apresentações com projetores, com persistência automática no navegador.
9. **Alternância de Tema (Modo Escuro / Modo Claro):**
   * Paleta executiva e sóbria projetada para salas de reunião e leitura institucional.
