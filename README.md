# Scenario Extractor

## Descrição

O Scenario Extractor é uma ferramenta web front-end projetada para processar arquivos XML contendo dados de cenários financeiros. Ele extrai informações relevantes, as consolida, exibe em uma tabela interativa com funcionalidades de ordenação e filtragem, e permite a exportação dos dados filtrados para um arquivo Excel (.xlsx) estilizado.

## Funcionalidades Principais

* Upload de múltiplos arquivos XML via arrastar e soltar (drag-and-drop) ou seleção de arquivos.
* Processamento e extração de dados de formatos XML específicos (atualmente: `pyrEvoDoc` e `Barclays`).
* Consolidação de dados de múltiplos arquivos, tratando duplicatas (baseado em CUSIP/ISIN) e mesclando informações (tipos de documento).
* Exibição dos dados em uma tabela HTML com:
    * Cabeçalho de múltiplas linhas.
    * Primeira coluna fixa (`sticky`).
    * Ordenação por qualquer coluna clicável.
    * Filtragem por Tipo de Produto (dropdown).
    * Filtragem individual por coluna (campo de busca no cabeçalho).
    * Agrupamento visual (cores alternadas) baseado na coluna ordenada.
* Indicadores visuais para filtros ativos.
* Botão para resetar todos os filtros aplicados.
* Exportação dos dados *atualmente visíveis* (filtrados) para um arquivo Excel (.xlsx) com múltiplas planilhas (agrupadas por Tipo de Produto) e formatação (cabeçalhos mesclados, estilos, largura de coluna automática).
* Interface responsiva e com tema escuro (dark mode).

## Como Funciona (Arquitetura Geral)

O projeto segue uma arquitetura front-end baseada em HTML, CSS e JavaScript (usando Módulos ES6):

1.  **`index.html`**: Define a estrutura da página e os elementos da interface do usuário.
2.  **`style.css`**: Controla toda a aparência visual, layout e responsividade.
3.  **JavaScript Módulos (`.js` files)**:
    * **`main.js`**: Orquestra a aplicação, gerencia o estado, manipula eventos da UI e coordena as chamadas para os outros módulos.
    * **`extractor.js`**: Contém a lógica específica para analisar (parse) os diferentes formatos XML e extrair os dados necessários.
    * **`table.js`**: Responsável por renderizar a tabela HTML inicial, atualizar eficientemente as linhas (filtragem/ordenação) via manipulação do DOM, calcular larguras de coluna e aplicar estilos visuais (agrupamento, ícones ativos).
    * **`export.js`**: Lida com a geração do arquivo Excel estilizado usando a biblioteca `xlsx-js-style`.

---

## Detalhamento por Arquivo

### 📄 `index.html`

**Propósito:** Este arquivo define a estrutura semântica e os elementos visuais básicos da página web. Ele serve como o "esqueleto" onde o CSS aplicará estilos e o JavaScript adicionará interatividade e conteúdo dinâmico.

**Estrutura Principal:**

* **`<head>`**:
    * Metadados básicos (`charset`, `title`).
    * Links para fontes externas (Google Fonts - Poppins).
    * Link para a folha de estilos principal (`style.css`).
    * Inclusão da biblioteca externa `xlsx-js-style` via CDN para a funcionalidade de exportação Excel.
* **`<body>`**:
    * **`<div class="container">`**: Contêiner principal para centralizar o conteúdo com margens laterais.
        * **`<header class="main-header">`**: Cabeçalho visual da página com logo (SVG), título (`<h1>Scenario Extractor</h1>`) e subtítulo.
        * **`<div class="panel">`**: Painel principal com fundo distinto que agrupa os controles de interação.
            * **`<div class="drop-zone">`**: Área designada para o upload de arquivos via arrastar e soltar. Contém um `<input type="file" id="file-input">` (escondido) e um `<label>` estilizado que ativa o input.
            * **`<div id="file-list-container">`**: Contêiner (inicialmente escondido) para exibir a lista de nomes dos arquivos selecionados (`#file-list`).
            * **`<div class="controls">`**: Agrupa os botões de ação e filtros.
                * **`<div class="filter-container">`**: Contém o dropdown (`<select id="productTypeFilter">`) para filtrar por "Product Type" (inicialmente escondido).
                * **`<button class="btn btn-primary btnPreviewData">`**: Botão para iniciar o processamento dos arquivos XML (inicialmente desabilitado).
                * **`<button class="btn btn-secondary btnExportExcel">`**: Botão para exportar os dados da tabela para Excel (inicialmente escondido).
                * **`<button class="btn btn-secondary btnResetFilters">`**: Botão para limpar todos os filtros aplicados (inicialmente escondido).
            * **`<div class="message">`**: Área para exibir mensagens de erro ou status para o usuário.
        * **`<div class="loader-container">`**: Contêiner (inicialmente escondido) que exibe uma animação de carregamento (`.loader`) durante o processamento.
        * **`<div class="dataTable-container">`**: Contêiner principal onde a tabela de dados será renderizada pelo JavaScript.
            * **`<div class="dataTable">`**: Div onde a `<table>` gerada dinamicamente será inserida.
    * **`<footer class="main-footer">`**: Rodapé simples da página.
    * **`<script src="main.js" type="module">`**: **Importante:** Carrega o script JavaScript principal (`main.js`). O atributo `type="module"` habilita o uso de `import` e `export` (Módulos ES6) nos arquivos JavaScript.

**Observações:**

* Muitos elementos começam escondidos (`style="display: none;"`) e são exibidos dinamicamente pelo JavaScript conforme necessário (ex: lista de arquivos, botões de exportar/resetar, loader, tabela).
* IDs (`#file-input`, `#file-list`, etc.) são usados para referenciar elementos específicos no JavaScript.
* Classes (`.panel`, `.btn`, `.sticky-col`, etc.) são usadas primariamente para aplicar estilos CSS e, secundariamente, para selecionar grupos de elementos no JavaScript (`querySelectorAll`).

### 📄 `style.css`

**Propósito:** Este arquivo define toda a apresentação visual e o layout da aplicação Scenario Extractor. Ele utiliza variáveis CSS para um tema consistente (dark mode) e aplica estilos aos elementos HTML definidos no `index.html`, além de classes adicionadas dinamicamente pelo JavaScript.

**Estrutura e Detalhamento:**

1.  **Variáveis CSS (`:root`)**:
    * Define cores primárias (`--scissero-blue`, `--dark-bg`, `--panel-bg`, `--text-primary`, etc.).
    * Define a família de fontes (`--font-family`).
    * Define cores de estado (`--success-color`, `--error-color`).
    * Define os fundos para o agrupamento visual das linhas da tabela (`--group-a-bg`, `--group-b-bg`) e suas contrapartes sólidas para a coluna fixa (`--group-a-sticky-bg`, `--group-b-sticky-bg`).
    * **Benefício:** Permite fácil alteração do tema e garante consistência visual em toda a aplicação.

2.  **Reset e Estilos Globais (`*`, `body`)**:
    * `*::before, *::after`: Aplica `box-sizing: border-box` a todos os elementos para um controle de layout mais intuitivo (padding e borda não aumentam o tamanho total). Remove margens e paddings padrão.
    * `body`: Define a fonte padrão, cor de fundo escura, cor de texto principal, altura mínima e esconde o scroll horizontal.

3.  **Layout Principal (`.container`, `.main-header`, `.main-footer`, `.panel`)**:
    * `.container`: Limita a largura máxima do conteúdo e o centraliza na página.
    * `.main-header`, `.main-footer`: Estiliza o cabeçalho e rodapé da página (alinhamento, margens, cores).
    * `.panel`: Estiliza o painel central onde ficam os controles (fundo, borda, sombra, cantos arredondados).

4.  **Componentes de UI**:
    * **Drop Zone (`.drop-zone`, `.drop-zone-label`, `.upload-icon`)**: Estiliza a área de upload com borda tracejada, ícone, texto e efeitos visuais de hover e arrasto (`:hover`, `.drag-over`).
    * **File List (`.file-list-container`, `#file-list`)**: Estiliza a área que exibe os arquivos selecionados (fundo, borda, scroll vertical se necessário).
    * **Botões (`.btn`, `.btn-primary`, `.btn-secondary`, `.btn:disabled`)**: Define a aparência padrão dos botões (cores, fonte, padding, cantos arredondados) e seus estados (hover, desabilitado). Inclui ícones SVG dentro dos botões.
    * **Dropdown (`.custom-select`, `.filter-container`)**: Estiliza o `<select>` para combinar com o tema escuro.
    * **Mensagens (`.message`)**: Estiliza a área de feedback para o usuário, especialmente para erros (fundo vermelho claro, borda, cor).
    * **Loader (`.loader-container`, `.loader`, `@keyframes spin`)**: Define a animação de carregamento circular (spinner).

5.  **Tabela de Dados (`.dataTable-container`, `.dataTable`, `table`, `th`, `td`, `thead`, `tbody`)**:
    * `.dataTable-container`: Adiciona margem acima da tabela.
    * `.dataTable`: Contêiner da tabela com altura máxima e scroll automático (`overflow: auto`). Estiliza a barra de rolagem (`::-webkit-scrollbar`).
    * `table`: Configurações básicas da tabela (largura 100%, bordas colapsadas).
    * `th`, `td`: Define padding, alinhamento e quebra de linha padrão para células de cabeçalho e dados.
    * `thead`: Posiciona o cabeçalho de forma fixa no topo durante o scroll vertical (`position: sticky; top: 0; z-index: 2;`).
    * `th`: Estiliza as células do cabeçalho (fundo, fonte, bordas). Regras específicas para cabeçalhos de categoria (`th.category`).
    * `tbody`: Estilos aplicados ao corpo da tabela.
    * **Coluna Fixa (`.sticky-col`, `thead .sticky-col`, `tbody .sticky-col`)**:
        * Aplica `position: sticky; left: 0;` para fixar a primeira coluna durante o scroll horizontal.
        * Define `z-index` diferentes para garantir que o cabeçalho fixo (`z-index: 4`) fique sobre as células de dados fixas (`z-index: 1` ou `3` no hover) e sobre as colunas que rolam.
        * Garante que o fundo (`background-color`) da coluna fixa no cabeçalho e no corpo seja sólido para cobrir o conteúdo que rola por baixo.
    * **Agrupamento de Linhas (`tbody tr.group-a td`, `tbody tr.group-b td`, `tbody tr.group-a td.sticky-col`, `tbody tr.group-b td.sticky-col`)**:
        * Define os fundos das linhas usando as variáveis `--group-a-bg` (transparente por padrão) e `--group-b-bg` (levemente escuro).
        * Crucialmente, aplica fundos **sólidos** correspondentes (`--group-a-sticky-bg`, `--group-b-sticky-bg`) à célula `.sticky-col` da linha, garantindo que a coluna fixa mantenha a consistência visual com o restante da linha durante o scroll e a ordenação.
    * **Efeitos de Hover (`tbody tr:hover td`, `tbody tr:hover td.sticky-col`)**: Define um fundo sutil para a linha inteira quando o mouse passa sobre ela, incluindo a coluna fixa.

6.  **Filtros e Ícones do Cabeçalho**:
    * `thead th:not(.sticky-col)`: Aplica `position: relative` apenas aos cabeçalhos *não fixos* para permitir o posicionamento absoluto dos ícones internos sem quebrar o `sticky` da primeira coluna.
    * `.filter-icon`: Posiciona o ícone da lupa (&#128269;) no canto direito (`right: 10px;`) da célula do cabeçalho. Define estilo e transições.
    * `.filter-icon.active`: Muda a cor do ícone para azul (`var(--scissero-blue)`) quando o filtro da coluna está ativo (classe adicionada via JavaScript).
    * `::after` (nas regras `thead th`): Posiciona o pseudo-elemento usado para a seta de ordenação (atualmente escondido com `display: none;`). Define o posicionamento (`right: 15px;`) e a aparência (bordas que formam um triângulo). Classes `.sorted-asc::after` e `.sorted-desc::after` controlam qual borda do triângulo é colorida para indicar a direção.
    * `.filter-input-wrapper`: Contêiner para o campo de busca, posicionado relativamente dentro do `<th>`.
    * `.column-filter-input`: Estiliza o campo de texto `<input>` do filtro (fundo, borda, cor, padding).
    * `.clear-filter`: Estiliza e posiciona o botão '×' para limpar o filtro dentro do campo de input, **sem** `z-index` para permitir que a coluna fixa passe por cima dele.

7.  **Larguras de Coluna**:
    * `thead th[data-sort-key]`: Define uma `min-width` base de `150px` para todas as colunas que são ordenáveis/filtráveis. Isso garante um tamanho mínimo inicial. (Regras específicas por coluna foram removidas e a largura final é ajustada via JavaScript).

8.  **Classes de Utilidade**:
    * `.hidden`: Classe simples (`display: none;`) adicionada/removida pelo JavaScript (`table.js`) para esconder/mostrar linhas da tabela durante a filtragem, otimizando a performance.

### 📄 `main.js`

**Propósito:** Este arquivo é o ponto de entrada e o orquestrador principal da aplicação Scenario Extractor. Ele é carregado pelo `index.html` com `type="module"`. Suas responsabilidades incluem:

* Importar funcionalidades dos outros módulos (`extractor.js`, `table.js`, `export.js`).
* Obter referências aos elementos-chave do DOM.
* Gerenciar o estado global da aplicação (dados carregados, estado da ordenação, filtros ativos).
* Definir e anexar os *event listeners* (ouvintes de eventos) para interações do usuário (cliques, arrastar arquivos, digitação).
* Chamar as funções apropriadas dos outros módulos em resposta a esses eventos.
* Controlar o fluxo geral da aplicação (ex: mostrar/esconder o loader, exibir mensagens).

**Detalhamento:**

1.  **Importações (`import ... from ...`)**:
    * Importa funções específicas de cada módulo, tornando a separação de responsabilidades explícita:
        * `extractor.js`: Funções para detectar o formato XML e extrair os dados (`detectXMLFormat`, `extractPyrEvoDocData`, `extractBarclaysData`).
        * `table.js`: Funções para manipulação e renderização da tabela (`renderInitialTable`, `updateTableRows`, `sortTable`, `applyRowGrouping`, `getRawValue`, `updateHeaderUI`, `adjustHeaderWidths`).
        * `export.js`: Função para gerar o arquivo Excel (`exportExcel`).

2.  **Referências do DOM**:
    * Utiliza `document.getElementById` e `document.querySelector` para obter e armazenar referências a todos os elementos HTML interativos (botões, inputs, divs de contêiner, etc.) em constantes para acesso rápido e eficiente.

3.  **Estado da Aplicação (Variáveis Globais)**:
    * `consolidatedData = []`: Array que armazena os objetos de dados extraídos e consolidados dos arquivos XML. É a "fonte da verdade" para a tabela.
    * `headerStructure = []`: Array que armazena a estrutura do cabeçalho da tabela (gerada por `renderInitialTable`), usada principalmente pela função de exportação.
    * `maxAssetsForExport = 0`: Armazena o número máximo de colunas "Asset" encontradas nos dados, usado para renderizar o cabeçalho e exportar corretamente.
    * `currentSort = { key: null, direction: 'asc' }`: Objeto que rastreia a coluna atualmente ordenada e a direção (ascendente/descendente).
    * `columnFilters = {}`: Objeto que armazena os valores de filtro ativos para cada coluna (chave é o `data-sort-key`, valor é o texto do filtro em minúsculas).
    * `tableBody = null`: Armazena uma referência ao elemento `<tbody>` da tabela após a renderização inicial, usado pelas funções de atualização do DOM.
    * `tableRowsMap = new Map()`: Um `Map` que associa o ID único de cada linha de dados (`data-id`, geralmente CUSIP ou ISIN) ao seu elemento `<tr>` correspondente no DOM. Essencial para a atualização eficiente via manipulação do DOM.

4.  **Funções Utilitárias**:
    * `debounce(func, delay)`: Função de ordem superior que recebe uma função e um atraso. Retorna uma nova função que só executará a função original após um período de inatividade (1000ms), evitando execuções excessivas durante a digitação rápida nos filtros.

5.  **Funções Controladoras Principais**:
    * `updateFileList()`: Atualiza a lista visual de arquivos selecionados na UI.
    * `populateProductTypeFilter(data)`: Preenche o dropdown de filtro "Product Type" com os tipos únicos encontrados nos dados carregados.
    * `checkFiltersActive()`: Verifica se *algum* filtro (dropdown ou coluna) está ativo. Usado para decidir se o botão "Resetar" deve ser exibido.
    * `updateResetButtonVisibility()`: Mostra ou esconde o botão "Resetar Filtros" com base no resultado de `checkFiltersActive()`.
    * `resetAllFilters()`: Limpa o estado `columnFilters`, reseta o dropdown `productTypeFilter` para "All", chama `updateTableRows` para reexibir todas as linhas, `updateHeaderUI` para limpar os inputs/ícones do cabeçalho, `applyRowGrouping` para recalcular os grupos, e esconde o botão "Resetar".
    * `previewExtractedXML()`:
        * Função assíncrona principal acionada pelo botão "Process Data".
        * Limpa o estado anterior, mostra o loader.
        * Itera sobre os arquivos selecionados.
        * Para cada arquivo, chama `detectXMLFormat` e a função de extração apropriada (`extract...Data`) do módulo `extractor.js`.
        * Consolida os dados em `consolidatedData` usando um `Map`.
        * Se dados válidos forem encontrados:
            * Calcula `maxAssetsForExport`.
            * Chama `populateProductTypeFilter`.
            * **Importante:** Chama `renderInitialTable` (do `table.js`) para criar a tabela HTML inicial e obter as referências (`tableBody`, `tableRowsMap`).
            * **Importante:** Chama `adjustHeaderWidths` (do `table.js`) para calcular e aplicar as larguras mínimas aos cabeçalhos.
            * Mostra a tabela e os botões "Export Excel".
        * Trata erros e esconde o loader no bloco `finally`.

6.  **Event Listeners**:
    * **Drag and Drop (`dropZone`)**: Listeners para `dragover`, `dragleave`, `drop` que gerenciam a aparência da drop zone e chamam `updateFileList` quando arquivos são soltos.
    * **File Input (`fileInput`)**: Listener para `change` que chama `updateFileList`.
    * **Botões (`btnPreview`, `btnExport`, `btnResetFilters`)**: Listeners `click` que chamam as funções correspondentes (`previewExtractedXML`, `exportExcel` (do `export.js`), `resetAllFilters`).
    * **Product Type Filter (`productTypeFilter`)**: Listener `change` que chama `updateTableRows` para aplicar o filtro principal e `applyRowGrouping`. Também atualiza a visibilidade do botão "Resetar".
    * **Column Filter Input (`document`, 'input')**: Listener `input` (com `debounce`) no documento. Quando o evento ocorre em um `.column-filter-input`:
        * Atualiza o estado `columnFilters`.
        * Chama a versão *debounced* (`debouncedFilterUpdate`) que, por sua vez, chama `updateTableRows`, `applyRowGrouping`, `updateHeaderUI` e `updateResetButtonVisibility`.
    * **Header Click (`document`, 'click')**: Listener `click` delegado ao documento para lidar com cliques no cabeçalho da tabela:
        * **Ícone Lupa (`.filter-icon`)**: Mostra o input de filtro correspondente.
        * **Botão Limpar (`.clear-filter`)**: Limpa o filtro da coluna específica, chama `updateTableRows`, `applyRowGrouping`, `updateHeaderUI` e `updateResetButtonVisibility`.
        * **Input de Filtro (`.column-filter-input`)**: Interrompe a propagação do evento para evitar que a ordenação seja acionada ao clicar no input.
        * **Outra Parte do Cabeçalho (`th[data-sort-key]`)**:
            * Chama `sortTable` (do `table.js`) para reordenar o array `consolidatedData`.
            * Chama `updateTableRows` (com `orderChanged = true`) para reorganizar as linhas `<tr>` no DOM.
            * Chama `applyRowGrouping` para atualizar as cores de fundo.
            * Chama `updateHeaderUI` para atualizar as classes de ordenação no cabeçalho.

### 📄 `extractor.js`

**Propósito:** Este módulo é responsável exclusivamente pela lógica de **análise (parsing) e extração de dados** dos diferentes formatos de arquivos XML suportados pela aplicação. Ele não interage diretamente com o DOM nem gerencia o estado da aplicação; apenas recebe um nó XML como entrada e retorna um objeto de dados padronizado (ou `null` se a extração falhar).

**Detalhamento:**

1.  **Funções Auxiliares (Não Exportadas - "Privadas" ao Módulo)**:
    * **`setDate(strDate)`**:
        * **Entrada:** Uma string de data (esperada no formato YYYY-MM-DD).
        * **Lógica:**
            * Verifica se a string de entrada é válida.
            * Cria um objeto `Date` a partir da string, **importante:** adicionando `'T00:00:00Z'` para garantir que a data seja interpretada como UTC e evitar problemas de fuso horário.
            * Verifica se o objeto `Date` resultante é válido.
            * Extrai o dia (`getUTCDate`), o mês abreviado em inglês (`toLocaleString`) e os dois últimos dígitos do ano (`getUTCFullYear`).
            * Formata a data no padrão `DD-MMM-YY` (ex: `19-Oct-25`).
        * **Saída:** String da data formatada ou string vazia se a entrada for inválida.
    * **`findFirstContent(node, selectors)`**:
        * **Entrada:** Um nó do DOM XML (`node`) e um array de strings (`selectors`) contendo seletores CSS.
        * **Lógica:**
            * Itera sobre o array `selectors`.
            * Para cada `selector`, tenta encontrar o primeiro elemento correspondente dentro do `node` usando `node.querySelector(selector)`.
            * Se um elemento for encontrado e tiver conteúdo de texto (`element.textContent`), retorna o texto após remover espaços em branco extras (`trim()`).
            * Se nenhum seletor encontrar um elemento com conteúdo, retorna uma string vazia.
        * **Saída:** O conteúdo textual do primeiro seletor que encontrar um resultado, ou string vazia. **Benefício:** Torna a extração mais resiliente a pequenas variações na estrutura XML, permitindo buscar um dado em locais alternativos.
    * **`calculateTenor(startDateStr, endDateStr)`**:
        * **Entrada:** Duas strings de data (esperadas no formato YYYY-MM-DD).
        * **Lógica:**
            * Verifica se as strings de entrada são válidas.
            * Cria objetos `Date` a partir das strings.
            * Verifica se os objetos `Date` são válidos.
            * Calcula a diferença de meses entre as duas datas.
            * Formata o resultado: Se for um múltiplo exato de 12 meses (e maior que 12), retorna em anos (ex: `2Y`). Caso contrário, retorna em meses (ex: `6M`, `18M`).
        * **Saída:** String representando o prazo (tenor) ou string vazia se as entradas forem inválidas.

2.  **Funções Exportadas**:
    * **`detectXMLFormat(xmlNode)`**:
        * **Entrada:** O nó raiz do documento XML parseado.
        * **Lógica:**
            * Verifica a existência de elementos específicos que identificam cada formato:
                * Se encontrar `<pyrEvoDoc>`, retorna `'pyrEvoDoc'`.
                * Se encontrar `<priip>`, retorna `'Barclays'`.
            * Se nenhum for encontrado, retorna `'unknown'`.
        * **Saída:** String indicando o formato detectado.
    * **`extractPyrEvoDocData(xmlNode)`**:
        * **Entrada:** O nó raiz de um documento XML no formato `pyrEvoDoc`.
        * **Lógica:**
            * Seleciona nós principais (`<product>`, `<tradableForm>`, `<asset>`).
            * **Funções Auxiliares Internas (aninhadas):** Define várias funções específicas para extrair dados complexos dentro deste formato:
                * `getUnderlying()`: Determina o tipo de ativo subjacente (Single, WorstOf, etc.) e extrai os tickers.
                * `getProducts()`: Extrai o tipo de produto (BREN, REN, Reverse Convertible) e o tenor.
                * `upsideLeverage()`, `upsideCap()`: Extrai dados específicos de produtos BREN/REN.
                * `getCoupon()`: Extrai frequência, nível de barreira e memória dos cupons.
                * `getEarlyStrike()`: Determina se houve "early strike" comparando datas.
                * `detectClient()`: Tenta identificar o cliente (JPM, GS, UBS, 3P) baseado nos nomes de contraparte/dealer.
                * `detectDocType()`: Verifica o tipo do documento (Term Sheet, Final PS, Fact Sheet).
                * `getDetails()`: Extrai detalhes condicionais baseados no tipo de produto (BREN vs. RC), como nível de buffer/barreira, frequência de call, período non-call, e a relação entre barreira de juros e buffer/KI.
                * `findIdentifier()`: Busca por identificadores específicos (CUSIP, ISIN) dentro de uma estrutura repetitiva.
            * **Extração Principal:**
                * Chama `findIdentifier` para obter o CUSIP (obrigatório, retorna `null` se não encontrado).
                * Chama as funções auxiliares internas para extrair todos os campos necessários.
                * Usa `setDate` para formatar todas as datas.
                * Usa `calculateTenor` (a global) para calcular o `callNonCallPeriod`.
            * **Estruturação:** Monta e retorna um objeto padronizado contendo todos os dados extraídos, incluindo `format: 'pyrEvoDoc'` e o `identifier` (CUSIP).
        * **Saída:** Objeto com os dados extraídos ou `null` se o CUSIP não for encontrado.
    * **`extractBarclaysData(xmlNode)`**:
        * **Entrada:** O nó raiz de um documento XML no formato `priip` (Barclays).
        * **Lógica:**
            * Usa `findFirstContent` para extrair os dados diretamente dos seletores CSS correspondentes a cada campo no formato Barclays.
            * Usa `setDate` para formatar as datas.
            * Usa `calculateTenor` (a global) para calcular o `productTenor` e `callNonCallPeriod`.
            * Determina o tipo de ativo (Single/WorstOf) e extrai os nomes.
            * Calcula e formata `couponBarrierLevel` e `detailBufferBarrierLevel`.
            * Preenche campos não existentes neste formato com `"N/A"`.
            * **Estruturação:** Monta e retorna um objeto padronizado similar ao de `pyrEvoDoc`, mas com `format: 'Barclays'` e `identifier` sendo o ISIN.
        * **Saída:** Objeto com os dados extraídos ou `null` se o ISIN não for encontrado.

### 📄 `table.js`

**Propósito:** Este módulo encapsula toda a lógica relacionada à **manipulação, renderização e atualização** da tabela HTML de dados. Ele é responsável por:

* Gerar o HTML inicial da tabela (cabeçalho e corpo).
* Calcular e aplicar larguras de coluna dinamicamente.
* Ordenar os dados brutos (o array `consolidatedData`).
* Atualizar a interface da tabela de forma eficiente (sem recriar tudo) em resposta a filtros e ordenações, usando manipulação direta do DOM.
* Aplicar o agrupamento visual (cores alternadas) às linhas visíveis.
* Atualizar a UI do cabeçalho (ícones, inputs, classes de ordenação).
* Fornecer funções auxiliares para obter valores brutos e comparáveis das células.

**Detalhamento:**

1.  **Funções Auxiliares Exportadas**:
    * **`getSortableValue(value)`**:
        * **Entrada:** Um valor de célula (pode ser string, número, null, etc.).
        * **Lógica:** Converte o valor de entrada em um tipo que pode ser comparado corretamente pela função `sort()` do JavaScript:
            * Retorna `-Infinity` para valores `null`, `undefined`, `""` ou `"N/A"` (garante que fiquem agrupados no início ou fim da ordenação).
            * Detecta datas no formato `DD-MMM-YY`, as converte para timestamp (número) para ordenação cronológica correta.
            * Extrai números de strings (ex: '100%' vira `100`), permitindo ordenação numérica.
            * Como fallback, retorna a string em minúsculas para ordenação alfabética.
        * **Saída:** Um valor (número ou string minúscula) apropriado para comparação.
    * **`getRawValue(row, key)`**:
        * **Entrada:** Um objeto de dados da linha (`row`) e a chave da coluna (`key`, ex: `'productClient'`, `'asset_0'`, `'prodCusip'`).
        * **Lógica:** Acessa e retorna o valor correspondente à chave no objeto da linha.
            * Trata o caso especial das colunas de Asset (`key` começando com `'asset_'`), acessando o índice correto no array `row.assets`.
            * Trata o caso especial da primeira coluna, que pode ter a chave `'identifier'` ou `'prodCusip'`.
            * Retorna string vazia (`""`) se a chave não existir no objeto.
        * **Saída:** O valor bruto da célula como string ou o valor do ativo. Usado principalmente para a filtragem (comparação de texto).

2.  **Funções Principais Exportadas**:
    * **`sortTable(sortKey, currentSort, consolidatedData)`**:
        * **Entrada:** A chave da coluna a ser ordenada (`sortKey`), o objeto de estado da ordenação atual (`currentSort`), e o array completo de dados (`consolidatedData`).
        * **Lógica:**
            * Atualiza o objeto `currentSort` (invertendo a direção ou mudando a chave).
            * Reordena o array `consolidatedData` **no local** (in-place) usando `consolidatedData.sort()`.
            * A função de comparação dentro do `sort()` utiliza `getSortableValue(getRawValue(...))` para obter os valores comparáveis corretos para cada linha e coluna.
        * **Saída:** Nenhuma (modifica `consolidatedData` e `currentSort` diretamente). **Importante:** Esta função *apenas* ordena os dados, ela *não* atualiza o DOM.
    * **`applyRowGrouping(dataTable, currentSort)`**:
        * **Entrada:** A referência ao elemento `<table>` (`dataTable`) e o estado atual da ordenação (`currentSort`).
        * **Lógica:**
            * Encontra o `<tbody>`.
            * Obtém uma lista de todas as linhas (`<tr>`) que **não** estão escondidas (`:not(.hidden)`).
            * Determina o índice (`sortKeyIndex`) da coluna que está atualmente ordenada (necessário porque o `colspan` no cabeçalho bagunça a contagem simples).
            * Itera sobre as linhas *visíveis*.
            * Compara o valor da célula na coluna ordenada com o valor da linha anterior.
            * Aplica alternadamente as classes `group-a` e `group-b` às linhas visíveis para criar o efeito de "zebra striping" agrupado.
            * Remove as classes `group-a` e `group-b` das linhas *escondidas* para evitar problemas visuais se elas forem reexibidas.
        * **Saída:** Nenhuma (modifica as classes das `<tr>` no DOM).
    * **`updateHeaderUI(columnFilters, currentSort = null)`**:
        * **Entrada:** O estado atual dos filtros de coluna (`columnFilters`) e, opcionalmente, o estado atual da ordenação (`currentSort`).
        * **Lógica:** Itera sobre todos os cabeçalhos (`<th>`) que possuem `data-sort-key`:
            * **Ícones Ativos:** Verifica se existe um filtro ativo para a coluna em `columnFilters`. Adiciona/remove a classe `.active` no `.filter-icon`. Se um filtro foi limpo externamente (pelo botão Resetar), garante que o input seja escondido e o título/ícone reapareçam.
            * **Classes de Ordenação:** Se `currentSort` foi fornecido, remove as classes `.sorted-asc`/`.sorted-desc` e adiciona a classe correta ao `<th>` correspondente à `currentSort.key`.
        * **Saída:** Nenhuma (modifica classes e estilos no DOM do `<thead>`).
    * **`updateTableRows(consolidatedData, tableRowsMap, columnFilters, productTypeFilter, orderChanged = false)`**:
        * **Entrada:** O array de dados (`consolidatedData`, potencialmente reordenado), o `Map` de linhas (`tableRowsMap`), os filtros (`columnFilters`, `productTypeFilter`) e um booleano `orderChanged` indicando se a ordem dos dados foi alterada (pela função `sortTable`).
        * **Lógica (Otimização de Performance):**
            * **Filtragem:** Itera sobre `consolidatedData`. Para cada `row`, pega sua `<tr>` correspondente do `tableRowsMap`. Aplica a lógica de filtro (Product Type e Colunas). Adiciona ou remove a classe `.hidden` na `<tr>` apropriada. Mantém um registro (`visibleDataIds`) das linhas que devem ficar visíveis.
            * **Reordenação (se `orderChanged === true`):**
                * Cria um `DocumentFragment` (buffer de DOM).
                * Itera sobre `consolidatedData` (que *já está na nova ordem*).
                * Para cada `row` que deve estar *visível* (`visibleDataIds.includes(...)`), adiciona sua `<tr>` (obtida do `tableRowsMap`) ao `DocumentFragment`.
                * Limpa o `<tbody>` existente (`tbody.innerHTML = '';`).
                * Adiciona o `DocumentFragment` (contendo as linhas visíveis na nova ordem) ao `<tbody>` de uma só vez.
            * **Mensagem "Sem Resultados":** Verifica se alguma linha está visível após a filtragem/reordenação. Se nenhuma estiver, adiciona uma linha especial `<tr><td colspan="...">No data matches...</td></tr>`. Se a mensagem existia e agora há resultados, remove a linha da mensagem.
        * **Saída:** Nenhuma (modifica o DOM do `<tbody>`). **Benefício:** Evita a recriação completa do HTML, tornando filtros e ordenações muito mais rápidos.
    * **`adjustHeaderWidths()`**:
        * **Entrada:** Nenhuma (opera diretamente no DOM).
        * **Lógica:**
            * Seleciona todos os cabeçalhos `<th>` com `data-sort-key`.
            * Define um buffer de espaço (`ICON_SPACE_BUFFER`) para os ícones e espaçamento.
            * Define uma lista (`EXTRA_PADDING_COLS`) de colunas que precisam de padding extra e a quantidade (`EXTRA_PADDING_AMOUNT`).
            * Define a largura mínima base (`baseMinWidth`).
            * Itera sobre cada `<th>`:
                * Mede a largura do texto do título (`titleSpan.scrollWidth`).
                * Calcula a largura total necessária (texto + buffer + padding CSS da célula).
                * Garante que essa largura seja pelo menos `baseMinWidth`.
                * Adiciona `EXTRA_PADDING_AMOUNT` se a coluna estiver na lista `EXTRA_PADDING_COLS`.
                * Aplica a `min-width` final diretamente ao estilo inline do `<th>` (`th.style.minWidth = ...`).
        * **Saída:** Nenhuma (modifica o estilo inline dos `<th>`). **Benefício:** Ajusta automaticamente a largura das colunas para caber o título e os ícones, eliminando a necessidade de `min-width` específicos no CSS.
    * **`renderInitialTable(dataTable, data, maxAssetsForExport, columnFilters, currentSort, productTypeFilter)`**:
        * **Entrada:** Referência ao contêiner da tabela, dados, configurações e estados iniciais.
        * **Lógica:**
            * **Geração de `headerStructure`:** Define a estrutura do cabeçalho (categorias, colunas filhas, chaves de ordenação) com base nos dados carregados (número de assets, presença de BREN/REN, formatos XML).
            * **Geração de HTML (Cabeçalho e Corpo):** Cria as strings HTML para `<thead>` e `<tbody>` usando a `headerStructure` e iterando sobre os `data`. **Importante:** Adiciona `data-id` a cada `<tr>` no corpo.
            * **Inserção no DOM:** Insere o HTML gerado na `dataTable` usando `innerHTML`.
            * **Mapeamento de Linhas:** Seleciona todas as `<tr>` recém-criadas no `<tbody>` e as armazena no `tableRowsMap` (chave=`data-id`, valor=elemento `<tr>`).
            * **Chamadas de UI Inicial:** Chama `updateHeaderUI`, `updateTableRows`, e `applyRowGrouping` para garantir que o estado visual inicial (filtros pré-aplicados, se houver, agrupamento) esteja correto.
        * **Saída:** Um objeto contendo `{ headerStructure, tableBody, tableRowsMap }`, que são salvos no estado global em `main.js`.

### 📄 `export.js`

**Propósito:** Este módulo é dedicado exclusivamente à funcionalidade de **exportar os dados da tabela para um arquivo Excel (.xlsx)**. Ele utiliza a biblioteca `xlsx-js-style` (incluída via CDN no `index.html`) para gerar um arquivo Excel formatado, incluindo múltiplas planilhas, cabeçalhos mesclados, estilos de célula e larguras de coluna automáticas.

**Detalhamento:**

1.  **Importações**:
    * `import { getRawValue } from './table.js'`: Importa a função `getRawValue` do módulo `table.js`. Isso é **crucial** porque a exportação precisa aplicar os mesmos filtros de coluna que estão sendo usados na tabela visível, e `getRawValue` é a função que sabe como obter o valor correto de uma célula para a filtragem.

2.  **Funções Auxiliares (Não Exportadas)**:
    * **`sanitizeSheetName(name)`**:
        * **Entrada:** Uma string (geralmente o `productType`).
        * **Lógica:** Remove caracteres inválidos para nomes de planilhas do Excel ( `\`, `/`, `*`, `?`, `[`, `]`, `:`) e trunca o nome para o limite máximo de 31 caracteres.
        * **Saída:** Uma string segura para ser usada como nome de planilha.

3.  **Função Principal Exportada**:
    * **`exportExcel(consolidatedData, columnFilters, productTypeFilter, maxAssetsForExport)`**:
        * **Entrada:** O array completo de dados (`consolidatedData`), o estado atual dos filtros de coluna (`columnFilters`), a referência ao dropdown de tipo de produto (`productTypeFilter`), e o número máximo de colunas Asset (`maxAssetsForExport`).
        * **Lógica:**
            * **Filtragem:**
                * Cria uma cópia (`dataToExport`) de `consolidatedData`.
                * Aplica o filtro do `productTypeFilter` (se não for "all").
                * Aplica **todos** os `columnFilters` ativos, usando a função importada `getRawValue` para obter os dados corretos de cada linha/coluna para a comparação.
            * **Verificação de Dados Vazios:** Se `dataToExport` ficar vazio após a filtragem, exibe um `alert` para o usuário e interrompe a função.
            * **Agrupamento por Planilha:**
                * Se o `productTypeFilter` estiver em "all", agrupa `dataToExport` por `productType` (usando `reduce`) para criar múltiplas planilhas.
                * Se um `productType` específico estiver selecionado, cria apenas um grupo contendo os dados daquele tipo.
            * **Criação do Workbook:** Inicia um novo workbook Excel (`XLSX.utils.book_new()`).
            * **Loop por Planilha (Grupo):** Itera sobre cada `productType` nos dados agrupados.
                * **Geração da Planilha:**
                    * Obtém os dados (`sheetData`) para a planilha atual. Pula se estiver vazio.
                    * Gera um nome de planilha seguro usando `sanitizeSheetName`.
                    * **Recria `localHeaderStructure`:** Determina dinamicamente a estrutura do cabeçalho *específica para esta planilha*, considerando o formato dos dados (`Barclays` vs `pyrEvoDoc`), a presença de BREN/REN, etc. (lógica similar à `renderInitialTable`).
                    * **Cria Linhas de Cabeçalho:** Gera dois arrays (`headerRow1`, `headerRow2`) representando as duas linhas do cabeçalho, incluindo valores `null` para as células que serão mescladas na primeira linha.
                    * **Converte Dados para AoA:** Mapeia `sheetData` para um "Array de Arrays" (`dataAoA`), onde cada array interno representa uma linha de dados na ordem correta das colunas definidas em `localHeaderStructure`.
                    * **Cria Worksheet:** Usa `XLSX.utils.aoa_to_sheet()` para criar o objeto da planilha a partir dos cabeçalhos e `dataAoA`.
                    * **Define Merges:** Calcula e define as células mescladas para o cabeçalho (`worksheet['!merges']`).
                    * **Aplica Estilos e Calcula Larguras:**
                        * Itera sobre todas as células da planilha (`Object.keys(worksheet)`).
                        * Para cada célula, aplica estilos padrão (`font`, `alignment`, `border`) usando a propriedade `.s` do `xlsx-js-style`.
                        * Aplica estilos especiais (negrito, cor de fundo) às células do cabeçalho (linhas 0 e 1).
                        * Calcula a largura necessária para o conteúdo da célula e mantém o controle da largura máxima necessária para cada coluna (`colWidths`).
                    * **Define Larguras das Colunas:** Aplica as larguras calculadas à planilha (`worksheet['!cols']`).
                    * **Adiciona Planilha ao Workbook:** Anexa a `worksheet` formatada ao `workbook` (`XLSX.utils.book_append_sheet`).
            * **Download:** Após o loop (todas as planilhas foram criadas), dispara o download do arquivo `.xlsx` usando `XLSX.writeFile()`.
        * **Saída:** Nenhuma (inicia o download do arquivo no navegador).
