[Leia este documento em Português](README.md)

# Scenario Extractor

## Description

The Scenario Extractor is a front-end web tool designed to process XML files containing financial scenario data. It extracts relevant information, consolidates it, displays it in an interactive table with sorting and filtering functionalities, and allows exporting the filtered data to a stylized Excel (.xlsx) file.

## Main Features

* Upload of multiple XML files via drag-and-drop or file selection.
* Processing and data extraction from specific XML formats (currently: `pyrEvoDoc` and `Barclays`).
* Consolidation of data from multiple files, handling duplicates (based on CUSIP/ISIN) and merging information (document types).
* Display of data in an HTML table with:
    * Multi-row header.
    * Fixed first column (`sticky`).
    * Sorting by any clickable column.
    * Filtering by Product Type (dropdown).
    * Individual column filtering (search field in the header).
    * Visual grouping (alternating colors) based on the sorted column.
* Visual indicators for active filters.
* Button to reset all applied filters.
* Export of *currently visible* (filtered) data to an Excel (.xlsx) file with multiple worksheets (grouped by Product Type) and formatting (merged headers, styles, automatic column width).
* Responsive interface with a dark mode theme.

## How it Works (General Architecture)

The project follows a front-end architecture based on HTML, CSS, and JavaScript (using ES6 Modules):

1.  **`index.html`**: Defines the page structure and user interface elements.
2.  **`style.css`**: Controls all visual appearance, layout, and responsiveness.
3.  **JavaScript Modules (`.js` files)**:
    * **`main.js`**: Orchestrates the application, manages state, handles UI events, and coordinates calls to other modules.
    * **`extractor.js`**: Contains the specific logic to parse the different XML formats and extract the necessary data.
    * **`table.js`**: Responsible for rendering the initial HTML table, efficiently updating rows (filtering/sorting) via DOM manipulation, calculating column widths, and applying visual styles (grouping, active icons).
    * **`export.js`**: Handles the generation of the stylized Excel file using the `xlsx-js-style` library.

---

## File Breakdown

### 📄 `index.html`

**Purpose:** This file defines the semantic structure and basic visual elements of the web page. It serves as the "skeleton" where CSS will apply styles and JavaScript will add interactivity and dynamic content.

**Main Structure:**

* **`<head>`**:
    * Basic metadata (`charset`, `title`).
    * Links to external fonts (Google Fonts - Poppins).
    * Link to the main stylesheet (`style.css`).
    * Inclusion of the external `xlsx-js-style` library via CDN for the Excel export functionality.
* **`<body>`**:
    * **`<div class="container">`**: Main container to center the content with side margins.
        * **`<header class="main-header">`**: Visual page header with logo (SVG), title (`<h1>Scenario Extractor</h1>`), and subtitle.
        * **`<div class="panel">`**: Main panel with a distinct background that groups the interaction controls.
            * **`<div class="drop-zone">`**: Designated area for file upload via drag-and-drop. Contains a hidden `<input type="file" id="file-input">` and a stylized `<label>` that activates the input.
            * **`<div id="file-list-container">`**: Container (initially hidden) to display the list of selected file names (`#file-list`).
            * **`<div class="controls">`**: Groups the action buttons and filters.
                * **`<div class="filter-container">`**: Contains the dropdown (`<select id="productTypeFilter">`) to filter by "Product Type" (initially hidden).
                * **`<button class="btn btn-primary btnPreviewData">`**: Button to start processing the XML files (initially disabled).
                * **`<button class="btn btn-secondary btnExportExcel">`**: Button to export the table data to Excel (initially hidden).
                * **`<button class="btn btn-secondary btnResetFilters">`**: Button to clear all applied filters (initially hidden).
            * **`<div class="message">`**: Area to display error or status messages to the user.
        * **`<div class="loader-container">`**: Container (initially hidden) that displays a loading animation (`.loader`) during processing.
        * **`<div class="dataTable-container">`**: Main container where the data table will be rendered by JavaScript.
            * **`<div class="dataTable">`**: Div where the dynamically generated `<table>` will be inserted.
    * **`<footer class="main-footer">`**: Simple page footer.
    * **`<script src="main.js" type="module">`**: **Important:** Loads the main JavaScript script (`main.js`). The `type="module"` attribute enables the use of `import` and `export` (ES6 Modules) in the JavaScript files.

**Notes:**

* Many elements start hidden (`style="display: none;"`) and are displayed dynamically by JavaScript as needed (e.g., file list, export/reset buttons, loader, table).
* IDs (`#file-input`, `#file-list`, etc.) are used to reference specific elements in JavaScript.
* Classes (`.panel`, `.btn`, `.sticky-col`, etc.) are used primarily to apply CSS styles and, secondarily, to select groups of elements in JavaScript (`querySelectorAll`).

### 📄 `style.css`

**Purpose:** This file defines the entire visual presentation and layout of the Scenario Extractor application. It uses CSS variables for a consistent theme (dark mode) and applies styles to the HTML elements defined in `index.html`, as well as classes added dynamically by JavaScript.

**Structure and Breakdown:**

1.  **CSS Variables (`:root`)**:
    * Defines primary colors (`--scissero-blue`, `--dark-bg`, `--panel-bg`, `--text-primary`, etc.).
    * Defines the font family (`--font-family`).
    * Defines state colors (`--success-color`, `--error-color`).
    * Defines the backgrounds for visual row grouping in the table (`--group-a-bg`, `--group-b-bg`) and their solid counterparts for the fixed column (`--group-a-sticky-bg`, `--group-b-sticky-bg`).
    * **Benefit:** Allows for easy theme changes and ensures visual consistency throughout the application.

2.  **Reset and Global Styles (`*`, `body`)**:
    * `*::before, *::after`: Applies `box-sizing: border-box` to all elements for more intuitive layout control (padding and border do not increase the total size). Removes default margins and paddings.
    * `body`: Defines the default font, dark background color, main text color, minimum height, and hides the horizontal scroll.

3.  **Main Layout (`.container`, `.main-header`, `.main-footer`, `.panel`)**:
    * `.container`: Limits the maximum width of the content and centers it on the page.
    * `.main-header`, `.main-footer`: Styles the page header and footer (alignment, margins, colors).
    * `.panel`: Styles the central panel where the controls are located (background, border, shadow, rounded corners).

4.  **UI Components**:
    * **Drop Zone (`.drop-zone`, `.drop-zone-label`, `.upload-icon`)**: Styles the upload area with a dashed border, icon, text, and visual hover and drag effects (`:hover`, `.drag-over`).
    * **File List (`.file-list-container`, `#file-list`)**: Styles the area that displays the selected files (background, border, vertical scroll if necessary).
    * **Buttons (`.btn`, `.btn-primary`, `.btn-secondary`, `.btn:disabled`)**: Defines the default appearance of buttons (colors, font, padding, rounded corners) and their states (hover, disabled). Includes SVG icons inside the buttons.
    * **Dropdown (`.custom-select`, `.filter-container`)**: Styles the `<select>` element to match the dark theme.
    * **Messages (`.message`)**: Styles the user feedback area, especially for errors (light red background, border, color).
    * **Loader (`.loader-container`, `.loader`, `@keyframes spin`)**: Defines the circular loading animation (spinner).

5.  **Data Table (`.dataTable-container`, `.dataTable`, `table`, `th`, `td`, `thead`, `tbody`)**:
    * `.dataTable-container`: Adds margin above the table.
    * `.dataTable`: Table container with a max height and automatic scroll (`overflow: auto`). Styles the scrollbar (`::-webkit-scrollbar`).
    * `table`: Basic table settings (100% width, collapsed borders).
    * `th`, `td`: Defines default padding, alignment, and line break for header and data cells.
    * `thead`: Positions the header in a fixed position at the top during vertical scroll (`position: sticky; top: 0; z-index: 2;`).
    * `th`: Styles the header cells (background, font, borders). Specific rules for category headers (`th.category`).
    * `tbody`: Styles applied to the table body.
    * **Fixed Column (`.sticky-col`, `thead .sticky-col`, `tbody .sticky-col`)**:
        * Applies `position: sticky; left: 0;` to fix the first column during horizontal scroll.
        * Defines different `z-index` values to ensure the fixed header (`z-index: 4`) stays above the fixed data cells (`z-index: 1` or `3` on hover) and above the scrolling columns.
        * Ensures the `background-color` of the fixed column in the header and body is solid to cover the content scrolling underneath.
    * **Row Grouping (`tbody tr.group-a td`, `tbody tr.group-b td`, `tbody tr.group-a td.sticky-col`, `tbody tr.group-b td.sticky-col`)**:
        * Defines the row backgrounds using the variables `--group-a-bg` (transparent by default) and `--group-b-bg` (slightly dark).
        * Crucially, applies corresponding **solid** backgrounds (`--group-a-sticky-bg`, `--group-b-sticky-bg`) to the `.sticky-col` cell of the row, ensuring the fixed column maintains visual consistency with the rest of the row during scrolling and sorting.
    * **Hover Effects (`tbody tr:hover td`, `tbody tr:hover td.sticky-col`)**: Defines a subtle background for the entire row on mouse hover, including the fixed column.

6.  **Header Filters and Icons**:
    * `thead th:not(.sticky-col)`: Applies `position: relative` only to *non-fixed* headers to allow absolute positioning of internal icons without breaking the first column's `sticky`.
    * `.filter-icon`: Positions the magnifying glass icon (&#128269;) in the right corner (`right: 10px;`) of the header cell. Defines style and transitions.
    * `.filter-icon.active`: Changes the icon color to blue (`var(--scissero-blue)`) when the column's filter is active (class added via JavaScript).
    * `::after` (in `thead th` rules): Positions the pseudo-element used for the sorting arrow (currently hidden with `display: none;`). Defines the position (`right: 15px;`) and appearance (borders that form a triangle). `.sorted-asc::after` and `.sorted-desc::after` classes control which border of the triangle is colored to indicate the direction.
    * `.filter-input-wrapper`: Container for the search field, positioned relatively inside the `<th>`.
    * `.column-filter-input`: Styles the `<input>` text field for the filter (background, border, color, padding).
    * `.clear-filter`: Styles and positions the '×' button to clear the filter inside the input field, **without** a `z-index` to allow the fixed column to pass over it.

7.  **Column Widths**:
    * `thead th[data-sort-key]`: Defines a base `min-width` of `150px` for all sortable/filterable columns. This ensures a minimum initial size. (Specific per-column rules were removed, and the final width is adjusted via JavaScript).

8.  **Utility Classes**:
    * `.hidden`: Simple class (`display: none;`) added/removed by JavaScript (`table.js`) to hide/show table rows during filtering, optimizing performance.

### 📄 `main.js`

**Purpose:** This file is the entry point and main orchestrator of the Scenario Extractor application. It is loaded by `index.html` with `type="module"`. Its responsibilities include:

* Importing functionalities from other modules (`extractor.js`, `table.js`, `export.js`).
* Getting references to key DOM elements.
* Managing the global application state (loaded data, sorting state, active filters).
* Defining and attaching event listeners for user interactions (clicks, dragging files, typing).
* Calling the appropriate functions from other modules in response to these events.
* Controlling the general application flow (e.g., showing/hiding the loader, displaying messages).

**Breakdown:**

1.  **Imports (`import ... from ...`)**:
    * Imports specific functions from each module, making the separation of concerns explicit:
        * `extractor.js`: Functions to detect XML format and extract data (`detectXMLFormat`, `extractPyrEvoDocData`, `extractBarclaysData`).
        * `table.js`: Functions for table manipulation and rendering (`renderInitialTable`, `updateTableRows`, `sortTable`, `applyRowGrouping`, `getRawValue`, `updateHeaderUI`, `adjustHeaderWidths`).
        * `export.js`: Function to generate the Excel file (`exportExcel`).

2.  **DOM References**:
    * Uses `document.getElementById` and `document.querySelector` to get and store references to all interactive HTML elements (buttons, inputs, container divs, etc.) in constants for fast and efficient access.

3.  **Application State (Global Variables)**:
    * `consolidatedData = []`: Array that stores the extracted and consolidated data objects from the XML files. It is the "source of truth" for the table.
    * `headerStructure = []`: Array that stores the table header structure (generated by `renderInitialTable`), used mainly by the export function.
    * `maxAssetsForExport = 0`: Stores the maximum number of "Asset" columns found in the data, used to render the header and export correctly.
    * `currentSort = { key: null, direction: 'asc' }`: Object that tracks the currently sorted column and direction (ascending/descending).
    * `columnFilters = {}`: Object that stores the active filter values for each column (key is `data-sort-key`, value is the filter text in lowercase).
    * `tableBody = null`: Stores a reference to the `<tbody>` element of the table after initial rendering, used by DOM update functions.
    * `tableRowsMap = new Map()`: A `Map` that associates the unique ID of each data row (`data-id`, usually CUSIP or ISIN) with its corresponding `<tr>` element in the DOM. Essential for efficient updates via DOM manipulation.

4.  **Utility Functions**:
    * `debounce(func, delay)`: Higher-order function that takes a function and a delay. Returns a new function that will only execute the original function after a period of inactivity (1000ms), preventing excessive executions during rapid typing in filters.

5.  **Main Controller Functions**:
    * `updateFileList()`: Updates the visual list of selected files in the UI.
    * `populateProductTypeFilter(data)`: Populates the "Product Type" filter dropdown with the unique types found in the loaded data.
    * `checkFiltersActive()`: Checks if *any* filter (dropdown or column) is active. Used to decide if the "Reset" button should be displayed.
    * `updateResetButtonVisibility()`: Shows or hides the "Reset Filters" button based on the result of `checkFiltersActive()`.
    * `resetAllFilters()`: Clears the `columnFilters` state, resets the `productTypeFilter` dropdown to "All", calls `updateTableRows` to re-display all rows, `updateHeaderUI` to clear the header inputs/icons, `applyRowGrouping` to recalculate groups, and hides the "Reset" button.
    * `previewExtractedXML()`:
        * Main asynchronous function triggered by the "Process Data" button.
        * Clears the previous state, shows the loader.
        * Iterates over the selected files.
        * For each file, calls `detectXMLFormat` and the appropriate extraction function (`extract...Data`) from the `extractor.js` module.
        * Consolidates the data into `consolidatedData` using a `Map`.
        * If valid data is found:
            * Calculates `maxAssetsForExport`.
            * Calls `populateProductTypeFilter`.
            * **Important:** Calls `renderInitialTable` (from `table.js`) to create the initial HTML table and get references (`tableBody`, `tableRowsMap`).
            * **Important:** Calls `adjustHeaderWidths` (from `table.js`) to calculate and apply the minimum widths to the headers.
            * Shows the table and the "Export Excel" buttons.
        * Handles errors and hides the loader in the `finally` block.

6.  **Event Listeners**:
    * **Drag and Drop (`dropZone`)**: Listeners for `dragover`, `dragleave`, `drop` that manage the drop zone's appearance and call `updateFileList` when files are dropped.
    * **File Input (`fileInput`)**: Listener for `change` that calls `updateFileList`.
    * **Buttons (`btnPreview`, `btnExport`, `btnResetFilters`)**: `click` listeners that call the corresponding functions (`previewExtractedXML`, `exportExcel` (from `export.js`), `resetAllFilters`).
    * **Product Type Filter (`productTypeFilter`)**: `change` listener that calls `updateTableRows` to apply the main filter and `applyRowGrouping`. Also updates the "Reset" button's visibility.
    * **Column Filter Input (`document`, 'input')**: `input` listener (with `debounce`) on the document. When the event occurs on a `.column-filter-input`:
        * Updates the `columnFilters` state.
        * Calls the *debounced* version (`debouncedFilterUpdate`) which, in turn, calls `updateTableRows`, `applyRowGrouping`, `updateHeaderUI`, and `updateResetButtonVisibility`.
    * **Header Click (`document`, 'click')**: Delegated `click` listener on the document to handle clicks in the table header:
        * **Magnifying Glass Icon (`.filter-icon`)**: Shows the corresponding filter input.
        * **Clear Button (`.clear-filter`)**: Clears the specific column's filter, calls `updateTableRows`, `applyRowGrouping`, `updateHeaderUI`, and `updateResetButtonVisibility`.
        * **Filter Input (`.column-filter-input`)**: Stops event propagation to prevent sorting from being triggered when clicking in the input.
        * **Other Header Part (`th[data-sort-key]`)**:
            * Calls `sortTable` (from `table.js`) to re-sort the `consolidatedData` array.
            * Calls `updateTableRows` (with `orderChanged = true`) to reorganize the `<tr>` rows in the DOM.
            * Calls `applyRowGrouping` to update the background colors.
            * Calls `updateHeaderUI` to update the sorting classes in the header.

### 📄 `extractor.js`

**Purpose:** This module is exclusively responsible for the **parsing and data extraction** logic for the different XML file formats supported by the application. It does not interact directly with the DOM nor manage the application state; it only receives an XML node as input and returns a standardized data object (or `null` if extraction fails).

**Breakdown:**

1.  **Helper Functions (Not Exported - "Private" to the Module)**:
    * **`setDate(strDate)`**:
        * **Input:** A date string (expected in YYYY-MM-DD format).
        * **Logic:**
            * Checks if the input string is valid.
            * Creates a `Date` object from the string, **importantly:** adding `'T00:00:00Z'` to ensure the date is interpreted as UTC and avoid timezone issues.
            * Checks if the resulting `Date` object is valid.
            * Extracts the day (`getUTCDate`), the abbreviated English month (`toLocaleString`), and the last two digits of the year (`getUTCFullYear`).
            * Formats the date in the `DD-MMM-YY` pattern (e.g., `19-Oct-25`).
        * **Output:** Formatted date string or an empty string if the input is invalid.
    * **`findFirstContent(node, selectors)`**:
        * **Input:** An XML DOM node (`node`) and an array of strings (`selectors`) containing CSS selectors.
        * **Logic:**
            * Iterates over the `selectors` array.
            * For each `selector`, tries to find the first matching element within the `node` using `node.querySelector(selector)`.
            * If an element is found and has text content (`element.textContent`), returns the text after trimming extra whitespace (`trim()`).
            * If no selector finds an element with content, returns an empty string.
        * **Output:** The text content of the first selector that finds a result, or an empty string. **Benefit:** Makes the extraction more resilient to minor variations in the XML structure, allowing it to fetch data from alternative locations.
    * **`calculateTenor(startDateStr, endDateStr)`**:
        * **Input:** Two date strings (expected in YYYY-MM-DD format).
        * **Logic:**
            * Checks if the input strings are valid.
            * Creates `Date` objects from the strings.
            * Checks if the `Date` objects are valid.
            * Calculates the difference in months between the two dates.
            * Formats the result: If it's an exact multiple of 12 months (and greater than 12), returns in years (e.g., `2Y`). Otherwise, returns in months (e.g., `6M`, `18M`).
        * **Output:** String representing the term (tenor) or an empty string if the inputs are invalid.

2.  **Exported Functions**:
    * **`detectXMLFormat(xmlNode)`**:
        * **Input:** The root node of the parsed XML document.
        * **Logic:**
            * Checks for the existence of specific elements that identify each format:
                * If it finds `<pyrEvoDoc>`, returns `'pyrEvoDoc'`.
                * If it finds `<priip>`, returns `'Barclays'`.
            * If neither is found, returns `'unknown'`.
        * **Output:** String indicating the detected format.
    * **`extractPyrEvoDocData(xmlNode)`**:
        * **Input:** The root node of an XML document in `pyrEvoDoc` format.
        * **Logic:**
            * Selects main nodes (`<product>`, `<tradableForm>`, `<asset>`).
            * **Internal Helper Functions (nested):** Defines several specific functions to extract complex data within this format:
                * `getUnderlying()`: Determines the underlying asset type (Single, WorstOf, etc.) and extracts tickers.
                * `getProducts()`: Extracts the product type (BREN, REN, Reverse Convertible) and tenor.
                * `upsideLeverage()`, `upsideCap()`: Extracts specific BREN/REN product data.
                * `getCoupon()`: Extracts frequency, barrier level, and memory of coupons.
                * `getEarlyStrike()`: Determines if an "early strike" occurred by comparing dates.
                * `detectClient()`: Tries to identify the client (JPM, GS, UBS, 3P) based on counterparty/dealer names.
                * `detectDocType()`: Checks the document type (Term Sheet, Final PS, Fact Sheet).
                * `getDetails()`: Extracts conditional details based on product type (BREN vs. RC), such as buffer/barrier level, call frequency, non-call period, and the relationship between the interest barrier and the buffer/KI.
                * `findIdentifier()`: Searches for specific identifiers (CUSIP, ISIN) within a repetitive structure.
            * **Main Extraction:**
                * Calls `findIdentifier` to get the CUSIP (required, returns `null` if not found).
                * Calls the internal helper functions to extract all necessary fields.
                * Uses `setDate` to format all dates.
                * Uses the global `calculateTenor` to calculate `callNonCallPeriod`.
            * **Structuring:** Assembles and returns a standardized object containing all extracted data, including `format: 'pyrEvoDoc'` and the `identifier` (CUSIP).
        * **Output:** Object with extracted data or `null` if CUSIP is not found.
    * **`extractBarclaysData(xmlNode)`**:
        * **Input:** The root node of an XML document in `priip` (Barclays) format.
        * **Logic:**
            * Uses `findFirstContent` to extract data directly from CSS selectors corresponding to each field in the Barclays format.
            * Uses `setDate` to format dates.
            * Uses the global `calculateTenor` to calculate `productTenor` and `callNonCallPeriod`.
            * Determines the asset type (Single/WorstOf) and extracts names.
            * Calculates and formats `couponBarrierLevel` and `detailBufferBarrierLevel`.
            * Fills non-existent fields in this format with `"N/A"`.
            * **Structuring:** Assembles and returns a standardized object similar to `pyrEvoDoc`'s, but with `format: 'Barclays'` and `identifier` being the ISIN.
        * **Output:** Object with extracted data or `null` if ISIN is not found.

### 📄 `table.js`

**Purpose:** This module encapsulates all logic related to the **manipulation, rendering, and updating** of the HTML data table. It is responsible for:

* Generating the initial HTML for the table (header and body).
* Calculating and applying column widths dynamically.
* Sorting the raw data (the `consolidatedData` array).
* Efficiently updating the table interface (without recreating everything) in response to filters and sorting, using direct DOM manipulation.
* Applying visual grouping (alternating colors) to visible rows.
* Updating the header UI (icons, inputs, sorting classes).
* Providing helper functions to get raw, comparable values from cells.

**Breakdown:**

1.  **Exported Helper Functions**:
    * **`getSortableValue(value)`**:
        * **Input:** A cell value (can be string, number, null, etc.).
        * **Logic:** Converts the input value into a type that can be correctly compared by JavaScript's `sort()` function:
            * Returns `-Infinity` for `null`, `undefined`, `""`, or `"N/A"` values (ensures they are grouped at the beginning or end of the sort).
            * Detects dates in `DD-MMM-YY` format, converts them to a timestamp (number) for correct chronological sorting.
            * Extracts numbers from strings (e.g., '100%' becomes `100`), allowing numerical sorting.
            * As a fallback, returns the lowercase string for alphabetical sorting.
        * **Output:** A value (number or lowercase string) appropriate for comparison.
    * **`getRawValue(row, key)`**:
        * **Input:** A row data object (`row`) and the column key (`key`, e.g., `'productClient'`, `'asset_0'`, `'prodCusip'`).
        * **Logic:** Accesses and returns the value corresponding to the key in the row object.
            * Handles the special case of Asset columns (`key` starting with `'asset_'`), accessing the correct index in the `row.assets` array.
            * Handles the special case of the first column, which might have the key `'identifier'` or `'prodCusip'`.
            * Returns an empty string (`""`) if the key doesn't exist in the object.
        * **Output:** The raw cell value as a string or the asset value. Used primarily for filtering (text comparison).

2.  **Exported Main Functions**:
    * **`sortTable(sortKey, currentSort, consolidatedData)`**:
        * **Input:** The key of the column to be sorted (`sortKey`), the current sort state object (`currentSort`), and the complete data array (`consolidatedData`).
        * **Logic:**
            * Updates the `currentSort` object (reversing direction or changing the key).
            * Re-sorts the `consolidatedData` array **in-place** using `consolidatedData.sort()`.
            * The comparison function inside `sort()` uses `getSortableValue(getRawValue(...))` to get the correct comparable values for each row and column.
        * **Output:** None (modifies `consolidatedData` and `currentSort` directly). **Important:** This function *only* sorts the data; it does *not* update the DOM.
    * **`applyRowGrouping(dataTable, currentSort)`**:
        * **Input:** Reference to the `<table>` element (`dataTable`) and the current sort state (`currentSort`).
        * **Logic:**
            * Finds the `<tbody>`.
            * Gets a list of all rows (`<tr>`) that are **not** hidden (`:not(.hidden)`).
            * Determines the index (`sortKeyIndex`) of the currently sorted column (necessary because `colspan` in the header messes up a simple count).
            * Iterates over the *visible* rows.
            * Compares the cell value in the sorted column with the value from the previous row.
            * Alternately applies the `group-a` and `group-b` classes to visible rows to create the grouped "zebra striping" effect.
            * Removes `group-a` and `group-b` classes from *hidden* rows to prevent visual issues if they are re-displayed.
        * **Output:** None (modifies the classes of `<tr>` elements in the DOM).
    * **`updateHeaderUI(columnFilters, currentSort = null)`**:
        * **Input:** The current state of column filters (`columnFilters`) and, optionally, the current sort state (`currentSort`).
        * **Logic:** Iterates over all headers (`<th>`) that have a `data-sort-key`:
            * **Active Icons:** Checks if an active filter exists for the column in `columnFilters`. Adds/removes the `.active` class on the `.filter-icon`. If a filter was cleared externally (by the Reset button), ensures the input is hidden and the title/icon reappear.
            * **Sort Classes:** If `currentSort` was provided, removes `.sorted-asc`/`.sorted-desc` classes and adds the correct class to the `<th>` corresponding to `currentSort.key`.
        * **Output:** None (modifies classes and styles in the `<thead>` DOM).
    * **`updateTableRows(consolidatedData, tableRowsMap, columnFilters, productTypeFilter, orderChanged = false)`**:
        * **Input:** The data array (`consolidatedData`, potentially re-sorted), the row `Map` (`tableRowsMap`), the filters (`columnFilters`, `productTypeFilter`), and a boolean `orderChanged` indicating if the data order was changed (by the `sortTable` function).
        * **Logic (Performance Optimization):**
            * **Filtering:** Iterates over `consolidatedData`. For each `row`, gets its corresponding `<tr>` from `tableRowsMap`. Applies the filter logic (Product Type and Columns). Adds or removes the `.hidden` class on the appropriate `<tr>`. Keeps a record (`visibleDataIds`) of rows that should remain visible.
            * **Reordering (if `orderChanged === true`):**
                * Creates a `DocumentFragment` (DOM buffer).
                * Iterates over `consolidatedData` (which is *already in the new order*).
                * For each `row` that should be *visible* (`visibleDataIds.includes(...)`), adds its `<tr>` (obtained from `tableRowsMap`) to the `DocumentFragment`.
                * Clears the existing `<tbody>` (`tbody.innerHTML = '';`).
                * Appends the `DocumentFragment` (containing the visible rows in the new order) to the `<tbody>` all at once.
            * **"No Results" Message:** Checks if any rows are visible after filtering/reordering. If none are, adds a special row `<tr><td colspan="...">No data matches...</td></tr>`. If the message existed and there are now results, removes the message row.
        * **Output:** None (modifies the `<tbody>` DOM). **Benefit:** Avoids complete HTML recreation, making filters and sorting much faster.
    * **`adjustHeaderWidths()`**:
        * **Input:** None (operates directly on the DOM).
        * **Logic:**
            * Selects all `<th>` headers with `data-sort-key`.
            * Defines a space buffer (`ICON_SPACE_BUFFER`) for icons and spacing.
            * Defines a list (`EXTRA_PADDING_COLS`) of columns that need extra padding and the amount (`EXTRA_PADDING_AMOUNT`).
            * Defines the base minimum width (`baseMinWidth`).
            * Iterates over each `<th>`:
                * Measures the width of the title text (`titleSpan.scrollWidth`).
                * Calculates the total required width (text + buffer + cell's CSS padding).
                * Ensures this width is at least `baseMinWidth`.
                * Adds `EXTRA_PADDING_AMOUNT` if the column is in the `EXTRA_PADDING_COLS` list.
                * Applies the final `min-width` directly to the `<th>`'s inline style (`th.style.minWidth = ...`).
        * **Output:** None (modifies the inline style of the `<th>` elements). **Benefit:** Automatically adjusts column widths to fit the title and icons, eliminating the need for specific `min-width` in CSS.
    * **`renderInitialTable(dataTable, data, maxAssetsForExport, columnFilters, currentSort, productTypeFilter)`**:
        * **Input:** Reference to the table container, data, and initial settings/states.
        * **Logic:**
            * **`headerStructure` Generation:** Defines the header structure (categories, child columns, sort keys) based on the loaded data (number of assets, presence of BREN/REN, XML formats).
            * **HTML Generation (Header and Body):** Creates the HTML strings for `<thead>` and `<tbody>` using the `headerStructure` and iterating over the `data`. **Important:** Adds a `data-id` to each `<tr>` in the body.
            * **DOM Insertion:** Inserts the generated HTML into the `dataTable` using `innerHTML`.
            * **Row Mapping:** Selects all newly created `<tr>` elements in the `<tbody>` and stores them in the `tableRowsMap` (key=`data-id`, value=`<tr>` element).
            * **Initial UI Calls:** Calls `updateHeaderUI`, `updateTableRows`, and `applyRowGrouping` to ensure the initial visual state (pre-applied filters, if any, grouping) is correct.
        * **Output:** An object containing `{ headerStructure, tableBody, tableRowsMap }`, which are saved to the global state in `main.js`.

### 📄 `export.js`

**Purpose:** This module is exclusively dedicated to the functionality of **exporting the table data to an Excel (.xlsx) file**. It uses the `xlsx-js-style` library (included via CDN in `index.html`) to generate a formatted Excel file, including multiple worksheets, merged headers, cell styles, and automatic column widths.

**Breakdown:**

1.  **Imports**:
    * `import { getRawValue } from './table.js'`: Imports the `getRawValue` function from the `table.js` module. This is **crucial** because the export needs to apply the same column filters that are being used in the visible table, and `getRawValue` is the function that knows how to get the correct value from a cell for filtering.

2.  **Helper Functions (Not Exported)**:
    * **`sanitizeSheetName(name)`**:
        * **Input:** A string (usually the `productType`).
        * **Logic:** Removes invalid characters for Excel worksheet names ( `\`, `/`, `*`, `?`, `[`, `]`, `:`) and truncates the name to the 31-character limit.
        * **Output:** A safe string to be used as a worksheet name.

3.  **Exported Main Function**:
    * **`exportExcel(consolidatedData, columnFilters, productTypeFilter, maxAssetsForExport)`**:
        * **Input:** The complete data array (`consolidatedData`), the current state of column filters (`columnFilters`), the reference to the product type dropdown (`productTypeFilter`), and the maximum number of Asset columns (`maxAssetsForExport`).
        * **Logic:**
            * **Filtering:**
                * Creates a copy (`dataToExport`) of `consolidatedData`.
                * Applies the `productTypeFilter` (if not "all").
                * Applies **all** active `columnFilters`, using the imported `getRawValue` function to get the correct data from each row/column for comparison.
            * **Empty Data Check:** If `dataToExport` becomes empty after filtering, displays an `alert` to the user and stops the function.
            * **Grouping by Worksheet:**
                * If `productTypeFilter` is on "all", groups `dataToExport` by `productType` (using `reduce`) to create multiple worksheets.
                * If a specific `productType` is selected, creates only one group containing the data for that type.
            * **Workbook Creation:** Initializes a new Excel workbook (`XLSX.utils.book_new()`).
            * **Loop per Worksheet (Group):** Iterates over each `productType` in the grouped data.
                * **Worksheet Generation:**
                    * Gets the `sheetData` for the current worksheet. Skips if empty.
                    * Generates a safe worksheet name using `sanitizeSheetName`.
                    * **Recreates `localHeaderStructure`:** Dynamically determines the header structure *specific to this worksheet*, considering the data format (`Barclays` vs `pyrEvoDoc`), the presence of BREN/REN, etc. (logic similar to `renderInitialTable`).
                    * **Create Header Rows:** Generates two arrays (`headerRow1`, `headerRow2`) representing the two header rows, including `null` values for cells that will be merged in the first row.
                    * **Convert Data to AoA:** Maps `sheetData` to an "Array of Arrays" (`dataAoA`), where each inner array represents a data row in the correct column order defined by `localHeaderStructure`.
                    * **Create Worksheet:** Uses `XLSX.utils.aoa_to_sheet()` to create the worksheet object from the headers and `dataAoA`.
                    * **Define Merges:** Calculates and defines the merged cells for the header (`worksheet['!merges']`).
                    * **Apply Styles and Calculate Widths:**
                        * Iterates over all cells in the worksheet (`Object.keys(worksheet)`).
                        * For each cell, applies default styles (`font`, `alignment`, `border`) using the `.s` property from `xlsx-js-style`.
                        * Applies special styles (bold, background color) to header cells (rows 0 and 1).
                        * Calculates the width needed for the cell's content and keeps track of the maximum width required for each column (`colWidths`).
                    * **Define Column Widths:** Applies the calculated widths to the worksheet (`worksheet['!cols']`).
                    * **Add Worksheet to Workbook:** Appends the formatted `worksheet` to the `workbook` (`XLSX.utils.book_append_sheet`).
            * **Download:** After the loop (all worksheets have been created), triggers the download of the `.xlsx` file using `XLSX.writeFile()`.
        * **Output:** None (initiates the file download in the browser).