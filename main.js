import { detectXMLFormat, extractPyrEvoDocData, extractBarclaysData } from './extractor.js';
import { renderInitialTable, updateTableRows, sortTable, applyRowGrouping, getRawValue, updateHeaderUI, adjustHeaderWidths } from './table.js';
import { exportExcel } from './export.js';

const fileInput = document.getElementById("file-input");
const dropZone = document.querySelector(".drop-zone");
const fileList = document.getElementById("file-list");
const fileListContainer = document.getElementById("file-list-container");
const message = document.querySelector(".message");
const dataTable = document.querySelector(".dataTable");
const dataTableContainer = document.querySelector(".dataTable-container");
const btnPreview = document.querySelector(".btnPreviewData");
const btnExport = document.querySelector(".btnExportExcel");
const productTypeFilter = document.getElementById('productTypeFilter');
const loader = document.querySelector('.loader-container');
const btnResetFilters = document.querySelector('.btnResetFilters');

let consolidatedData = [];
let headerStructure = [];
let maxAssetsForExport = 0;
let currentSort = { key: null, direction: 'asc' };
let columnFilters = {};
let tableBody = null;
let tableRowsMap = new Map();

const debounce = (func, delay) => {
    let timeoutId;
    return (...args) => {
        clearTimeout(timeoutId);
        timeoutId = setTimeout(() => {
            func.apply(this, args);
        }, delay);
    };
};

function updateFileList() {
    const files = fileInput.files;
    if (files.length > 0) {
        fileList.innerHTML = "";
        Array.from(files).forEach(file => {
            const fileItem = document.createElement("div");
            fileItem.textContent = file.name;
            fileList.appendChild(fileItem);
        });
        fileListContainer.style.display = "block";
        btnPreview.disabled = false;
    } else {
        fileListContainer.style.display = "none";
        btnPreview.disabled = true;
    }
}

function populateProductTypeFilter(data) {
    const productTypes = new Set(data.map(item => item.productType).filter(Boolean));
    productTypeFilter.innerHTML = '<option value="all">All Product Types</option>';
    [...productTypes].sort().forEach(type => {
        productTypeFilter.innerHTML += `<option value="${type}">${type}</option>`;
    });
}

function checkFiltersActive() {
    const isProductFilterActive = productTypeFilter.value !== 'all';
    const isColumnFilterActive = Object.values(columnFilters).some(value => value);
    return isProductFilterActive || isColumnFilterActive;
}

function updateResetButtonVisibility() {
    if (checkFiltersActive()) {
        btnResetFilters.style.display = 'inline-flex';
    } else {
        btnResetFilters.style.display = 'none';
    }
}

function resetAllFilters() {
    columnFilters = {};
    productTypeFilter.value = 'all';

    updateTableRows(consolidatedData, tableRowsMap, columnFilters, productTypeFilter);
    updateHeaderUI(columnFilters);
    applyRowGrouping(dataTable, currentSort);

    updateResetButtonVisibility();
}


async function previewExtractedXML() {
    const files = Array.from(fileInput.files);
    if (files.length === 0) {
        message.textContent = "Please select at least one XML file.";
        return;
    }

    message.textContent = "";
    dataTable.innerHTML = "";
    btnExport.style.display = "none";
    btnResetFilters.style.display = 'none';
    document.querySelector('.filter-container').style.display = 'none';
    btnPreview.disabled = true;
    loader.style.display = 'flex';
    dataTableContainer.style.display = 'none';
    currentSort = { key: null, direction: 'asc' };
    columnFilters = {}; 
    tableBody = null;
    tableRowsMap.clear();

    try {
        const productMap = new Map();
        for (const file of files) {
            const parser = new DOMParser();
            const content = await file.text();
            const xmlNode = parser.parseFromString(content, "application/xml");
            if (xmlNode.querySelector('parsererror')) {
                throw new Error(`Error: The file ${file.name} is corrupt or not a valid XML.`);
            }
            const format = detectXMLFormat(xmlNode);
            let itemData;
            if (format === 'pyrEvoDoc') {
                itemData = extractPyrEvoDocData(xmlNode);
            } else if (format === 'Barclays') {
                itemData = extractBarclaysData(xmlNode);
            } else {
                console.warn(`Skipping file ${file.name} due to unknown format.`);
                continue;
            }
            if (!itemData || !itemData.identifier) continue;
            if (productMap.has(itemData.identifier)) {
                const existingItem = productMap.get(itemData.identifier);
                existingItem.termSheet = (existingItem.termSheet === 'Y' || itemData.termSheet === 'Y') ? 'Y' : 'N';
                existingItem.finalPS = (existingItem.finalPS === 'Y' || itemData.finalPS === 'Y') ? 'Y' : 'N';
                existingItem.factSheet = (existingItem.factSheet === 'Y' || itemData.factSheet === 'Y') ? 'Y' : 'N';
            } else {
                productMap.set(itemData.identifier, itemData);
            }
        }
        consolidatedData = Array.from(productMap.values());
        if (consolidatedData.length > 0) {
            maxAssetsForExport = consolidatedData.reduce((max, item) => Math.max(max, item.assets ? item.assets.length : 0), 0);
            
            populateProductTypeFilter(consolidatedData);
            document.querySelector('.filter-container').style.display = 'block';

            const renderResult = renderInitialTable(dataTable, consolidatedData, maxAssetsForExport, columnFilters, currentSort, productTypeFilter);
            headerStructure = renderResult.headerStructure; 
            tableBody = renderResult.tableBody;
            tableRowsMap = renderResult.tableRowsMap;
            
            adjustHeaderWidths(); 
            
            btnExport.style.display = "inline-flex";
            dataTableContainer.style.display = 'block';
        } else {
            message.textContent = `No valid data could be extracted from the selected files.`;
        }
    } catch (error) {
        console.error("An unexpected error occurred:", error);
        message.textContent = error.message || "An unexpected error occurred.";
    } finally {
        loader.style.display = 'none';
        btnPreview.disabled = false;
    }
}

dropZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropZone.classList.add("drag-over");
});
dropZone.addEventListener("dragleave", () => dropZone.classList.remove("drag-over"));
dropZone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropZone.classList.remove("drag-over");
    fileInput.files = e.dataTransfer.files;
    updateFileList();
});
fileInput.addEventListener("change", updateFileList);

btnPreview.addEventListener("click", previewExtractedXML);
btnExport.addEventListener("click", () => {
    exportExcel(consolidatedData, columnFilters, productTypeFilter, maxAssetsForExport);
});
btnResetFilters.addEventListener('click', resetAllFilters);

productTypeFilter.addEventListener('change', (event) => {
    if (!tableBody) return;
    updateTableRows(consolidatedData, tableRowsMap, columnFilters, productTypeFilter);
    applyRowGrouping(dataTable, currentSort);
    updateResetButtonVisibility();
});

const debouncedFilterUpdate = debounce(() => {
    if (!tableBody) return;
    updateTableRows(consolidatedData, tableRowsMap, columnFilters, productTypeFilter);
    applyRowGrouping(dataTable, currentSort);
    updateHeaderUI(columnFilters);
    updateResetButtonVisibility();
}, 1000); 

document.addEventListener('input', (event) => {
    const input = event.target.closest('.column-filter-input');
    if (input) {
        const filterKey = input.dataset.filterKey;
        const filterValue = input.value.toLowerCase();
        columnFilters[filterKey] = filterValue;
        debouncedFilterUpdate();
    }
});

document.addEventListener('click', (event) => {
    const filterIcon = event.target.closest('.filter-icon');
    const clearButton = event.target.closest('.clear-filter');
    const filterInput = event.target.closest('.column-filter-input'); 
    const header = event.target.closest('thead th[data-sort-key]');

    if (filterIcon) {
        event.preventDefault();
        event.stopPropagation();
        const th = filterIcon.closest('th');
        th.querySelector('.header-title-text').style.display = 'none';
        th.querySelector('.filter-icon').style.display = 'none';
        th.querySelector('.filter-input-wrapper').style.display = 'flex';
        th.querySelector('.column-filter-input').focus();
    } else if (clearButton) {
        event.preventDefault();
        event.stopPropagation();
        const wrapper = clearButton.closest('.filter-input-wrapper');
        const th = clearButton.closest('th');
        const input = wrapper.querySelector('input');
        const sortKey = input.dataset.filterKey;

        columnFilters[sortKey] = '';
        input.value = '';
        wrapper.style.display = 'none';
        th.querySelector('.header-title-text').style.display = '';
        th.querySelector('.filter-icon').style.display = '';
        
        if(tableBody) {
             updateTableRows(consolidatedData, tableRowsMap, columnFilters, productTypeFilter);
             applyRowGrouping(dataTable, currentSort);
             updateHeaderUI(columnFilters);
             updateResetButtonVisibility();
        }
    } else if (filterInput) { 
        event.stopPropagation(); 
    } else if (header && tableBody) {
        const sortKey = header.dataset.sortKey;
        
        sortTable(sortKey, currentSort, consolidatedData);
        
        updateTableRows(consolidatedData, tableRowsMap, columnFilters, productTypeFilter, true);
        applyRowGrouping(dataTable, currentSort);
        updateHeaderUI(columnFilters, currentSort);
    }
});