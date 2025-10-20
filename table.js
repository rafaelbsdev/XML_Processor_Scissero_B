export const getSortableValue = (value) => {
    if (value === null || value === undefined || value === "" || value === "N/A") {
        return -Infinity;
    }
    const stringValue = String(value);
    const dateMatch = stringValue.match(/^(\d{2})-([A-Za-z]{3})-(\d{2})$/);
    if (dateMatch) {
                const date = new Date(`${dateMatch[2]} ${dateMatch[1]}, 20${dateMatch[3]}`);
        if (!isNaN(date.getTime())) {
            return date.getTime();
        }
    }
        const num = parseFloat(stringValue.replace(/[^0-9.-]+/g, ""));
    if (!isNaN(num)) {
        return num;
    }
    return stringValue.toLowerCase();
};


export const getRawValue = (row, key) => {
    if (key.startsWith('asset_')) {
        const index = parseInt(key.split('_')[1], 10);
        return (row.assets && row.assets[index]) ? row.assets[index] : "";
    }
    if (key === 'identifier' || key === 'prodCusip') {
        return row.identifier || row.prodCusip || "";
    }
        if (key === 'callMonitoringType' || key === 'couponRateAnnualised') {
         return row[key] || "";
    }
    return row[key] || "";
};


export function sortTable(sortKey, currentSort, consolidatedData) {
    if (currentSort.key === sortKey) {
        currentSort.direction = currentSort.direction === 'asc' ? 'desc' : 'asc';
    } else {
        currentSort.key = sortKey;
        currentSort.direction = 'asc';
    }

    consolidatedData.sort((a, b) => {
        let valA, valB;
        valA = getSortableValue(getRawValue(a, sortKey));
        valB = getSortableValue(getRawValue(b, sortKey));

        if (valA === -Infinity && valB === -Infinity) return 0;
        if (valA === -Infinity) return currentSort.direction === 'asc' ? -1 : 1;
        if (valB === -Infinity) return currentSort.direction === 'asc' ? 1 : -1;

        if (valA < valB) return currentSort.direction === 'asc' ? -1 : 1;
        if (valA > valB) return currentSort.direction === 'asc' ? 1 : -1;
        return 0;
    });
}

export function applyRowGrouping(dataTable, currentSort) {
    const tableBody = dataTable.querySelector('tbody');
    if (!tableBody || !currentSort.key) return;

    const rows = Array.from(tableBody.querySelectorAll('tr:not(.hidden):not(.no-results)'));     if (rows.length === 0) return;

    let currentGroupClass = 'group-a';
    let lastValue = null;

    const headerCells = Array.from(document.querySelectorAll('thead th[data-sort-key]'));
    let sortKeyIndex = -1;

        for(let i=0; i< headerCells.length; i++) {
        if(headerCells[i].dataset.sortKey === currentSort.key) {
             let visibleIndex = 0;
            const allHeadersSecondRow = document.querySelectorAll('thead .second-tr th');
            const allHeadersFirstRowSingle = document.querySelectorAll('thead .first-tr th[rowspan="2"]');

            allHeadersFirstRowSingle.forEach(th => {
                 if (th.dataset.sortKey === currentSort.key) {
                     sortKeyIndex = visibleIndex;
                 }
                 visibleIndex++;
            });

            if (sortKeyIndex === -1) {
                allHeadersSecondRow.forEach(th => {
                     if (th.dataset.sortKey === currentSort.key) {
                        sortKeyIndex = visibleIndex;
                     }
                     visibleIndex++;
                });
            }
            break;
        }
    }

    if (sortKeyIndex === -1) {
                                rows.forEach(row => row.classList.remove('group-a', 'group-b'));
        return;
    }

        rows.forEach(row => {
        row.classList.remove('group-a', 'group-b');         const cell = row.cells[sortKeyIndex];
        if (!cell) {
             console.warn("Cell not found at index", sortKeyIndex, "for row", row);
             return;
        };

        const currentValue = cell.textContent;
        if (lastValue !== null && currentValue !== lastValue) {
            currentGroupClass = currentGroupClass === 'group-a' ? 'group-b' : 'group-a';
        }

        row.classList.add(currentGroupClass);
        lastValue = currentValue;
    });

     const hiddenRows = Array.from(tableBody.querySelectorAll('tr.hidden'));
     hiddenRows.forEach(row => row.classList.remove('group-a', 'group-b'));
}

export function updateHeaderUI(columnFilters, currentSort = null) {
     document.querySelectorAll('thead th[data-sort-key]').forEach(th => {
        const sortKey = th.dataset.sortKey;
        const filterIcon = th.querySelector('.filter-icon');
        const filterInput = th.querySelector('.column-filter-input');
        const titleText = th.querySelector('.header-title-text');
        const inputWrapper = th.querySelector('.filter-input-wrapper');

                if (filterIcon) {
            if (columnFilters[sortKey]) {
                filterIcon.classList.add('active');
            } else {
                filterIcon.classList.remove('active');
            }
        }

                if (!columnFilters[sortKey] && filterInput && inputWrapper && inputWrapper.style.display !== 'none') {
            filterInput.value = '';
            inputWrapper.style.display = 'none';
            if (titleText) titleText.style.display = '';
            if (filterIcon) filterIcon.style.display = '';
        }

                th.classList.remove('sorted-asc', 'sorted-desc');
        if (currentSort && currentSort.key === sortKey) {
            th.classList.add(currentSort.direction === 'asc' ? 'sorted-asc' : 'sorted-desc');
        }
     });
}

export function updateTableRows(consolidatedData, tableRowsMap, columnFilters, productTypeFilter, orderChanged = false) {
     const tableBody = document.querySelector(".dataTable tbody");
    if (!tableBody) return;

    const selectedProductType = productTypeFilter.value;
    const activeColumnFilters = Object.entries(columnFilters).filter(([key, value]) => value);
    let visibleDataIds = [];

    consolidatedData.forEach(row => {
        const rowId = row.identifier || row.prodCusip;
        const tr = tableRowsMap.get(rowId);
        if (!tr) {
            console.warn("Row element not found for ID:", rowId);
            return;
        }
        let isVisible = true;
        if (selectedProductType !== 'all' && row.productType !== selectedProductType) { isVisible = false; }
        if (isVisible && activeColumnFilters.length > 0) {
            isVisible = activeColumnFilters.every(([key, value]) => {
                const rawValue = String(getRawValue(row, key)).toLowerCase();
                return rawValue.includes(value);
            });
        }
        if (isVisible) { tr.classList.remove('hidden'); visibleDataIds.push(rowId); }
        else { tr.classList.add('hidden'); }
    });

    if (orderChanged) {
        const fragment = document.createDocumentFragment();
        consolidatedData.forEach(row => {
            const rowId = row.identifier || row.prodCusip;
                        if (visibleDataIds.includes(rowId) && tableRowsMap.has(rowId)) {
                fragment.appendChild(tableRowsMap.get(rowId));
            }
        });
                Array.from(tableBody.querySelectorAll('tr[data-id]')).forEach(tr => tableBody.removeChild(tr));
        tableBody.appendChild(fragment);     }

        const visibleRowCount = visibleDataIds.length;     let noResultsRow = tableBody.querySelector('.no-results');

    if (visibleRowCount === 0) {
         if (!noResultsRow) {
              const currentHeaders = document.querySelectorAll('thead th[data-sort-key]');
              const totalCols = currentHeaders.length;
              noResultsRow = tableBody.insertRow(0);               noResultsRow.classList.add('no-results');
              const cell = noResultsRow.insertCell();
              cell.colSpan = totalCols > 0 ? totalCols : 1;
              cell.style.textAlign = 'center';
              cell.style.padding = '2rem';
              cell.textContent = 'No data matches the current filters.';
         }
    } else if (noResultsRow) {
         tableBody.removeChild(noResultsRow);
    }
}


export function adjustHeaderWidths() {
    const headers = document.querySelectorAll('thead th[data-sort-key]');
    const ICON_SPACE_BUFFER = 55;     const EXTRA_PADDING_COLS = [         'detailCappedUncapped',
        'detailBufferKIBarrier',         'detailBufferBarrierLevel',
        'callMonitoringType',
        'couponRateAnnualised',
        'callNonCallPeriod'
    ];
    const EXTRA_PADDING_AMOUNT = 10;     const baseMinWidth = 150; 
    headers.forEach(th => {
        const titleSpan = th.querySelector('.header-title-text');
        const sortKey = th.dataset.sortKey;

        if (titleSpan && sortKey) {
            const textWidth = titleSpan.scrollWidth;

                        const computedStyle = window.getComputedStyle(th);
            const paddingLeft = parseFloat(computedStyle.paddingLeft);
            const paddingRight = parseFloat(computedStyle.paddingRight);
            const horizontalPadding = paddingLeft + paddingRight;

                        let calculatedWidth = textWidth + ICON_SPACE_BUFFER + horizontalPadding;

                        let widthToApply = Math.max(baseMinWidth, calculatedWidth);

                        if (EXTRA_PADDING_COLS.includes(sortKey)) {
                widthToApply += EXTRA_PADDING_AMOUNT;
            }

                        th.style.minWidth = `${widthToApply}px`;
        }
    });
}


export function renderInitialTable(dataTable, data, maxAssetsForExport, columnFilters, currentSort, productTypeFilter) {

    const showConditionalCols = data.some(row => row.format !== 'Barclays');

        let headerStructure = [];
    let assetHeaders = Array.from({ length: maxAssetsForExport }, (_, i) => ({ title: `Asset ${i + 1}`, sortKey: `asset_${i}`}));
    if (maxAssetsForExport === 1) assetHeaders = [{ title: "Asset", sortKey: 'asset_0' }];
    const hasBrenRenProducts = data.some(row => row.productType === 'BREN' || row.productType === 'REN');
    const formatsInData = new Set(data.map(row => row.format));
    let idColumnTitle = "CUSIP / ISIN";
    let showSecondIsinColumn = false;
    const showDocTypeColumns = showConditionalCols && !(formatsInData.size === 1 && formatsInData.has('Barclays'));
    if (formatsInData.size === 1) {
        if (formatsInData.has('Barclays')) idColumnTitle = "ISIN";
        if (formatsInData.has('pyrEvoDoc')) { idColumnTitle = "CUSIP"; showSecondIsinColumn = true; }
    }
    const firstColSortKey = formatsInData.has('pyrEvoDoc') ? 'prodCusip' : 'identifier';
    let initialHeaders = [{ title: idColumnTitle, isSticky: true, sortKey: firstColSortKey }];
    if (showSecondIsinColumn) { initialHeaders.push({ title: "ISIN", sortKey: 'prodIsin' }); }

    const detailsChildren = [{ title: "Upside Cap", sortKey: 'upsideCap'}, { title: "Upside Leverage", sortKey: 'upsideLeverage'}];
    if (hasBrenRenProducts) { detailsChildren.push({ title: "Capped / Uncapped", sortKey: 'detailCappedUncapped'}); }
    detailsChildren.push({ title: "Buffer / Barrier", sortKey: 'detailBufferKIBarrier'}, { title: "Barrier/Buffer Level", sortKey: 'detailBufferBarrierLevel'}, { title: "Interest v Barrier/Buffer", sortKey: 'detailInterestBarrierTriggerValue'});

    const callChildren = [{ title: "Frequency", sortKey: 'callFrequency' }, { title: "Non-call period", sortKey: 'callNonCallPeriod' }];
    if (showConditionalCols) {
        callChildren.push({ title: "Call Monitoring Type", sortKey: 'callMonitoringType'});
    }

    const couponChildren = [{ title: "Frequency", sortKey: 'couponFrequency'}, { title: "Barrier Level", sortKey: 'couponBarrierLevel'}, { title: "Memory", sortKey: 'couponMemory'}];
     if (showConditionalCols) {
         couponChildren.push({ title: "Coupon rate annualised", sortKey: 'couponRateAnnualised'});
     }

    headerStructure = [...initialHeaders,
        { title: "Underlying", children: [{ title: "Asset Type", sortKey: 'underlyingAssetType'}, ...assetHeaders] },
        { title: "Product Details", children: [{ title: "Product Type", sortKey: 'productType'}, { title: "Client", sortKey: 'productClient'}, { title: "Tenor", sortKey: 'productTenor'}] },
        { title: "Coupons", children: couponChildren },
        { title: "CALL", children: callChildren },
        { title: "Details", children: detailsChildren },
        { title: "DATES IN BOOKINGS", children: [{ title: "Strike", sortKey: 'dateBookingStrikeDate'}, { title: "Pricing", sortKey: 'dateBookingPricingDate'}, { title: "Maturity", sortKey: 'maturityDate'}, { title: "Valuation", sortKey: 'valuationDate'}, { title: "Early Strike", sortKey: 'earlyStrike'}] }
    ];
    if (showDocTypeColumns) { headerStructure.push({ title: "Doc Type", children: [{ title: "Term Sheet", sortKey: 'termSheet'}, { title: "Final PS", sortKey: 'finalPS'}, { title: "Fact Sheet", sortKey: 'factSheet'}] }); }

        let headerHtml = "<tr class='first-tr'>";
    headerStructure.forEach(({ title, children, isSticky, sortKey }) => {
        const span = children ? `colspan="${children.length}"` : `rowspan="2"`;
        const stickyClass = isSticky ? 'sticky-col' : '';
        const categoryClass = children ? 'category' : '';
        const sortKeyAttr = !children && sortKey ? `data-sort-key="${sortKey}"` : '';
        headerHtml += `<th class="${stickyClass} ${categoryClass}" ${span} ${sortKeyAttr}>`;
        if (!children && sortKey) {
             const filterValue = columnFilters[sortKey] || '';
             const titleStyle = filterValue ? 'style="display:none;"' : '';
             const inputStyle = filterValue ? 'style="display:flex;"' : 'style="display:none;"';
             headerHtml += `<span class="header-title-text" ${titleStyle}>${title.toUpperCase()}</span>`;
             headerHtml += `<span class="filter-icon" title="Filter column" ${titleStyle}>&#128269;</span>`;
             headerHtml += `<div class="filter-input-wrapper" ${inputStyle}><input type="text" class="column-filter-input" data-filter-key="${sortKey}" value="${filterValue}" placeholder="Filter..."><span class="clear-filter" title="Clear filter">×</span></div>`;
        } else { headerHtml += title.toUpperCase(); }
        headerHtml += `</th>`;
    });
    headerHtml += "</tr><tr class='second-tr'>";
    headerStructure.forEach(({ children }) => {
        if (Array.isArray(children)) {
            children.forEach(child => {
                const sortKey = child.sortKey;
                 const filterValue = columnFilters[sortKey] || '';
                 const titleStyle = filterValue ? 'style="display:none;"' : '';
                 const inputStyle = filterValue ? 'style="display:flex;"' : 'style="display:none;"';
                 headerHtml += `<th data-sort-key="${sortKey}">`;
                 headerHtml += `<span class="header-title-text" ${titleStyle}>${child.title}</span>`;
                 headerHtml += `<span class="filter-icon" title="Filter column" ${titleStyle}>&#128269;</span>`;
                 headerHtml += `<div class="filter-input-wrapper" ${inputStyle}><input type="text" class="column-filter-input" data-filter-key="${sortKey}" value="${filterValue}" placeholder="Filter..."><span class="clear-filter" title="Clear filter">×</span></div>`;
                 headerHtml += `</th>`;
            });
        }
    });
    headerHtml += "</tr>";

        let bodyHtml = "";
    if (data.length === 0) {
        const totalCols = headerStructure.reduce((acc, h) => acc + (h.children ? h.children.length : 1), 0);
        bodyHtml += `<tr><td colspan="${totalCols}" style="text-align: center; padding: 2rem;">No data loaded.</td></tr>`;
    } else {
        data.forEach(row => {
            const rowId = row.identifier || row.prodCusip;
            bodyHtml += `<tr data-id="${rowId}" data-product-type="${row.productType || ''}">`;
            bodyHtml += `<td class="sticky-col">${row.prodCusip || row.identifier}</td>`;
            if (showSecondIsinColumn) { bodyHtml += `<td>${row.prodIsin || ""}</td>`; }
            bodyHtml += `<td>${row.underlyingAssetType || ""}</td>`;
            for (let i = 0; i < maxAssetsForExport; i++) { bodyHtml += `<td>${(row.assets && row.assets[i]) ? row.assets[i] : ""}</td>`; }
            bodyHtml += `<td>${row.productType || ""}</td>`; bodyHtml += `<td>${row.productClient || ""}</td>`; bodyHtml += `<td>${row.productTenor || ""}</td>`;
            bodyHtml += `<td>${row.couponFrequency || ""}</td>`; bodyHtml += `<td>${row.couponBarrierLevel || ""}</td>`; bodyHtml += `<td>${row.couponMemory || ""}</td>`;
            if (showConditionalCols) { bodyHtml += `<td>${row.couponRateAnnualised || ""}</td>`; }             bodyHtml += `<td>${row.callFrequency || ""}</td>`; bodyHtml += `<td>${row.callNonCallPeriod || ""}</td>`;
            if (showConditionalCols) { bodyHtml += `<td>${row.callMonitoringType || ""}</td>`; }             bodyHtml += `<td>${row.upsideCap || ""}</td>`; bodyHtml += `<td>${row.upsideLeverage || ""}</td>`;
            if (hasBrenRenProducts) { bodyHtml += `<td>${row.detailCappedUncapped || ""}</td>`; }
            bodyHtml += `<td>${row.detailBufferKIBarrier || ""}</td>`; bodyHtml += `<td>${row.detailBufferBarrierLevel || ""}</td>`; bodyHtml += `<td>${row.detailInterestBarrierTriggerValue || ""}</td>`;
            bodyHtml += `<td>${row.dateBookingStrikeDate || ""}</td>`; bodyHtml += `<td>${row.dateBookingPricingDate || ""}</td>`; bodyHtml += `<td>${row.maturityDate || ""}</td>`; bodyHtml += `<td>${row.valuationDate || ""}</td>`; bodyHtml += `<td>${row.earlyStrike || ""}</td>`;
            if (showDocTypeColumns) { bodyHtml += `<td>${row.termSheet || ""}</td>`; bodyHtml += `<td>${row.finalPS || ""}</td>`; bodyHtml += `<td>${row.factSheet || ""}</td>`; }
            bodyHtml += "</tr>";
        });
    }

        dataTable.innerHTML = `<table><thead>${headerHtml}</thead><tbody>${bodyHtml}</tbody></table>`;

        const tableBody = dataTable.querySelector('tbody');
    const tableRowsMap = new Map();
    tableBody.querySelectorAll('tr[data-id]').forEach(tr => {         tableRowsMap.set(tr.dataset.id, tr);
    });

        updateHeaderUI(columnFilters, currentSort);
        applyRowGrouping(dataTable, currentSort);

        return { headerStructure, tableBody, tableRowsMap };
}