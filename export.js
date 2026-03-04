import { getRawValue } from './table.js';

const sanitizeSheetName = (name) => {
    if (!name) return 'Uncategorized';
    return name.replace(/[\\/*?\[\]:]/g, "").substring(0, 31);
};

const createSheet = (sheetData, sheetName, maxAssetsForExport) => {
    if (!sheetData || sheetData.length === 0) {
        console.warn(`No data provided for sheet: ${sheetName}`);
        return null;
    }

    const hasBarclays = sheetData.some(row => row.format === 'Barclays');
    const showCouponMonitoringCols = sheetData.some(row => row.format !== 'Barclays');
    const showDocTypeColumns = showCouponMonitoringCols;

    const format = sheetData[0]?.format;
    const idColumnTitle = format === 'Barclays' ? 'ISIN' : 'CUSIP';
    const isOnlyPyrEvoDoc = sheetData.every(row => row.format === 'pyrEvoDoc');
    const showSecondIsinColumn = (format === 'pyrEvoDoc' && (showCouponMonitoringCols || isOnlyPyrEvoDoc));

    const isBrenRenSheet = sheetData.some(row => row.productType === 'BREN' || row.productType === 'REN');

    let assetHeaders = Array.from({ length: maxAssetsForExport }, (_, i) => `Asset ${i + 1}`);
    if (maxAssetsForExport === 1) assetHeaders = ["Asset"];

    let initialHeaders = [{ title: idColumnTitle }];
    if (showSecondIsinColumn) initialHeaders.push({ title: "ISIN" });

    const detailsChildren = ["Upside Cap", "Upside Leverage"];
    if (isBrenRenSheet) detailsChildren.push("Capped / Uncapped");
    detailsChildren.push("Buffer / Barrier", "Barrier/Buffer Level", "Interest v Barrier/Buffer");
    if (hasBarclays) {
        detailsChildren.push("Programme Name");
        detailsChildren.push("Jurisdiction");
    }

    const callChildren = ["Frequency", "Non-call period"];
    if (showCouponMonitoringCols) { callChildren.push("Call Monitoring Type"); }

    const couponChildren = ["Frequency", "Barrier Level", "Memory"];
    if (showCouponMonitoringCols) { couponChildren.push("Coupon rate annualised"); }

    let localHeaderStructure = [
        ...initialHeaders,
        { title: "Underlying", children: ["Asset Type", ...assetHeaders] },
        { title: "Product Details", children: ["Product Type", "Client", "Tenor"] },
        { title: "Coupons", children: couponChildren },
        { title: "CALL", children: callChildren },
        { title: "Details", children: detailsChildren },
        { title: "DATES IN BOOKINGS", children: ["Strike", "Pricing", "Maturity", "Valuation", "Early Strike"] }
    ];
    if (showDocTypeColumns) {
        localHeaderStructure.push({ title: "Doc Type", children: ["Term Sheet", "Final PS", "Fact Sheet"] });
    }

    const headerRow1 = [], headerRow2 = [];
    localHeaderStructure.forEach(h => {
        headerRow1.push(h.title.toUpperCase());
        if (h.children) {
            headerRow2.push(...h.children);
            for (let i = 1; i < h.children.length; i++) headerRow1.push(null);
        } else {
            headerRow2.push(null);
        }
    });

    const dataAoA = sheetData.map(row => {
        const rowAsArray = [];
        // Iniciais
        rowAsArray.push(row.prodCusip || row.identifier);
        if (showSecondIsinColumn) rowAsArray.push(row.prodIsin || "");
        // Underlying
        rowAsArray.push(row.underlyingAssetType || "");
        for (let i = 0; i < maxAssetsForExport; i++) rowAsArray.push(row.assets && row.assets[i] ? row.assets[i] : "");
        // Product Details
        rowAsArray.push(row.productType || "", row.productClient || "", row.productTenor || "");
        // Coupons
        rowAsArray.push(row.couponFrequency || "", row.couponBarrierLevel || "", row.couponMemory || "");
        if (showCouponMonitoringCols) rowAsArray.push(row.couponRateAnnualised || "");
        // Call
        rowAsArray.push(row.callFrequency || "", row.callNonCallPeriod || "");
        if (showCouponMonitoringCols) rowAsArray.push(row.callMonitoringType || "");
        // Details
        rowAsArray.push(row.upsideCap || "", row.upsideLeverage || "");
        if (isBrenRenSheet) rowAsArray.push(row.detailCappedUncapped || "");
        rowAsArray.push(row.detailBufferKIBarrier || "", row.detailBufferBarrierLevel || "", row.detailInterestBarrierTriggerValue || "");
        if (hasBarclays) {
            rowAsArray.push(row.programmeName || "", row.jurisdiction || "");
        }
        // Dates
        rowAsArray.push(row.dateBookingStrikeDate || "", row.dateBookingPricingDate || "", row.maturityDate || "", row.valuationDate || "", row.earlyStrike || "");
        // Doc Types
        if (showDocTypeColumns) { rowAsArray.push(row.termSheet || "", row.finalPS || "", row.factSheet || ""); }
        return rowAsArray;
    });

    const worksheet = XLSX.utils.aoa_to_sheet([headerRow1, headerRow2, ...dataAoA]);

    const merges = [];
    let colIndex = 0;
    localHeaderStructure.forEach(header => {
        if (header.children) {
            if (header.children.length > 1) merges.push({ s: { r: 0, c: colIndex }, e: { r: 0, c: colIndex + header.children.length - 1 } });
            colIndex += header.children.length;
        } else {
            merges.push({ s: { r: 0, c: colIndex }, e: { r: 1, c: colIndex } });
            colIndex++;
        }
    });
    worksheet['!merges'] = merges;

    const colWidths = [];
    Object.keys(worksheet).forEach(cellAddress => {
        if (cellAddress[0] === '!') return;
        const cell = XLSX.utils.decode_cell(cellAddress);

        worksheet[cellAddress].s = {
            font: { name: "Arial", sz: 10 },
            alignment: { vertical: "center", horizontal: "left", wrapText: false },
            border: { top: { style: "thin" }, right: { style: "thin" }, bottom: { style: "thin" }, left: { style: "thin" } }
        };

        if (cell.r < 2) {
            worksheet[cellAddress].s.font.bold = true;
            worksheet[cellAddress].s.alignment.horizontal = "center";
            worksheet[cellAddress].s.fill = { fgColor: { rgb: "DDEBF7" } };
        }

        const value = worksheet[cellAddress].v;
        const defaultWidth = 12;
        let width = defaultWidth;
        if (value) {
            const stringValue = String(value);
            if (stringValue.match(/^\d{2}-[A-Za-z]{3}-\d{2}$/)) {
                width = 12;
            } else if (!isNaN(value) && typeof value === 'number') {
                 width = Math.max(defaultWidth, stringValue.length + 2);
            } else {
                width = Math.max(defaultWidth, stringValue.length + 4);
            }
        }
         if (cell.r === 1 && headerRow2[cell.c]) {
             width = Math.max(width, String(headerRow2[cell.c]).length + 4);
         }
          if (cell.r === 0 && headerRow1[cell.c] !== null) {
             const mergeInfo = (merges || []).find(m => m.s.r === 0 && m.s.c <= cell.c && m.e.c >= cell.c);
             if (mergeInfo && mergeInfo.e.c > mergeInfo.s.c) {
                  const numCols = mergeInfo.e.c - mergeInfo.s.c + 1;
                  width = Math.max(width, (String(headerRow1[cell.c]).length / numCols) + 4);
             } else if (!mergeInfo || mergeInfo.s.c === mergeInfo.e.c) {
                 width = Math.max(width, String(headerRow1[cell.c]).length + 4);
             }
         }

        if (!colWidths[cell.c] || width > (colWidths[cell.c].wch || 0)) {
            colWidths[cell.c] = { wch: Math.min(Math.ceil(width), 60) };
        }
    });
    worksheet['!cols'] = colWidths;

    return worksheet;
};

export function exportExcel(consolidatedData, columnFilters, productTypeFilter, maxAssetsForExport) {
    let dataToExport = [...consolidatedData];

    const selectedProductType = productTypeFilter.value;
    if (selectedProductType !== 'all') {
        dataToExport = dataToExport.filter(row => row.productType === selectedProductType);
    }
    const activeColumnFilters = Object.entries(columnFilters).filter(([key, value]) => value);
    if (activeColumnFilters.length > 0) {
        dataToExport = dataToExport.filter(row => activeColumnFilters.every(([key, value]) => String(getRawValue(row, key)).toLowerCase().includes(value)));
    }

    if (dataToExport.length === 0) {
        alert("No data to export for the selected filter.");
        return;
    }

    const workbook = XLSX.utils.book_new();

    const allDataSheet = createSheet(dataToExport, "All Data", maxAssetsForExport);
    if (allDataSheet) {
        XLSX.utils.book_append_sheet(workbook, allDataSheet, "All Data");
    }

    let groupedData;
    const uniqueProductTypesInData = new Set(dataToExport.map(row => row.productType || 'Uncategorized'));
    if (selectedProductType === 'all' && uniqueProductTypesInData.size > 1) {
        groupedData = dataToExport.reduce((acc, row) => {
            const key = row.productType || 'Uncategorized';
            if (!acc[key]) acc[key] = [];
            acc[key].push(row);
            return acc;
        }, {});

        const sortedProductTypes = Object.keys(groupedData).sort((a, b) => groupedData[b].length - groupedData[a].length);

        sortedProductTypes.forEach(productType => {
            const sheetData = groupedData[productType];
            const sheetName = sanitizeSheetName(productType);
            const individualSheet = createSheet(sheetData, sheetName, maxAssetsForExport);
            if (individualSheet) {
                XLSX.utils.book_append_sheet(workbook, individualSheet, sheetName);
            }
        });
    }

    XLSX.writeFile(workbook, "ExtractedXML_Data_Global.xlsx");
}