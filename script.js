const inputXML = document.querySelector(".inputXML");
const message = document.querySelector(".message");
const dataTable = document.querySelector(".dataTable");
const btnPreview = document.querySelector(".btnPreviewData");
const btnExport = document.querySelector(".btnExportExcel");
const productTypeFilter = document.getElementById('productTypeFilter');

let consolidatedData = [];
let headerStructure = [];
let maxAssetsForExport = 0;

btnPreview.addEventListener("click", previewExtractedXML);
btnExport.addEventListener("click", exportExcel);

productTypeFilter.addEventListener('change', (event) => {
    const selectedType = event.target.value;
    const tableRows = document.querySelectorAll('.dataTable tbody tr');
    tableRows.forEach(row => {
        const rowProductType = row.dataset.productType;
        if (selectedType === 'all' || rowProductType === selectedType) {
            row.style.display = '';
        } else {
            row.style.display = 'none';
        }
    });
});

const setDate = (strDate) => {
    if (!strDate) return "";
    const date = new Date(strDate + 'T00:00:00Z');
    if (isNaN(date.getTime())) return "";
    const day = date.getUTCDate().toString().padStart(2, '0');
    const month = date.toLocaleString('en-US', { month: 'short', timeZone: 'UTC' });
    const year = date.getUTCFullYear().toString().slice(-2);
    return `${day}-${month}-${year}`;
};

const uniqueArray = (array, objKey) => {
    const uniqValues = new Set();
    array.forEach((item) => {
        if (item && typeof item === 'object' && item[objKey]) {
            uniqValues.add(item[objKey]);
        }
    });
    return [...uniqValues];
};

const findFirstContent = (node, selectors) => {
    if (!node) return "";
    for (const selector of selectors) {
        const element = node.querySelector(selector);
        if (element && element.textContent) {
            return element.textContent.trim();
        }
    }
    return "";
};

const sanitizeSheetName = (name) => {
    return name.replace(/[\\/*?\[\]:]/g, "").substring(0, 31);
};

const calculateTenor = (startDateStr, endDateStr) => {
    if (!startDateStr || !endDateStr) return "";
    const startDate = new Date(startDateStr);
    const endDate = new Date(endDateStr);
    if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) return "";
    
    let months = (endDate.getFullYear() - startDate.getFullYear()) * 12;
    months -= startDate.getMonth();
    months += endDate.getMonth();
    months = months <= 0 ? 0 : months;

    return months > 12 && months % 12 === 0 ? `${months / 12}Y` : `${months}M`;
};

const detectXMLFormat = (xmlNode) => {
    if (xmlNode.querySelector("pyrEvoDoc")) return 'pyrEvoDoc';
    if (xmlNode.querySelector("priip")) return 'Barclays';
    return 'unknown';
};

const extractPyrEvoDocData = (xmlNode) => {
    const product = xmlNode.querySelector("product");
    const tradableForm = xmlNode.querySelector("tradableForm");
    const asset = xmlNode.querySelector("asset");

    const getUnderlying = (assetNode, key) => {
        const assets = assetNode.querySelectorAll("assets");
        if (typeof key === "number") {
            return findFirstContent(assets[key], ["bloombergTickerSuffix"]);
        } else {
            if (assets.length === 0) return "N/A";
            const assetType = findFirstContent(assets[0], ["assetType"]).replace("Exchange_Traded_Fund", "ETF");
            const basketType = findFirstContent(assetNode, ["basketType"]) || "Multiple";
            return assets.length === 1 ? `Single ${assetType}` : `${basketType} ${assetType}`;
        }
    };
    const getProducts = (productNode, type) => {
        if (type === "tenor") {
            const tenorMonths = findFirstContent(productNode, ["tenor > months"]);
            if (!tenorMonths) return "";
            const months = parseInt(tenorMonths, 10);
            return months > 12 && months % 12 === 0 ? `${months / 12}Y` : `${months}M`;
        }
        return findFirstContent(productNode, ["bufferedReturnEnhancedNote > productType", "reverseConvertible > description", "productName"]);
    };
    const upsideLeverage = (productNode) => {
        const leverage = findFirstContent(productNode, ["bufferedReturnEnhancedNote > upsideLeverage"]);
        return (leverage && `${leverage}X`) || "N/A";
    };
    const upsideCap = (productNode) => {
        const cap = findFirstContent(productNode, ["bufferedReturnEnhancedNote > upsideCap"]);
        return (cap && `${cap}%`) || "N/A";
    };
    const getCoupon = (productNode, type) => {
        const couponSchedule = productNode.querySelector("reverseConvertible > couponSchedule");
        if (!couponSchedule) return "N/A";
        switch (type) {
            case "frequency": return findFirstContent(couponSchedule, ["frequency"]) || "N/A";
            case "level":
                const level = findFirstContent(couponSchedule, ["coupons > contingentLevelLegal > level", "contigentLevelLegal > level"]);
                return (level && `${level}%`) || "N/A";
            case "memory":
                const hasMemory = findFirstContent(couponSchedule, ["hasMemory"]);
                return hasMemory ? (hasMemory === "false" ? "N" : "Y") : "N/A";
            default: return "N/A";
        }
    };
    const getEarlyStrike = (xmlNode) => {
        const strikeRaw = findFirstContent(xmlNode, ["securitized > issuance > prospectusStartDate", "strikeDate > date"]);
        const pricingRaw = findFirstContent(xmlNode, ["securitized > issuance > clientOrderTradeDate"]);
        return strikeRaw && pricingRaw && strikeRaw !== pricingRaw ? "Y" : "N";
    };
    const detectClient = (xmlNode) => {
        const cp = findFirstContent(xmlNode, ["counterparty > name"]);
        const dealer = findFirstContent(xmlNode, ["dealer > name"]);
        const blob = (cp + " " + dealer).toLowerCase();
        if (blob.includes("jpm")) return "JPM PB";
        if (blob.includes("goldman")) return "GS";
        if (blob.includes("bauble") || blob.includes("ubs")) return "UBS";
        return "3P";
    };
    const detectDocType = (xmlNode) => {
        const docType = findFirstContent(xmlNode, ["documentType"]).toUpperCase();
        return {
            termSheet: docType.includes("TERMSHEET") ? "Y" : "N",
            finalPS: docType.includes("PRICING_SUPPLEMENT") ? "Y" : "N",
            factSheet: docType.includes("FACT_SHEET") ? "Y" : "N",
        };
    };
    const getDetails = (productNode, type, tradableFormNode) => {
        const productType = getProducts(productNode);
        if (productType === "BREN" || productType === "REN") {
            switch (type) {
                case "strikelevel": return "Buffer";
                case "bufferlevel":
                    const buffer = findFirstContent(productNode, ["bufferedReturnEnhancedNote > buffer"]);
                    return buffer ? `${100 - parseFloat(buffer)}%` : "";
                case "frequency": return "At Maturity";
                case "noncall": return "N/A";
            }
        }
        
        const isBufferType = parseFloat(findFirstContent(productNode, ["reverseConvertible > strike > level"])) < 100;
        const phoenixType = productNode.querySelector("reverseConvertible > issuerCallable") ? "issuerCallable" : "autocallSchedule";

        switch (type) {
            case "strikelevel":
                return isBufferType ? "Buffer" : "KI Barrier";

            case "bufferlevel":
                if (isBufferType) {
                    const bufferLevelStr = findFirstContent(productNode, ["reverseConvertible > buffer > level"]);
                    return bufferLevelStr ? `${parseFloat(bufferLevelStr)}%` : "";
                } else {
                    const kiBarrierStr = findFirstContent(productNode, ["knockInBarrier > barrierSchedule > barrierLevel > level"]);
                    return kiBarrierStr ? `${parseFloat(kiBarrierStr)}%` : "";
                }

            case "frequency": return findFirstContent(productNode, [`reverseConvertible > ${phoenixType} > barrierSchedule > frequency`]);
            
            case "noncall":
                const issueDateStr = findFirstContent(tradableFormNode, ["securitized > issuance > issueDate"]);
                const firstCallDateStr = findFirstContent(productNode, [`reverseConvertible > ${phoenixType} > barrierSchedule > firstDate`]);
                return calculateTenor(issueDateStr, firstCallDateStr);
            
            default:
                const interestBarrierLevel = parseFloat(findFirstContent(productNode, ["reverseConvertible > couponSchedule > contigentLevelLegal > level"]));
                let comparisonLevel;
                let comparisonLabel;

                if (isBufferType) {
                    comparisonLevel = parseFloat(findFirstContent(productNode, ["reverseConvertible > buffer > level"]));
                    comparisonLabel = "Buffer";
                } else {
                    comparisonLevel = parseFloat(findFirstContent(productNode, ["knockInBarrier > barrierSchedule > barrierLevel > level"]));
                    comparisonLabel = "KI Barrier";
                }

                if (isNaN(interestBarrierLevel) || isNaN(comparisonLevel)) return "N/A";
                
                const comparisonSymbol = interestBarrierLevel > comparisonLevel ? ">" : interestBarrierLevel < comparisonLevel ? "<" : "=";
                return `Interest Barrier ${comparisonSymbol} ${comparisonLabel}`;
        }
    };
    const findIdentifier = (tradableFormNode, type) => {
        const identifiers = tradableFormNode.querySelectorAll('identifiers');
        for (const idNode of identifiers) {
            const typeNode = idNode.querySelector('type');
            if (typeNode && typeNode.textContent.trim().toUpperCase() === type) {
                const codeNode = idNode.querySelector('code');
                return codeNode ? codeNode.textContent.trim() : "";
            }
        }
        return "";
    };

    const cusip = findIdentifier(tradableForm, 'CUSIP');
    if (!cusip) return null;

    const product_type = getProducts(product);
    const { termSheet, finalPS, factSheet } = detectDocType(xmlNode);
    const cappedValue = findFirstContent(product, ['bufferedReturnEnhancedNote > capped', 'capped']);

    return {
        format: 'pyrEvoDoc',
        identifier: cusip,
        prodCusip: cusip,
        prodIsin: findIdentifier(tradableForm, 'ISIN'),
        underlyingAssetType: getUnderlying(asset),
        assets: Array.from(asset.querySelectorAll("assets")).map(node => findFirstContent(node, ["bloombergTickerSuffix"])),
        productType: product_type,
        productClient: detectClient(xmlNode),
        productTenor: getProducts(product, "tenor"),
        couponFrequency: getCoupon(product, "frequency"),
        couponBarrierLevel: getCoupon(product, "level"),
        couponMemory: getCoupon(product, "memory"),
        upsideCap: upsideCap(product),
        upsideLeverage: upsideLeverage(product),
        callFrequency: getDetails(product, "frequency", tradableForm),
        callNonCallPeriod: getDetails(product, "noncall", tradableForm),
        detailCappedUncapped: (product_type === 'BREN' || product_type === 'REN')
            ? (cappedValue === 'true' ? 'Capped' : (cappedValue === 'false' ? 'Uncapped' : ''))
            : '',
        detailBufferKIBarrier: getDetails(product, "strikelevel"),
        detailBufferBarrierLevel: getDetails(product, "bufferlevel", tradableForm),
        detailInterestBarrierTriggerValue: getDetails(product, null, tradableForm),
        dateBookingStrikeDate: setDate(findFirstContent(xmlNode, ["securitized > issuance > prospectusStartDate", "strikeDate > date"])),
        dateBookingPricingDate: setDate(findFirstContent(tradableForm, ["securitized > issuance > clientOrderTradeDate"])),
        maturityDate: setDate(findFirstContent(product, ["maturity > date", "redemptionDate", "settlementDate"])),
        valuationDate: setDate(findFirstContent(product, ["finalObsDate > date", "finalObservation"])),
        earlyStrike: getEarlyStrike(xmlNode),
        termSheet: termSheet,
        finalPS: finalPS,
        factSheet: factSheet,
    };
};

const extractBarclaysData = (xmlNode) => {
    const isin = findFirstContent(xmlNode, ["codes > isin"]);
    if (!isin) return null;

    const issueDateStr = findFirstContent(xmlNode, ["trade > dates > issueDate"]);
    const maturityDateStr = findFirstContent(xmlNode, ["trade > dates > maturityDate"]);
    const assetsNodes = xmlNode.querySelectorAll("referenceAsset > underlyings > item");
    const assetType = findFirstContent(assetsNodes[0], ["type"]);

    const barrierLevelValue = parseFloat(findFirstContent(xmlNode, ["payoff > paymentAtMaturity > knockInBarrierLevelRelative > value"]));
    const formattedBarrierLevel = isNaN(barrierLevelValue) ? "N/A" : `${barrierLevelValue * 100}%`;

    return {
        format: 'Barclays',
        identifier: isin,
        prodCusip: "",
        prodIsin: isin,
        underlyingAssetType: assetsNodes.length > 1 ? `WorstOf ${assetType}` : `Single ${assetType}`,
        assets: Array.from(assetsNodes).map(node => findFirstContent(node, ["name"])),
        productType: findFirstContent(xmlNode, ["product > clientProductType"]),
        productClient: findFirstContent(xmlNode, ["manufacturer > nameShort"]),
        productTenor: calculateTenor(issueDateStr, maturityDateStr),
        couponFrequency: findFirstContent(xmlNode, ["payoff > couponEvents > couponObservationDatesInterval"]) || "N/A",
        couponBarrierLevel: `${parseFloat(findFirstContent(xmlNode, ["payoff > couponEvents > schedule > item > barrierLevelRelative > value"])) * 100}%` || "N/A",
        couponMemory: findFirstContent(xmlNode, ["payoff > couponEvents > schedule > item > memory"]) === 'true' ? 'Y' : 'N',
        upsideCap: "N/A",
        upsideLeverage: "N/A",
        callFrequency: findFirstContent(xmlNode, ["payoff > callEvents > autoCallObservationDatesInterval"]) || "N/A",
        callNonCallPeriod: calculateTenor(issueDateStr, findFirstContent(xmlNode, ["payoff > callEvents > schedule > item > settlementDate"])),
        detailCappedUncapped: "N/A",
        detailBufferKIBarrier: "KI Barrier",
        detailBufferBarrierLevel: formattedBarrierLevel,
        detailInterestBarrierTriggerValue: "N/A",
        dateBookingStrikeDate: setDate(findFirstContent(xmlNode, ["trade > dates > initialValuationDate"])),
        dateBookingPricingDate: setDate(findFirstContent(xmlNode, ["trade > dates > tradeDate"])),
        maturityDate: setDate(maturityDateStr),
        valuationDate: setDate(findFirstContent(xmlNode, ["trade > dates > finalValuationDate"])),
        earlyStrike: "N/A",
        termSheet: "N",
        finalPS: "N",
        factSheet: "N",
    };
};

async function previewExtractedXML() {
    const files = Array.from(inputXML.files).filter((file) => file.name.endsWith(".xml"));
    if (files.length === 0) {
        message.innerHTML = "Please select at least one XML file.";
        return;
    }

    message.innerHTML = `Processing ${files.length} file(s)...`;
    dataTable.innerHTML = "";
    btnExport.style.display = "none";
    document.querySelector('.filter-container').style.display = 'none';
    btnPreview.disabled = true;

    try {
        const productMap = new Map();

        for (const file of files) {
            const parser = new DOMParser();
            const content = await file.text();
            const xmlNode = parser.parseFromString(content, "application/xml");
            
            const parseError = xmlNode.querySelector('parsererror');
            if (parseError) {
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
            maxAssetsForExport = consolidatedData.reduce((max, item) => Math.max(max, item.assets.length), 0);
            renderTable(consolidatedData, maxAssetsForExport);
            message.innerHTML = "";
            btnExport.style.display = "inline-block";
        } else {
            message.innerHTML = `No valid data could be extracted from the selected files.`;
        }

    } catch (error) {
        console.error("An unexpected error occurred:", error);
        message.innerHTML = error.message || "An unexpected error occurred. Check the console for details.";
    } finally {
        btnPreview.disabled = false;
    }
}

function renderTable(data, maxAssets) {
    if (data.length === 0) {
        message.innerHTML = `No data found in the selected XML files.`;
        return;
    }
    const filterContainer = document.querySelector('.filter-container');
    productTypeFilter.innerHTML = ''; 
    const productTypes = uniqueArray(data, 'productType');
    const allOption = document.createElement('option');
    allOption.value = 'all';
    allOption.textContent = 'All Types';
    productTypeFilter.appendChild(allOption);
    productTypes.forEach(type => {
        if (type) {
            const option = document.createElement('option');
            option.value = type;
            option.textContent = type;
            productTypeFilter.appendChild(option);
        }
    });
    filterContainer.style.display = 'inline-block';
    
    let assetHeaders = Array.from({ length: maxAssets }, (_, i) => `Asset ${i + 1}`);
    if (maxAssets === 1) assetHeaders = ["Asset"];
    
    const hasBrenRenProducts = data.some(row => row.productType === 'BREN' || row.productType === 'REN');
    
    const formatsInData = new Set(data.map(row => row.format));
    let idColumnTitle = "CUSIP / ISIN";
    let showSecondIsinColumn = false;

    if (formatsInData.size === 1) {
        if (formatsInData.has('Barclays')) idColumnTitle = "ISIN";
        if (formatsInData.has('pyrEvoDoc')) {
            idColumnTitle = "CUSIP";
            showSecondIsinColumn = true;
        }
    }
    
    let initialHeaders = [{ title: idColumnTitle }];
    if (showSecondIsinColumn) {
        initialHeaders.push({ title: "ISIN" });
    }

    const detailsChildren = ["Upside Cap", "Upside Leverage"];
    if (hasBrenRenProducts) {
        detailsChildren.push("Capped / Uncapped");
    }
    detailsChildren.push("Buffer / Barrier", "Barrier/Buffer Level", "Interest v Barrier/Buffer");

    const callChildren = ["Frequency", "Non-call period"];

    headerStructure = [
        ...initialHeaders,
        { title: "Underlying", children: ["Asset Type", ...assetHeaders] },
        { title: "Product Details", children: ["Product Type", "Client", "Tenor"] },
        { title: "Coupons", children: ["Frequency", "Barrier Level", "Memory"] },
        { title: "CALL", children: callChildren },
        { title: "Details", children: detailsChildren },
        { title: "DATES IN BOOKINGS", children: ["Strike", "Pricing", "Maturity", "Valuation", "Early Strike"] },
        { title: "Doc Type", children: ["Term Sheet", "Final PS", "Fact Sheet"] }
    ];

    let htmlTable = "<table><thead><tr class='first-tr'>";
    headerStructure.forEach(({ title, children }) => {
        const span = children ? `colspan="${children.length}" class="category"` : `rowspan="2"`;
        htmlTable += `<th ${span}>${title.toUpperCase()}</th>`;
    });
    htmlTable += "</tr><tr class='second-tr'>";
    headerStructure.forEach(({ children }) => {
        if (children) children.forEach(child => (htmlTable += `<th>${child}</th>`));
    });
    htmlTable += "</tr></thead><tbody>";
    data.forEach(row => {
        htmlTable += `<tr data-product-type="${row.productType || ''}">`;
        htmlTable += `<td>${row.prodCusip || row.identifier}</td>`;
        if (showSecondIsinColumn) {
            htmlTable += `<td>${row.prodIsin || ""}</td>`;
        }
        
        htmlTable += `<td>${row.underlyingAssetType || ""}</td>`;
        for (let i = 0; i < maxAssets; i++) {
            htmlTable += `<td>${row.assets[i] || ""}</td>`;
        }
        htmlTable += `<td>${row.productType || ""}</td>`;
        htmlTable += `<td>${row.productClient || ""}</td>`;
        htmlTable += `<td>${row.productTenor || ""}</td>`;
        htmlTable += `<td>${row.couponFrequency || ""}</td>`;
        htmlTable += `<td>${row.couponBarrierLevel || ""}</td>`;
        htmlTable += `<td>${row.couponMemory || ""}</td>`;

        htmlTable += `<td>${row.callFrequency || ""}</td>`;
        htmlTable += `<td>${row.callNonCallPeriod || ""}</td>`;

        htmlTable += `<td>${row.upsideCap || ""}</td>`;
        htmlTable += `<td>${row.upsideLeverage || ""}</td>`;
        if (hasBrenRenProducts) {
            htmlTable += `<td>${row.detailCappedUncapped || ""}</td>`;
        }
        htmlTable += `<td>${row.detailBufferKIBarrier || ""}</td>`;
        htmlTable += `<td>${row.detailBufferBarrierLevel || ""}</td>`;
        htmlTable += `<td>${row.detailInterestBarrierTriggerValue || ""}</td>`;
        
        htmlTable += `<td>${row.dateBookingStrikeDate || ""}</td>`;
        htmlTable += `<td>${row.dateBookingPricingDate || ""}</td>`;
        htmlTable += `<td>${row.maturityDate || ""}</td>`;
        htmlTable += `<td>${row.valuationDate || ""}</td>`;
        htmlTable += `<td>${row.earlyStrike || ""}</td>`;
        htmlTable += `<td>${row.termSheet || ""}</td>`;
        htmlTable += `<td>${row.finalPS || ""}</td>`;
        htmlTable += `<td>${row.factSheet || ""}</td>`;
        htmlTable += "</tr>";
    });
    htmlTable += "</tbody></table>";
    dataTable.innerHTML = htmlTable;
}

function exportExcel() {
    const selectedType = productTypeFilter.value;
    let filteredData = consolidatedData;

    if (selectedType !== 'all') {
        filteredData = consolidatedData.filter(row => row.productType === selectedType);
    }

    if (filteredData.length === 0) {
        message.innerHTML = "No data to export for the selected filter.";
        return;
    }

    const groupedData = filteredData.reduce((acc, row) => {
        const key = row.productType || 'Uncategorized';
        if (!acc[key]) acc[key] = [];
        acc[key].push(row);
        return acc;
    }, {});

    const workbook = XLSX.utils.book_new();

    for (const productType in groupedData) {
        const sheetData = groupedData[productType];
        const sheetName = sanitizeSheetName(productType);
        
        const format = sheetData[0]?.format;
        const idColumnTitle = format === 'Barclays' ? 'ISIN' : 'CUSIP';
        const showSecondIsinColumn = format === 'pyrEvoDoc';

        const isBrenRenSheet = sheetData.some(row => row.productType === 'BREN' || row.productType === 'REN');
        
        let assetHeaders = Array.from({ length: maxAssetsForExport }, (_, i) => `Asset ${i + 1}`);
        if (maxAssetsForExport === 1) assetHeaders = ["Asset"];
        
        let initialHeaders = [{ title: idColumnTitle }];
        if (showSecondIsinColumn) {
            initialHeaders.push({ title: "ISIN" });
        }

        const detailsChildren = ["Upside Cap", "Upside Leverage"];
        if (isBrenRenSheet) {
            detailsChildren.push("Capped / Uncapped");
        }
        detailsChildren.push("Buffer / Barrier", "Barrier/Buffer Level", "Interest v Barrier/Buffer");

        const callChildren = ["Frequency", "Non-call period"];

        const localHeaderStructure = [
            ...initialHeaders,
            { title: "Underlying", children: ["Asset Type", ...assetHeaders] },
            { title: "Product Details", children: ["Product Type", "Client", "Tenor"] },
            { title: "Coupons", children: ["Frequency", "Barrier Level", "Memory"] },
            { title: "CALL", children: callChildren },
            { title: "Details", children: detailsChildren },
            { title: "DATES IN BOOKINGS", children: ["Strike", "Pricing", "Maturity", "Valuation", "Early Strike"] },
            { title: "Doc Type", children: ["Term Sheet", "Final PS", "Fact Sheet"] }
        ];
        
        const headerRow1 = [];
        const headerRow2 = [];
        localHeaderStructure.forEach(h => {
            headerRow1.push(h.title.toUpperCase());
            if (h.children) {
                headerRow2.push(...h.children);
                for (let i = 1; i < h.children.length; i++) {
                    headerRow1.push(null);
                }
            } else {
                headerRow2.push(null);
            }
        });

        const dataAoA = sheetData.map(row => {
            const rowAsArray = [];
            rowAsArray.push(row.prodCusip || row.identifier);
            if (showSecondIsinColumn) {
                rowAsArray.push(row.prodIsin || "");
            }
            
            rowAsArray.push(row.underlyingAssetType);
            for (let i = 0; i < maxAssetsForExport; i++) {
                rowAsArray.push(row.assets[i] || "");
            }
            rowAsArray.push(
                row.productType, row.productClient, row.productTenor, row.couponFrequency,
                row.couponBarrierLevel, row.couponMemory, row.callFrequency, row.callNonCallPeriod,
                row.upsideCap, row.upsideLeverage
            );
            
            if (isBrenRenSheet) {
                rowAsArray.push(row.detailCappedUncapped);
            }
            
            rowAsArray.push(
                row.detailBufferKIBarrier, row.detailBufferBarrierLevel,
                row.detailInterestBarrierTriggerValue, row.dateBookingStrikeDate,
                row.dateBookingPricingDate, row.maturityDate, row.valuationDate, row.earlyStrike,
                row.termSheet, row.finalPS, row.factSheet
            );
            return rowAsArray;
        });

        const worksheet = XLSX.utils.aoa_to_sheet([headerRow1, headerRow2, ...dataAoA]);

        const merges = [];
        let colIndex = 0;
        localHeaderStructure.forEach(header => {
            if (header.children) {
                if (header.children.length > 1) {
                    merges.push({ s: { r: 0, c: colIndex }, e: { r: 0, c: colIndex + header.children.length - 1 } });
                }
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
                alignment: { vertical: "center", horizontal: "left" },
                border: { top: { style: "thin" }, right: { style: "thin" }, bottom: { style: "thin" }, left: { style: "thin" } }
            };
            if (cell.r < 2) {
                worksheet[cellAddress].s.font.bold = true;
                worksheet[cellAddress].s.alignment.horizontal = "center";
                worksheet[cellAddress].s.fill = { fgColor: { rgb: "DDEBF7" } };
            }
            const value = worksheet[cellAddress].v;
            const width = value ? value.toString().length + 2 : 12;
            if (!colWidths[cell.c] || width > colWidths[cell.c].wch) {
                colWidths[cell.c] = { wch: width };
            }
        });
        worksheet['!cols'] = colWidths;

        XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
    }
    
    XLSX.writeFile(workbook, "ExtractedXML_Data_Global.xlsx");
}