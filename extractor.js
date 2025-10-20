const setDate = (strDate) => {
    if (!strDate) return "";
    const date = new Date(strDate + 'T00:00:00Z');
    if (isNaN(date.getTime())) return "";
    const day = date.getUTCDate().toString().padStart(2, '0');
    const month = date.toLocaleString('en-US', { month: 'short', timeZone: 'UTC' });
    const year = date.getUTCFullYear().toString().slice(-2);
    return `${day}-${month}-${year}`;
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

export const detectXMLFormat = (xmlNode) => {
    if (xmlNode.querySelector("pyrEvoDoc")) return 'pyrEvoDoc';
    if (xmlNode.querySelector("priip")) return 'Barclays';
    return 'unknown';
};

export const extractPyrEvoDocData = (xmlNode) => {
    const product = xmlNode.querySelector("product");
    const tradableForm = xmlNode.querySelector("tradableForm");
    const asset = xmlNode.querySelector("asset");
    const getUnderlying = (assetNode, key) => {
        const assets = assetNode.querySelectorAll("assets");
        if (typeof key === "number") return findFirstContent(assets[key], ["bloombergTickerSuffix"]);
        if (assets.length === 0) return "N/A";
        const assetType = findFirstContent(assets[0], ["assetType"]).replace("Exchange_Traded_Fund", "ETF");
        const basketType = findFirstContent(assetNode, ["basketType"]) || "Multiple";
        return assets.length === 1 ? `Single ${assetType}` : `${basketType} ${assetType}`;
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
        return findFirstContent(productNode, ["bufferedReturnEnhancedNote > upsideLeverage"]) || "N/A";
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
        return { termSheet: docType.includes("TERMSHEET") ? "Y" : "N", finalPS: docType.includes("PRICING_SUPPLEMENT") ? "Y" : "N", factSheet: docType.includes("FACT_SHEET") ? "Y" : "N" };
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
            case "strikelevel": return isBufferType ? "Buffer" : "KI Barrier";
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
                return idNode.querySelector('code')?.textContent.trim() || "";
            }
        }
        return "";
    };

    const cusip = findIdentifier(tradableForm, 'CUSIP');
    if (!cusip) return null;
    const { termSheet, finalPS, factSheet } = detectDocType(xmlNode);
    const cappedValue = findFirstContent(product, ['bufferedReturnEnhancedNote > capped', 'capped']);
    const product_type = getProducts(product);

    return {
        format: 'pyrEvoDoc', identifier: cusip, prodCusip: cusip, prodIsin: findIdentifier(tradableForm, 'ISIN'),
        underlyingAssetType: getUnderlying(asset), assets: Array.from(asset.querySelectorAll("assets")).map(node => findFirstContent(node, ["bloombergTickerSuffix"])),
        productType: product_type, productClient: detectClient(xmlNode), productTenor: getProducts(product, "tenor"), couponFrequency: getCoupon(product, "frequency"),
        couponBarrierLevel: getCoupon(product, "level"), couponMemory: getCoupon(product, "memory"), upsideCap: upsideCap(product), upsideLeverage: upsideLeverage(product),
        callFrequency: getDetails(product, "frequency", tradableForm), callNonCallPeriod: getDetails(product, "noncall", tradableForm),
        detailCappedUncapped: (product_type === 'BREN' || product_type === 'REN') ? (cappedValue === 'true' ? 'Capped' : (cappedValue === 'false' ? 'Uncapped' : '')) : '',
        detailBufferKIBarrier: getDetails(product, "strikelevel"), detailBufferBarrierLevel: getDetails(product, "bufferlevel", tradableForm),
        detailInterestBarrierTriggerValue: getDetails(product, null, tradableForm),
        dateBookingStrikeDate: setDate(findFirstContent(xmlNode, ["securitized > issuance > prospectusStartDate", "strikeDate > date"])),
        dateBookingPricingDate: setDate(findFirstContent(tradableForm, ["securitized > issuance > clientOrderTradeDate"])),
        maturityDate: setDate(findFirstContent(product, ["maturity > date", "redemptionDate", "settlementDate"])),
        valuationDate: setDate(findFirstContent(product, ["finalObsDate > date", "finalObservation"])),
        earlyStrike: getEarlyStrike(xmlNode), termSheet, finalPS, factSheet,
    };
};

export const extractBarclaysData = (xmlNode) => {
    const isin = findFirstContent(xmlNode, ["codes > isin"]);
    if (!isin) return null;

    const issueDateStr = findFirstContent(xmlNode, ["trade > dates > issueDate"]);
    const maturityDateStr = findFirstContent(xmlNode, ["trade > dates > maturityDate"]);
    const assetsNodes = xmlNode.querySelectorAll("referenceAsset > underlyings > item");
    const assetType = findFirstContent(assetsNodes[0], ["type"]);

    const barrierLevelValue = parseFloat(findFirstContent(xmlNode, ["payoff > paymentAtMaturity > knockInBarrierLevelRelative > value"]));
    const formattedBarrierLevel = isNaN(barrierLevelValue) ? "N/A" : `${barrierLevelValue * 100}%`;
    const detailBufferKIBarrierValue = (formattedBarrierLevel === "N/A") ? "N/A" : "KI Barrier";

    return {
        format: 'Barclays', identifier: isin, prodCusip: "", prodIsin: isin,
        underlyingAssetType: assetsNodes.length > 1 ? `WorstOf ${assetType}` : `Single ${assetType}`,
        assets: Array.from(assetsNodes).map(node => findFirstContent(node, ["name"])),
        productType: findFirstContent(xmlNode, ["product > clientProductType"]),
        productClient: findFirstContent(xmlNode, ["manufacturer > nameShort"]),
        productTenor: calculateTenor(issueDateStr, maturityDateStr),
        couponFrequency: findFirstContent(xmlNode, ["payoff > couponEvents > couponObservationDatesInterval"]) || "N/A",
        couponBarrierLevel: `${parseFloat(findFirstContent(xmlNode, ["payoff > couponEvents > schedule > item > barrierLevelRelative > value"])) * 100}%` || "N/A",
        couponMemory: findFirstContent(xmlNode, ["payoff > couponEvents > schedule > item > memory"]) === 'true' ? 'Y' : 'N',
        upsideCap: "N/A", upsideLeverage: "N/A", callFrequency: findFirstContent(xmlNode, ["payoff > callEvents > autoCallObservationDatesInterval"]) || "N/A",
        callNonCallPeriod: calculateTenor(issueDateStr, findFirstContent(xmlNode, ["payoff > callEvents > schedule > item > settlementDate"])),
        detailCappedUncapped: "N/A", detailBufferKIBarrier: detailBufferKIBarrierValue, detailBufferBarrierLevel: formattedBarrierLevel,
        detailInterestBarrierTriggerValue: "N/A",
        dateBookingStrikeDate: setDate(findFirstContent(xmlNode, ["trade > dates > initialValuationDate"])),
        dateBookingPricingDate: setDate(findFirstContent(xmlNode, ["trade > dates > tradeDate"])),
        maturityDate: setDate(maturityDateStr), valuationDate: setDate(findFirstContent(xmlNode, ["trade > dates > finalValuationDate"])),
        earlyStrike: "N/A", termSheet: "N", finalPS: "N", factSheet: "N",
    };
};