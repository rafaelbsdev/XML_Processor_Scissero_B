const setDate = (strDate) => {
    if (!strDate) return "";
    const dateStr = strDate.includes('T') ? strDate : strDate + 'T00:00:00Z';
    const date = new Date(dateStr);
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
    const startStr = startDateStr.includes('T') ? startDateStr : startDateStr + 'T00:00:00Z';
    const endStr = endDateStr.includes('T') ? endDateStr : endDateStr + 'T00:00:00Z';
    const startDate = new Date(startStr);
    const endDate = new Date(endStr);
    if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) return "";
    let months = (endDate.getUTCFullYear() - startDate.getUTCFullYear()) * 12;
    months -= startDate.getUTCMonth();
    months += endDate.getUTCMonth();
    months = months < 0 ? 0 : months;
    if (months === 0 && endDate.getUTCDate() >= startDate.getUTCDate()) return '0M';
    if (months === 0) return "";

    return months >= 12 && months % 12 === 0 ? `${months / 12}Y` : `${months}M`;
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
    const reverseConvertible = product?.querySelector("reverseConvertible");

     if (!product || !tradableForm || !asset) {
         console.warn("Skipping pyrEvoDoc file due to missing essential nodes (product, tradableForm, or asset).");
         return null;
     }

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
        return findFirstContent(productNode, ["reverseConvertible > description", "productName", "bufferedReturnEnhancedNote > productType"]);
     };
    const upsideLeverage = (productNode) => { return findFirstContent(productNode, ["bufferedReturnEnhancedNote > upsideLeverage"]) || "N/A"; };
    const upsideCap = (productNode) => { const cap = findFirstContent(productNode, ["bufferedReturnEnhancedNote > upsideCap"]); return (cap && `${cap}%`) || "N/A"; };
    const getCoupon = (productNode, type) => {
        const couponSchedule = productNode?.querySelector("reverseConvertible > couponSchedule");
        if (!couponSchedule) return "N/A";
        switch (type) {
            case "frequency": return findFirstContent(couponSchedule, ["frequency"]) || "N/A";
            case "level": const level = findFirstContent(couponSchedule, ["coupons > contingentLevelLegal > level", "contigentLevelLegal > level"]); return (level && `${level}%`) || "N/A";
            case "memory": const hasMemory = findFirstContent(couponSchedule, ["hasMemory"]); return hasMemory ? (hasMemory === "false" ? "N" : "Y") : "N/A";
            case "annualisedRate": const rate = findFirstContent(couponSchedule, ["annualizedRate > level"]); return (rate && `${parseFloat(rate)}%`) || "N/A";
            default: return "N/A";
        }
     };
    const getEarlyStrike = (xmlNode) => { const strikeRaw = findFirstContent(xmlNode, ["securitized > issuance > prospectusStartDate", "strikeDate > date"]); const pricingRaw = findFirstContent(xmlNode, ["securitized > issuance > clientOrderTradeDate"]); return strikeRaw && pricingRaw && strikeRaw !== pricingRaw ? "Y" : "N"; };
    const detectClient = (xmlNode) => { const cp = findFirstContent(xmlNode, ["counterparty > name"]); const dealer = findFirstContent(xmlNode, ["dealer > name"]); const blob = (cp + " " + dealer).toLowerCase(); if (blob.includes("jpm")) return "JPM PB"; if (blob.includes("goldman")) return "GS"; if (blob.includes("bauble") || blob.includes("ubs")) return "UBS"; return "3P"; };
    const detectDocType = (xmlNode) => { const docType = findFirstContent(xmlNode, ["documentType"]).toUpperCase(); return { termSheet: docType.includes("TERMSHEET") ? "Y" : "N", finalPS: docType.includes("PRICING_SUPPLEMENT") ? "Y" : "N", factSheet: docType.includes("FACT_SHEET") ? "Y" : "N" }; };
    const getDetails = (productNode, type, tradableFormNode) => {
        const productType = getProducts(productNode);
        if (productType === "BREN" || productType === "REN") {
            switch (type) {
                case "strikelevel": return "Buffer";
                case "bufferlevel": const buffer = findFirstContent(productNode, ["bufferedReturnEnhancedNote > buffer"]); return buffer ? `${100 - parseFloat(buffer)}%` : "";
                case "frequency": return "At Maturity"; case "noncall": return "N/A"; case "callMonitoring": return "N/A";
            }
        }
        const rcNode = productNode?.querySelector("reverseConvertible"); if (!rcNode) { return type === "callMonitoring" ? "N/A" : ""; }
        const isBufferType = parseFloat(findFirstContent(rcNode, ["strike > level"])) < 100; const phoenixTypeNode = rcNode.querySelector("issuerCallable") || rcNode.querySelector("autocallSchedule"); const phoenixType = phoenixTypeNode ? (phoenixTypeNode.tagName === "issuerCallable" ? "issuerCallable" : "autocallSchedule") : null;
        switch (type) {
            case "strikelevel": return isBufferType ? "Buffer" : "KI Barrier";
            case "bufferlevel": if (isBufferType) { const bfl = findFirstContent(rcNode, ["buffer > level"]); return bfl ? `${parseFloat(bfl)}%` : ""; } else { const kib = findFirstContent(rcNode, ["knockInBarrier > barrierSchedule > barrierLevel > level"]); return kib ? `${parseFloat(kib)}%` : ""; }
            case "frequency": return phoenixType ? findFirstContent(phoenixTypeNode, ["barrierSchedule > frequency"]) : "N/A";
            case "noncall": if (!phoenixType) return "N/A"; const isd = findFirstContent(tradableFormNode, ["securitized > issuance > issueDate"]); const fcd = findFirstContent(phoenixTypeNode, ["barrierSchedule > firstDate"]); return calculateTenor(isd, fcd);
            case "callMonitoring": return phoenixTypeNode ? findFirstContent(phoenixTypeNode, ["monitoringType"]) : "N/A";
            default: const ibl = parseFloat(findFirstContent(rcNode, ["couponSchedule > contigentLevelLegal > level"])); let cl; let clbl; if (isBufferType) { cl = parseFloat(findFirstContent(rcNode, ["buffer > level"])); clbl = "Buffer"; } else { cl = parseFloat(findFirstContent(rcNode, ["knockInBarrier > barrierSchedule > barrierLevel > level"])); clbl = "KI Barrier"; } if (isNaN(ibl) || isNaN(cl)) return "N/A"; const cs = ibl > cl ? ">" : ibl < cl ? "<" : "="; return `Interest Barrier ${cs} ${clbl}`;
        }
     };
    const findIdentifier = (tradableFormNode, type) => {
        const identifiers = tradableFormNode.querySelectorAll('identifiers');
        for (const idNode of identifiers) { const typeNode = idNode.querySelector('type'); if (typeNode && typeNode.textContent.trim().toUpperCase() === type) { return idNode.querySelector('code')?.textContent.trim() || ""; } } return "";
     };

    const cusip = findIdentifier(tradableForm, 'CUSIP');
    if (!cusip) return null;
    if (!product?.querySelector("reverseConvertible") && !product?.querySelector("bufferedReturnEnhancedNote")) {
        console.warn(`Skipping CUSIP ${cusip} due to unsupported product structure within pyrEvoDoc.`);
        return null;
    }

    const { termSheet, finalPS, factSheet } = detectDocType(xmlNode);
    const cappedValue = findFirstContent(product, ['bufferedReturnEnhancedNote > capped', 'reverseConvertible > capped']);
    const product_type = getProducts(product);

    return {
        format: 'pyrEvoDoc', identifier: cusip, prodCusip: cusip, prodIsin: findIdentifier(tradableForm, 'ISIN'),
        underlyingAssetType: getUnderlying(asset), assets: Array.from(asset.querySelectorAll("assets")).map(node => findFirstContent(node, ["bloombergTickerSuffix"])),
        productType: product_type, productClient: detectClient(xmlNode), productTenor: getProducts(product, "tenor"),
        programmeName: "",
        jurisdiction: "",
        couponFrequency: getCoupon(product, "frequency"), couponBarrierLevel: getCoupon(product, "level"), couponMemory: getCoupon(product, "memory"), couponRateAnnualised: getCoupon(product, "annualisedRate"),
        upsideCap: upsideCap(product), upsideLeverage: upsideLeverage(product),
        callFrequency: getDetails(product, "frequency", tradableForm), callNonCallPeriod: getDetails(product, "noncall", tradableForm), callMonitoringType: getDetails(product, "callMonitoring", tradableForm),
        detailCappedUncapped: (product_type === 'BREN' || product_type === 'REN' || cappedValue) ? (cappedValue === 'true' ? 'Capped' : (cappedValue === 'false' ? 'Uncapped' : '')) : '',
        detailBufferKIBarrier: getDetails(product, "strikelevel", tradableForm), detailBufferBarrierLevel: getDetails(product, "bufferlevel", tradableForm), detailInterestBarrierTriggerValue: getDetails(product, null, tradableForm),
        dateBookingStrikeDate: setDate(findFirstContent(xmlNode, ["securitized > issuance > prospectusStartDate", "strikeDate > date"])), dateBookingPricingDate: setDate(findFirstContent(tradableForm, ["securitized > issuance > clientOrderTradeDate"])),
        maturityDate: setDate(findFirstContent(product, ["maturity > date", "redemptionDate", "settlementDate"])), valuationDate: setDate(findFirstContent(product, ["finalObsDate > date", "finalObservation"])),
        earlyStrike: getEarlyStrike(xmlNode), termSheet, finalPS, factSheet,
    };
};

export const extractBarclaysData = (xmlNode) => {
    const isin = findFirstContent(xmlNode, ["codes > isin"]);
    if (!isin) return null;
    const dataNode = xmlNode.querySelector("data");
    if (!dataNode) { console.warn("Could not find <data> node in Barclays XML for ISIN:", isin); return null; }

    const issueDateStr = findFirstContent(dataNode, ["trade > dates > issueDate"]);
    const maturityDateStr = findFirstContent(dataNode, ["trade > dates > maturityDate"]);
    const assetsNodes = dataNode.querySelectorAll("referenceAsset > underlyings > item");
    const assetType = findFirstContent(assetsNodes[0], ["type"]);

    const barrierLevelValue = parseFloat(findFirstContent(dataNode, ["payoff > paymentAtMaturity > knockInBarrierLevelRelative > value"]));
    const formattedBarrierLevel = isNaN(barrierLevelValue) ? "N/A" : `${barrierLevelValue * 100}%`;
    const detailBufferKIBarrierValue = (formattedBarrierLevel === "N/A") ? "N/A" : "KI Barrier";
    const couponBarrierRaw = findFirstContent(dataNode, ["payoff > couponEvents > schedule > item > barrierLevelRelative > value"]);
    const couponBarrierFormatted = couponBarrierRaw ? `${parseFloat(couponBarrierRaw) * 100}%` : "N/A";

    const programmeNameSelector = "legalDocumentation > programName";
    const jurisdictionSelector = "governingLaw";

    return {
        format: 'Barclays', identifier: isin, prodCusip: "", prodIsin: isin,
        underlyingAssetType: assetsNodes.length > 1 ? `WorstOf ${assetType}` : `Single ${assetType}`,
        assets: Array.from(assetsNodes).map(node => findFirstContent(node, ["name"])),
        productType: findFirstContent(dataNode, ["product > clientProductType"]),
        productClient: findFirstContent(dataNode, ["manufacturer > nameShort"]),
        productTenor: calculateTenor(issueDateStr, maturityDateStr),
        programmeName: findFirstContent(dataNode, [programmeNameSelector]),
        jurisdiction: findFirstContent(dataNode, [jurisdictionSelector]).toUpperCase(),
        couponFrequency: findFirstContent(dataNode, ["payoff > couponEvents > couponObservationDatesInterval"]) || "N/A",
        couponBarrierLevel: couponBarrierFormatted,
        couponMemory: findFirstContent(dataNode, ["payoff > couponEvents > schedule > item > memory"]) === 'true' ? 'Y' : 'N',
        couponRateAnnualised: "N/A",
        upsideCap: "N/A", upsideLeverage: "N/A",
        callFrequency: findFirstContent(dataNode, ["payoff > callEvents > autoCallObservationDatesInterval"]) || "N/A",
        callNonCallPeriod: calculateTenor(issueDateStr, findFirstContent(dataNode, ["payoff > callEvents > schedule > item > settlementDate"])),
        callMonitoringType: findFirstContent(dataNode, ["payoff > callEvents > monitoringType", "payoff > callEvents > barrierEventObservationType"]) || "N/A",
        detailCappedUncapped: "N/A", detailBufferKIBarrier: detailBufferKIBarrierValue, detailBufferBarrierLevel: formattedBarrierLevel,
        detailInterestBarrierTriggerValue: "N/A",
        dateBookingStrikeDate: setDate(findFirstContent(dataNode, ["trade > dates > initialValuationDate"])),
        dateBookingPricingDate: setDate(findFirstContent(dataNode, ["trade > dates > tradeDate"])),
        maturityDate: setDate(maturityDateStr), valuationDate: setDate(findFirstContent(dataNode, ["trade > dates > finalValuationDate"])),
        earlyStrike: "N/A", termSheet: "N", finalPS: "N", factSheet: "N",
    };
};