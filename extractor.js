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
    if (months === 0 && endDate.getUTCDate() < startDate.getUTCDate()) return "";

    return months >= 12 && months % 12 === 0 ? `${months / 12}Y` : `${months}M`;
};

export const detectXMLFormat = (xmlNode) => {
    if (xmlNode?.querySelector("pyrEvoDoc")) return 'pyrEvoDoc';
    if (xmlNode?.querySelector("priip")) return 'Barclays';
    return 'unknown';
};

export const extractPyrEvoDocData = (xmlNode) => {
    const productNode = xmlNode?.querySelector("product") ?? xmlNode?.querySelector("deals > deal > product");
    const tradableFormNode = xmlNode?.querySelector("tradableForm") ?? xmlNode?.querySelector("deals > deal > tradableForm");
    const assetNode = xmlNode?.querySelector("asset") ?? xmlNode?.querySelector("deals > deal > asset");

    if (!productNode || !tradableFormNode || !assetNode) {
        console.warn("Skipping pyrEvoDoc file: Missing essential nodes (product, tradableForm, or asset).");
        return null;
    }

    const findIdentifier = (node, type) => {
        const identifiers = node.querySelectorAll('identifiers');
        for (const idNode of identifiers) { const typeNode = idNode.querySelector('type'); if (typeNode && typeNode.textContent.trim().toUpperCase() === type) { return idNode.querySelector('code')?.textContent.trim() || ""; } } return "";
     };
    const cusip = findIdentifier(tradableFormNode, 'CUSIP');
    if (!cusip) { console.warn("Skipping pyrEvoDoc file: CUSIP not found."); return null; }

    const brenNode = productNode.querySelector("bufferedReturnEnhancedNote");
    const arnNode = productNode.querySelector("autocallableReviewNote");
    const pcdNode = productNode.querySelector("pointToPointCD");
    const rcNode = productNode.querySelector("reverseConvertible");
    const productSubtypeNode = brenNode || arnNode || pcdNode || rcNode;

    if (!productSubtypeNode) {
        console.warn(`Skipping CUSIP ${cusip}: Could not identify product subtype. Product Name: ${findFirstContent(productNode, ['productName'])}`);
        return null;
    }
    const subtype = productSubtypeNode.tagName;

    const getUnderlying = (node) => {
        const assets = node.querySelectorAll("assets"); if (assets.length === 0) return "N/A";
        const assetType = findFirstContent(assets[0], ["assetType"])?.replace("Exchange_Traded_Fund", "ETF") ?? "Unknown"; const basketType = findFirstContent(node, ["basketType"]) || "Multiple";
        return assets.length === 1 ? `Single ${assetType}` : `${basketType} ${assetType}`;
     };
    const getTenor = (node) => {
         const monthsStr = findFirstContent(node, ["tenor > months"]); if (!monthsStr) return "";
         const months = parseInt(monthsStr, 10); if (isNaN(months)) return "";
         return months >= 12 && months % 12 === 0 ? `${months / 12}Y` : `${months}M`;
     };
    const detectClient = (xmlNode) => {
         const tradeNode = xmlNode?.querySelector("trade") ?? xmlNode?.querySelector("deals > deal > trade"); const cp = findFirstContent(tradeNode, ["counterparty > name"]); const dealer = findFirstContent(tradableFormNode, ["securitized > issuance > dealer > name"]);
         const blob = (cp + " " + dealer).toLowerCase(); if (blob.includes("jpm")) return "JPM PB"; if (blob.includes("goldman")) return "GS"; if (blob.includes("bauble") || blob.includes("ubs")) return "UBS"; return "3P";
    };
    const detectDocType = (xmlNode) => {
         const docType = findFirstContent(xmlNode, ["header > documentType"])?.toUpperCase() || ""; return { termSheet: docType.includes("TERMSHEET"), finalPS: docType.includes("PRICING_SUPPLEMENT"), factSheet: docType.includes("FACT_SHEET") };
    };
    const getEarlyStrike = (tradableFormNode) => {
         const strikeRaw = findFirstContent(tradableFormNode, ["securitized > issuance > prospectusStartDate"]); const pricingRaw = findFirstContent(tradableFormNode, ["securitized > issuance > clientOrderTradeDate"]);
         return strikeRaw && pricingRaw && strikeRaw !== pricingRaw;
    };

    let productType = "";
    let couponFrequency = "N/A", couponBarrierLevel = "N/A", couponMemory = "N/A", couponRateAnnualised = "N/A";
    let upsideCap = "N/A", upsideLeverage = "N/A", detailCappedUncapped = "";
    let callFrequency = "N/A", callNonCallPeriod = "N/A", callMonitoringType = "N/A";
    let detailBufferKIBarrier = "N/A", detailBufferBarrierLevel = "N/A", detailInterestBarrierTriggerValue = "N/A";

    productType = findFirstContent(productSubtypeNode, ['productType']) || findFirstContent(rcNode, ['description']) || findFirstContent(productNode, ['productName']);

    switch (subtype) {
        case "bufferedReturnEnhancedNote":
            upsideLeverage = findFirstContent(productSubtypeNode, ["upsideLeverage"]);
            upsideLeverage = (upsideLeverage && !isNaN(parseFloat(upsideLeverage))) ? parseFloat(upsideLeverage).toFixed(2) : "N/A";

            upsideCap = findFirstContent(productSubtypeNode, ["upsideCap"]);
            upsideCap = (upsideCap && !isNaN(parseFloat(upsideCap))) ? `${parseFloat(upsideCap).toFixed(2)}%` : "N/A";

            // CAPPED/UNCAPPED COM DEFAULT
            const cappedBrenNode = productSubtypeNode.querySelector("capped");
            const isCappedBren = cappedBrenNode ? cappedBrenNode.textContent.trim() : null;
            if (isCappedBren === 'true') { detailCappedUncapped = 'Capped'; }
            else if (isCappedBren === 'false') { detailCappedUncapped = 'Uncapped'; }
            else { detailCappedUncapped = 'Uncapped'; } // Default

            callFrequency = "At Maturity";
            detailBufferKIBarrier = "Buffer";
            const bufferBren = findFirstContent(productSubtypeNode, ["buffer"]);
            detailBufferBarrierLevel = bufferBren && !isNaN(parseFloat(bufferBren)) ? `${(100 - parseFloat(bufferBren)).toFixed(2)}%` : "";
            break;

        case "autocallableReviewNote":
            upsideLeverage = findFirstContent(productSubtypeNode, ["upsideLeverage"]) || findFirstContent(productSubtypeNode, ["participation"]);
            upsideLeverage = (upsideLeverage && !isNaN(parseFloat(upsideLeverage))) ? parseFloat(upsideLeverage).toFixed(2) : "N/A";

            upsideCap = findFirstContent(productSubtypeNode, ["upsideCap"]);
            upsideCap = (upsideCap && !isNaN(parseFloat(upsideCap))) ? `${parseFloat(upsideCap).toFixed(2)}%` : "N/A";

            // CAPPED/UNCAPPED COM DEFAULT
            const cappedArnNode = productSubtypeNode.querySelector("capped");
            const isCappedArn = cappedArnNode ? cappedArnNode.textContent.trim() : null;
            if (isCappedArn === 'true') { detailCappedUncapped = 'Capped'; }
            else if (isCappedArn === 'false') { detailCappedUncapped = 'Uncapped'; }
            else { detailCappedUncapped = 'Uncapped'; } // Default

            // Call (KO Schedule)
            const koScheduleNode = productSubtypeNode.querySelector("koSchedule");
            if (koScheduleNode) {
                callFrequency = findFirstContent(koScheduleNode, ["frequency"]) || "N/A";
                callMonitoringType = findFirstContent(koScheduleNode, ["monitoringType"]) || "N/A";
                const issueDateArn = findFirstContent(tradableFormNode, ["securitized > issuance > issueDate"]);
                const firstCallDateArn = findFirstContent(koScheduleNode, ["barrierObservations:first-of-type > rebatePaymentDate", "firstDate"]);
                callNonCallPeriod = calculateTenor(issueDateArn, firstCallDateArn);
            }
             // Proteção (KI)
             const kiNodeArn = productSubtypeNode.querySelector("ki");
             if (kiNodeArn) {
                 detailBufferKIBarrier = "KI Barrier";
                 const kiLevelArn = findFirstContent(kiNodeArn, ["barrierSchedule > barrierLevel > level"]);
                 detailBufferBarrierLevel = (kiLevelArn && !isNaN(parseFloat(kiLevelArn))) ? `${parseFloat(kiLevelArn).toFixed(2)}%` : "";
             } else {
                 const ppArn = findFirstContent(tradableFormNode, ["securitized > issuance > principalProtection"]);
                 if (ppArn === "100") { detailBufferKIBarrier = "Principal Protected"; detailBufferBarrierLevel = "N/A"; }
             }
            break;

        case "pointToPointCD":
            upsideLeverage = findFirstContent(productSubtypeNode, ["participation"]);
            upsideLeverage = (upsideLeverage && !isNaN(parseFloat(upsideLeverage))) ? `${parseFloat(upsideLeverage)}%` : "N/A";

            upsideCap = findFirstContent(productSubtypeNode, ["upsideCap"]);
            upsideCap = (upsideCap && !isNaN(parseFloat(upsideCap))) ? `${parseFloat(upsideCap).toFixed(2)}%` : "N/A";

            // CAPPED/UNCAPPED COM DEFAULT
            const cappedPcdNode = productSubtypeNode.querySelector("capped");
            const isCappedPcd = cappedPcdNode ? cappedPcdNode.textContent.trim() : null;
             if (isCappedPcd === 'true') { detailCappedUncapped = 'Capped'; }
             else if (isCappedPcd === 'false') { detailCappedUncapped = 'Uncapped'; }
             else { detailCappedUncapped = 'Uncapped'; }

            // Proteção
            const ppPcd = findFirstContent(tradableFormNode, ["securitized > issuance > principalProtection"]);
            if (ppPcd === "100") { detailBufferKIBarrier = "Principal Protected"; detailBufferBarrierLevel = "N/A"; }
            break;

        case "reverseConvertible":
            const couponScheduleRc = productSubtypeNode.querySelector("couponSchedule");
            if (couponScheduleRc) {
                couponFrequency = findFirstContent(couponScheduleRc, ["frequency"]) || "N/A";
                const cpnLevelRc = findFirstContent(couponScheduleRc, ["coupons > contingentLevelLegal > level", "contigentLevelLegal > level"]);
                couponBarrierLevel = (cpnLevelRc && !isNaN(parseFloat(cpnLevelRc))) ? `${parseFloat(cpnLevelRc).toFixed(2)}%` : "N/A";
                const hasMemoryRc = findFirstContent(couponScheduleRc, ["hasMemory"]);
                couponMemory = hasMemoryRc ? (hasMemoryRc === "false" ? "N" : "Y") : "N/A";
                const rateRc = findFirstContent(couponScheduleRc, ["annualizedRate > level"]);
                couponRateAnnualised = (rateRc && !isNaN(parseFloat(rateRc))) ? `${parseFloat(rateRc).toFixed(2)}%` : "N/A";
            }
            // Call
            const phoenixTypeNodeRc = productSubtypeNode.querySelector("issuerCallable") || productSubtypeNode.querySelector("autocallSchedule");
            if (phoenixTypeNodeRc) {
                callFrequency = findFirstContent(phoenixTypeNodeRc, ["barrierSchedule > frequency"]) || "N/A";
                callMonitoringType = findFirstContent(phoenixTypeNodeRc, ["monitoringType"]) || "N/A";
                const issueDateRc = findFirstContent(tradableFormNode, ["securitized > issuance > issueDate"]);
                const firstCallDateRc = findFirstContent(phoenixTypeNodeRc, ["barrierSchedule > firstDate"]);
                callNonCallPeriod = calculateTenor(issueDateRc, firstCallDateRc);
            }
             // Proteção (Buffer ou KI)
             const strikeLevelRcStr = findFirstContent(productSubtypeNode, ["strike > level"]);
             const isBufferTypeRc = strikeLevelRcStr && parseFloat(strikeLevelRcStr) < 100;
             if (isBufferTypeRc) {
                 detailBufferKIBarrier = "Buffer";
                 const bufferLevelRc = findFirstContent(productSubtypeNode, ["buffer > level"]);
                 detailBufferBarrierLevel = bufferLevelRc && !isNaN(parseFloat(bufferLevelRc)) ? `${parseFloat(bufferLevelRc).toFixed(2)}%` : "";
             } else {
                 detailBufferKIBarrier = "KI Barrier";
                 const kiBarrierRc = findFirstContent(productSubtypeNode, ["knockInBarrier > barrierSchedule > barrierLevel > level"]);
                 detailBufferBarrierLevel = kiBarrierRc && !isNaN(parseFloat(kiBarrierRc)) ? `${parseFloat(kiBarrierRc).toFixed(2)}%` : "";
             }
             // Interest vs Barrier/Buffer
             if (couponScheduleRc) {
                 const interestBarrierLevelRc = parseFloat(findFirstContent(couponScheduleRc, ["contigentLevelLegal > level"]));
                 let comparisonLevelRc; let comparisonLabelRc;
                 if (isBufferTypeRc) { comparisonLevelRc = parseFloat(findFirstContent(productSubtypeNode, ["buffer > level"])); comparisonLabelRc = "Buffer"; }
                 else { comparisonLevelRc = parseFloat(findFirstContent(productSubtypeNode, ["knockInBarrier > barrierSchedule > barrierLevel > level"])); comparisonLabelRc = "KI Barrier"; }
                 if (!isNaN(interestBarrierLevelRc) && !isNaN(comparisonLevelRc)) {
                     const comparisonSymbolRc = interestBarrierLevelRc > comparisonLevelRc ? ">" : interestBarrierLevelRc < comparisonLevelRc ? "<" : "=";
                     detailInterestBarrierTriggerValue = `Interest Barrier ${comparisonSymbolRc} ${comparisonLabelRc}`;
                 } else { detailInterestBarrierTriggerValue = "N/A"; }
             }

            const cappedRcNode = productSubtypeNode.querySelector("capped");
            const isCappedRc = cappedRcNode ? cappedRcNode.textContent.trim() : null;
             if (isCappedRc === 'true') { detailCappedUncapped = 'Capped'; }
             else if (isCappedRc === 'false') { detailCappedUncapped = 'Uncapped'; }
             else { detailCappedUncapped = 'Uncapped'; }
            break;
    }

    const { termSheet, finalPS, factSheet } = detectDocType(xmlNode);
    const booleanTermSheet = termSheet ? 'Y' : 'N';
    const booleanFinalPS = finalPS ? 'Y' : 'N';
    const booleanFactSheet = factSheet ? 'Y' : 'N';
    const booleanEarlyStrike = getEarlyStrike(tradableFormNode) ? 'Y' : 'N';

    const strikeDate = setDate(findFirstContent(productSubtypeNode, ["strikeDate > date"]) || findFirstContent(tradableFormNode, ["securitized > issuance > prospectusStartDate"]) || findFirstContent(productNode, ["strikeDate > date"]));
    const maturityDate = setDate(findFirstContent(productSubtypeNode, ["maturity > date"]) || findFirstContent(rcNode, ["redemption > redemptionDate"]) || findFirstContent(productNode, ["maturity > date"]));
    const valuationDate = setDate(findFirstContent(productSubtypeNode, ["finalObsDate > date"]) || findFirstContent(rcNode, ["redemption > finalObservation"]) || findFirstContent(productNode, ["finalObsDate > date"]));
    const pricingDate = setDate(findFirstContent(tradableFormNode, ["securitized > issuance > clientOrderTradeDate"]));

    return {
        format: 'pyrEvoDoc', identifier: cusip, prodCusip: cusip, prodIsin: findIdentifier(tradableFormNode, 'ISIN'),
        underlyingAssetType: getUnderlying(assetNode), assets: Array.from(assetNode.querySelectorAll("assets")).map(node => findFirstContent(node, ["bloombergTickerSuffix"])),
        productType: productType, productClient: detectClient(xmlNode), productTenor: getTenor(productNode),
        programmeName: "", jurisdiction: "",
        couponFrequency: couponFrequency, couponBarrierLevel: couponBarrierLevel, couponMemory: couponMemory, couponRateAnnualised: couponRateAnnualised,
        upsideCap: upsideCap, upsideLeverage: upsideLeverage,
        callFrequency: callFrequency, callNonCallPeriod: callNonCallPeriod, callMonitoringType: callMonitoringType,
        detailCappedUncapped: detailCappedUncapped,
        detailBufferKIBarrier: detailBufferKIBarrier, detailBufferBarrierLevel: detailBufferBarrierLevel, detailInterestBarrierTriggerValue: detailInterestBarrierTriggerValue,
        dateBookingStrikeDate: strikeDate, dateBookingPricingDate: pricingDate,
        maturityDate: maturityDate, valuationDate: valuationDate,
        earlyStrike: booleanEarlyStrike,
        termSheet: booleanTermSheet, finalPS: booleanFinalPS, factSheet: booleanFactSheet,
    };
};

export const extractBarclaysData = (xmlNode) => {
    let isin = findFirstContent(xmlNode, ["priip > data > codes > isin"]);
    if (!isin) {
        isin = findFirstContent(xmlNode, ["priip > barclays > identifiers > identifier[type='ISIN']"]);
        if (!isin) return null;
        console.warn("Using fallback ISIN for Barclays XML:", isin);
    }
    const dataNode = xmlNode?.querySelector("priip > data");
    if (!dataNode) { console.warn("Could not find <data> node in Barclays XML for ISIN:", isin); return null; }

    const issueDateStr = findFirstContent(dataNode, ["trade > dates > issueDate"]);
    const maturityDateStr = findFirstContent(dataNode, ["trade > dates > maturityDate"]);
    const assetsNodes = dataNode.querySelectorAll("referenceAsset > underlyings > item");
    const assetType = findFirstContent(assetsNodes[0], ["type"]);

    const barrierLevelValue = parseFloat(findFirstContent(dataNode, ["payoff > paymentAtMaturity > knockInBarrierLevelRelative > value"]));
    const formattedBarrierLevel = !isNaN(barrierLevelValue) ? `${(barrierLevelValue * 100).toFixed(2)}%` : "N/A";
    const detailBufferKIBarrierValue = (formattedBarrierLevel === "N/A") ? "N/A" : "KI Barrier";
    const couponBarrierRaw = findFirstContent(dataNode, ["payoff > couponEvents > schedule > item > barrierLevelRelative > value"]);
    const couponBarrierFormatted = couponBarrierRaw && !isNaN(parseFloat(couponBarrierRaw)) ? `${(parseFloat(couponBarrierRaw) * 100).toFixed(2)}%` : "N/A";

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
        jurisdiction: findFirstContent(dataNode, [jurisdictionSelector])?.toUpperCase() || "",
        couponFrequency: findFirstContent(dataNode, ["payoff > couponEvents > couponObservationDatesInterval"]) || "N/A",
        couponBarrierLevel: couponBarrierFormatted,
        couponMemory: findFirstContent(dataNode, ["payoff > couponEvents > schedule > item > memory"]) === 'true' ? 'Y' : 'N',
        couponRateAnnualised: "N/A",
        upsideCap: "N/A", upsideLeverage: "N/A",
        callFrequency: findFirstContent(dataNode, ["payoff > callEvents > autoCallObservationDatesInterval"]) || "N/A",
        callNonCallPeriod: calculateTenor(issueDateStr, findFirstContent(dataNode, ["payoff > callEvents > schedule > item > settlementDate"])),
        callMonitoringType: findFirstContent(dataNode, ["payoff > callEvents > monitoringType", "payoff > callEvents > barrierEventObservationType"]) || "N/A",
        detailCappedUncapped: "N/A",
        detailBufferKIBarrier: detailBufferKIBarrierValue, detailBufferBarrierLevel: formattedBarrierLevel,
        detailInterestBarrierTriggerValue: "N/A",
        dateBookingStrikeDate: setDate(findFirstContent(dataNode, ["trade > dates > initialValuationDate"])),
        dateBookingPricingDate: setDate(findFirstContent(dataNode, ["trade > dates > tradeDate"])),
        maturityDate: setDate(maturityDateStr), valuationDate: setDate(findFirstContent(dataNode, ["trade > dates > finalValuationDate"])),
        earlyStrike: "N/A", termSheet: "N", finalPS: "N", factSheet: "N",
    };
};