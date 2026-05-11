import type {
  AnalysisResult,
  CodeCheck,
  DetailColumnMapping,
  DetailRecord,
  DriftSignal,
  LedgerColumnMapping,
  RouteDecision,
  SummaryBucket
} from "../types";
import {
  approxEqual,
  dominantMoneyFromRow,
  extractCurrency,
  getAOA,
  getCell,
  getCellMoney,
  nextMonthName,
  normalizeText,
  parseDateDisplay,
  parseMoney,
  rowHasUsefulCells
} from "./workbook";
import type * as XLSX from "xlsx";

const CONTROL_PHRASES = [
  "report in","do not report","reported in","paid in","confirmed as paid",
  "non-taxable","not subject to vat","vat reclass","import services","reverse charges",
  "target tax account","control total","bank fees","royalty fee","grand total","total"
];

const CATEGORY_ALIASES: Record<string, string[]> = {
  "Transf. pstg for target tax account": ["transf. pstg for target tax account","target tax account"],
  Other: ["other","confirmed as paid"],
  "Reported in June": ["reported in june","reported june","bank fees reported in june"],
  "Reported in July": ["reported in july","reported july","bank fees reported in july"],
  "Paid in September": ["paid in september"],
  "VAT Reclass": ["vat reclass","izettle bank fees"],
  "Non-taxable - reverse": ["non-taxable","not subject to vat","royalty fee"],
  "Import services": ["import services","reverse charges","225000-0614"],
  "Transfer Pricing": ["transfer pricing"]
};

function canonicalBucketName(value: string, year: string): string {
  const lower = normalizeText(value);
  if (lower.includes("target tax account")) return "Transf. pstg for target tax account";
  if (lower.includes("vat reclass") || lower.includes("izettle bank fees")) return "VAT Reclass";
  if (lower.includes("import services") || lower.includes("reverse charges") || lower.includes("225000-0614")) return "Import services";
  if (lower.includes("non-taxable") || lower.includes("not subject to vat") || lower.includes("royalty fee")) return "Non-taxable - reverse";
  if (lower.includes("transfer pricing")) return "Transfer Pricing";
  if (lower.includes("reported in june") || lower.includes("reported june")) return `Reported in June ${year}`;
  if (lower.includes("reported in july") || lower.includes("reported july")) return `Reported in July ${year}`;
  if (lower.includes("paid in september")) return `Paid in September ${year}`;
  if (lower.includes("other")) return "Other";
  return value;
}

function hasMonthReference(text: string, monthName: string, year: string): boolean {
  const lower = normalizeText(text);
  const month = normalizeText(monthName);
  const yearText = normalizeText(year);
  return lower.includes(`${month} ${yearText}`) || lower.includes(`${month}${yearText}`) || lower.includes(month);
}

function looksLikeControlRow(rowText: string): boolean {
  const lower = normalizeText(rowText);
  return CONTROL_PHRASES.some((phrase) => lower.includes(phrase));
}

function looksLikeTotalOnlyRow(rowText: string): boolean {
  return /\b(total|subtotal|grand total|control total)\b/i.test(rowText);
}

function inferRoute(
  textBlob: string,
  treatmentText: string,
  noteText: string,
  monthName: string,
  year: string
): { decision: RouteDecision; routeBucket: string; releaseEligible: boolean; trace: string[]; exceptions: string[] } {
  const lower = normalizeText(`${textBlob} ${treatmentText} ${noteText}`);
  const nextMonth = normalizeText(nextMonthName(monthName));
  const trace: string[] = [];
  const exceptions: string[] = [];

  if (!lower) {
    return { decision: "MANUAL_REVIEW", routeBucket: "Manual review", releaseEligible: false, trace: ["No routing text available"], exceptions: ["No classification text found"] };
  }
  if (lower.includes("transfer pricing")) {
    trace.push("Matched transfer pricing exclusion");
    return { decision: "EXCLUDE_TRANSFER_PRICING", routeBucket: "Transfer Pricing", releaseEligible: false, trace, exceptions };
  }
  if (lower.includes("non-taxable") || lower.includes("not subject to vat") || lower.includes("royalty fee")) {
    trace.push("Matched non-taxable exclusion");
    return { decision: "EXCLUDE_NON_TAXABLE", routeBucket: "Non-taxable - reverse", releaseEligible: false, trace, exceptions };
  }
  if (lower.includes("reported in june") || lower.includes("reported june") || lower.includes("bank fees reported in june")) {
    trace.push("Matched already reported June");
    return { decision: "EXCLUDE_ALREADY_REPORTED", routeBucket: `Reported in June ${year}`, releaseEligible: false, trace, exceptions };
  }
  if (lower.includes("reported in july") || lower.includes("reported july") || lower.includes("bank fees reported in july")) {
    trace.push("Matched already reported July");
    return { decision: "EXCLUDE_ALREADY_REPORTED", routeBucket: `Reported in July ${year}`, releaseEligible: false, trace, exceptions };
  }
  if (lower.includes(`paid in ${nextMonth}`)) {
    trace.push("Matched next-month deferral");
    return { decision: "DEFER_TO_NEXT_MONTH", routeBucket: `Paid in ${nextMonthName(monthName)} ${year}`, releaseEligible: false, trace, exceptions };
  }
  if (lower.includes("paid in september")) {
    trace.push("Matched September deferral");
    return { decision: "DEFER_TO_NEXT_MONTH", routeBucket: `Paid in September ${year}`, releaseEligible: false, trace, exceptions };
  }
  if (lower.includes("transf. pstg for target tax account") || lower.includes("target tax account")) {
    trace.push("Matched target tax account rule");
    return { decision: "REPORT_MONTH", routeBucket: "Transf. pstg for target tax account", releaseEligible: true, trace, exceptions };
  }
  if (lower.includes("vat reclass") || lower.includes("izettle bank fees")) {
    trace.push("Matched VAT reclass rule");
    return { decision: "REPORT_MONTH", routeBucket: "VAT Reclass", releaseEligible: true, trace, exceptions };
  }
  if (lower.includes("import services") || lower.includes("reverse charges") || lower.includes("225000-0614")) {
    trace.push("Matched import services rule");
    return { decision: "REPORT_MONTH", routeBucket: "Import services", releaseEligible: true, trace, exceptions };
  }
  if (lower.includes("other")) {
    if (lower.includes("confirm") || lower.includes("confirmed as paid") || lower.includes("paid") || lower.includes(`report in ${normalizeText(monthName)}`)) {
      trace.push("Matched confirmed other rule");
      return { decision: "REPORT_MONTH", routeBucket: "Other", releaseEligible: true, trace, exceptions };
    }
    trace.push("Other bucket found without explicit release language");
    return { decision: "MANUAL_REVIEW", routeBucket: "Other", releaseEligible: false, trace, exceptions: ["Other bucket exists without explicit confirmation to report"] };
  }
  if (hasMonthReference(lower, monthName, year) && lower.includes("report")) {
    trace.push("Matched explicit month reporting directive");
    return { decision: "REPORT_MONTH", routeBucket: "Reportable by directive", releaseEligible: true, trace, exceptions };
  }
  trace.push("No routing rule matched");
  return { decision: "MANUAL_REVIEW", routeBucket: "Manual review", releaseEligible: false, trace, exceptions: ["No routing rule matched from treatment/status text"] };
}

export function extractDetailRecords(
  workbook: XLSX.WorkBook,
  sheetName: string,
  headerRow: number,
  mapping: DetailColumnMapping,
  monthName: string,
  year: string,
  driftSignals: DriftSignal[]
): DetailRecord[] {
  const aoa = getAOA(workbook, sheetName);
  const records: DetailRecord[] = [];
  const duplicateCounter = new Map<string, number>();

  for (let i = headerRow + 1; i < aoa.length; i += 1) {
    const row = aoa[i] ?? [];
    if (!rowHasUsefulCells(row)) continue;

    const rowText = row.map((cell) => String(cell ?? "")).join(" | ");
    const mappedAmount = getCellMoney(row, mapping.amountCol);
    const fallbackAmount = mappedAmount === null ? dominantMoneyFromRow(row) : null;
    const amount = mappedAmount ?? fallbackAmount;

    const treatment = getCell(row, mapping.treatmentCol);
    const note = getCell(row, mapping.noteCol);
    const description = getCell(row, mapping.descriptionCol);
    const doc = getCell(row, mapping.docCol);
    const recordIdExplicit = getCell(row, mapping.idCol);
    const reference = getCell(row, mapping.referenceCol);
    const postingDate = parseDateDisplay(row[mapping.postingDateCol ?? -1]);
    const reportDate = parseDateDisplay(row[mapping.reportDateCol ?? -1]);
    const currency = extractCurrency(row[mapping.amountCol ?? -1], row[mapping.currencyCol ?? -1]);
    const likelyControl = looksLikeControlRow(rowText) && !doc && !reference && !recordIdExplicit;
    const likelyTotalOnly = looksLikeTotalOnlyRow(rowText) && !doc && !reference && !recordIdExplicit;

    if (likelyControl || likelyTotalOnly) continue;
    if (amount === null) {
      driftSignals.push({ severity: "warn", code: "ROW_AMOUNT_UNRESOLVED", message: `Skipped row ${i + 1} because no reliable amount could be determined.`, context: sheetName });
      continue;
    }

    const route = inferRoute(`${description} ${doc} ${reference}`, treatment || rowText, note, monthName, year);
    const recordId = recordIdExplicit || doc || reference || `${sheetName}-R${i + 1}`;
    duplicateCounter.set(recordId, (duplicateCounter.get(recordId) ?? 0) + 1);

    const exceptions = [...route.exceptions];
    const trace = [...route.trace];
    if (!currency) exceptions.push("Currency missing");
    if (!postingDate && !reportDate) exceptions.push("No posting/report date found");
    if (!treatment && !note && !description) exceptions.push("Treatment/status missing");
    if (fallbackAmount !== null && mappedAmount === null) {
      exceptions.push("Amount resolved from row fallback");
      trace.push("Used dominant-money fallback because mapped amount was missing");
    }

    const reportDecision = exceptions.length > route.exceptions.length && route.decision === "REPORT_MONTH" ? "MANUAL_REVIEW" : route.decision;
    records.push({
      rowNumber: i + 1,
      sheetName,
      recordId,
      description,
      documentNumber: doc,
      postingDate,
      reportDate,
      amount,
      currency,
      treatment,
      note,
      reference,
      decision: reportDecision,
      routeBucket: reportDecision === "MANUAL_REVIEW" ? "Manual review" : canonicalBucketName(route.routeBucket, year),
      releaseEligible: reportDecision === "REPORT_MONTH" && exceptions.length === route.exceptions.length && route.releaseEligible,
      exceptions,
      rawRow: row,
      trace
    });
  }

  duplicateCounter.forEach((count, recordId) => {
    if (count > 1) driftSignals.push({ severity: "warn", code: "DUPLICATE_RECORD_ID", message: `Record ID ${recordId} appears ${count} times in extracted detail.`, context: sheetName });
  });

  return records;
}

export function detectSummaryBuckets(
  workbook: XLSX.WorkBook,
  sheetName: string,
  year: string,
  driftSignals: DriftSignal[]
): SummaryBucket[] {
  const aoa = getAOA(workbook, sheetName);
  const buckets: SummaryBucket[] = [];

  for (let i = 0; i < aoa.length; i += 1) {
    const row = aoa[i] ?? [];
    const rowText = row.map((cell) => String(cell ?? "")).join(" | ");
    if (!looksLikeControlRow(rowText)) continue;

    for (let j = 0; j < row.length; j += 1) {
      const cellText = normalizeText(row[j]);
      if (!cellText) continue;
      const category = Object.entries(CATEGORY_ALIASES).find(([, aliases]) => aliases.some((alias) => cellText.includes(alias)))?.[0];
      if (!category) continue;

      let amount: number | null = null;
      let amountColumn = -1;
      for (let k = j + 1; k < Math.min(j + 8, row.length); k += 1) {
        const candidate = parseMoney(row[k]);
        if (candidate !== null) {
          amount = candidate;
          amountColumn = k;
          break;
        }
      }
      if (amount === null) amount = dominantMoneyFromRow(row);
      if (amount === null) {
        driftSignals.push({ severity: "warn", code: "SUMMARY_AMOUNT_MISSING", message: `Summary bucket ${category} on row ${i + 1} had no reliable nearby amount.`, context: sheetName });
        continue;
      }
      buckets.push({
        category: canonicalBucketName(category, year),
        amount,
        note: row.slice(Math.max(amountColumn + 1, j + 1)).map((cell) => String(cell ?? "").trim()).filter(Boolean).join(" | "),
        rowNumber: i + 1,
        sheetName,
        trace: [`Matched summary bucket ${category} on row ${i + 1}`]
      });
      break;
    }
  }

  const deduped = new Map<string, SummaryBucket>();
  for (const bucket of buckets) deduped.set(`${bucket.category}|${bucket.amount}|${bucket.rowNumber}`, bucket);
  const finalBuckets = Array.from(deduped.values());
  if (finalBuckets.length === 0) driftSignals.push({ severity: "error", code: "SUMMARY_BUCKETS_NOT_FOUND", message: "No summary buckets were detected in the selected summary sheet.", context: sheetName });
  return finalBuckets;
}

function summarizeByBucket(records: DetailRecord[]): Record<string, number> {
  return records.reduce<Record<string, number>>((acc, record) => {
    const key = record.routeBucket || "Unclassified";
    acc[key] = (acc[key] ?? 0) + record.amount;
    return acc;
  }, {});
}

export function compareBuckets(records: DetailRecord[], summaryBuckets: SummaryBucket[]) {
  const detailMap = summarizeByBucket(records);
  const summaryMap = summaryBuckets.reduce<Record<string, number>>((acc, bucket) => {
    acc[bucket.category] = (acc[bucket.category] ?? 0) + bucket.amount;
    return acc;
  }, {});
  return Array.from(new Set([...Object.keys(detailMap), ...Object.keys(summaryMap)])).sort((a,b)=>a.localeCompare(b)).map((category) => {
    const detailAmount = detailMap[category] ?? 0;
    const summaryAmount = summaryMap[category] ?? 0;
    return { category, detailAmount, summaryAmount, delta: detailAmount - summaryAmount, matched: approxEqual(detailAmount, summaryAmount) };
  });
}

function codeMatchesCell(code: string, cell: string): boolean {
  const normalizedCell = normalizeText(cell);
  const normalizedCode = normalizeText(code);
  if (!normalizedCell || !normalizedCode) return false;
  if (normalizedCell === normalizedCode || normalizedCell.startsWith(`${normalizedCode}-`) || normalizedCell.startsWith(`${normalizedCode} `)) return true;
  const boundaryRegex = new RegExp(`(^|[^0-9a-z])${normalizedCode}([^0-9a-z]|$)`, "i");
  return boundaryRegex.test(normalizedCell);
}

export function aggregateCodeSheet(
  workbook: XLSX.WorkBook,
  sheetName: string,
  headerRow: number,
  mapping: LedgerColumnMapping,
  codes: string[],
  driftSignals: DriftSignal[],
  sideLabel: string
): Map<string, { amount: number; hits: number }> {
  const aoa = getAOA(workbook, sheetName);
  const result = new Map<string, { amount: number; hits: number }>();
  codes.forEach((code) => result.set(code, { amount: 0, hits: 0 }));
  for (let i = headerRow + 1; i < aoa.length; i += 1) {
    const row = aoa[i] ?? [];
    if (!rowHasUsefulCells(row)) continue;
    const rowText = row.map((cell) => String(cell ?? "").trim());
    const codeValue = getCell(row, mapping.codeCol);
    const description = getCell(row, mapping.descriptionCol);
    const mappedAmount = getCellMoney(row, mapping.amountCol);
    const fallbackAmount = mappedAmount ?? dominantMoneyFromRow(row);

    codes.forEach((code) => {
      const matched = codeMatchesCell(code, codeValue) || codeMatchesCell(code, description) || rowText.some((cell) => codeMatchesCell(code, cell));
      if (matched) {
        const value = fallbackAmount ?? 0;
        const existing = result.get(code) ?? { amount: 0, hits: 0 };
        result.set(code, { amount: existing.amount + value, hits: existing.hits + 1 });
        if (mappedAmount === null && fallbackAmount !== null) {
          driftSignals.push({ severity: "info", code: "LEDGER_FALLBACK_AMOUNT", message: `${sideLabel} code ${code} on row ${i + 1} used fallback amount detection.`, context: sheetName });
        }
      }
    });
  }
  return result;
}

export function analyzeWorkbooks(params: {
  primaryWorkbook: XLSX.WorkBook;
  detailSheet: string;
  summarySheet: string;
  controlWorkbook: XLSX.WorkBook;
  revenueSheet: string;
  purchaseSheet: string;
  detailHeaderRow: number;
  revenueHeaderRow: number;
  purchaseHeaderRow: number;
  detailMapping: DetailColumnMapping;
  revenueMapping: LedgerColumnMapping;
  purchaseMapping: LedgerColumnMapping;
  reportMonth: string;
  reportYear: string;
  controlCodesText: string;
}): AnalysisResult {
  const driftSignals: DriftSignal[] = [];
  if (params.detailMapping.amountCol === null) driftSignals.push({ severity: "error", code: "AMOUNT_COLUMN_UNMAPPED", message: "Detail amount column is not mapped. Fallback extraction will be attempted row by row.", context: params.detailSheet });
  if (params.detailMapping.treatmentCol === null && params.detailMapping.noteCol === null) driftSignals.push({ severity: "warn", code: "ROUTING_TEXT_WEAK", message: "No dedicated treatment or note column is mapped. Row text fallback will be used.", context: params.detailSheet });
  if (params.revenueMapping.codeCol === null || params.purchaseMapping.codeCol === null) driftSignals.push({ severity: "warn", code: "CONTROL_CODE_MAPPING_WEAK", message: "One or both ledger code columns are unmapped. Code search will scan full rows.", context: `${params.revenueSheet} / ${params.purchaseSheet}` });

  const detailRows = extractDetailRecords(params.primaryWorkbook, params.detailSheet, params.detailHeaderRow, params.detailMapping, params.reportMonth, params.reportYear, driftSignals);
  const summaryBuckets = detectSummaryBuckets(params.primaryWorkbook, params.summarySheet || params.detailSheet, params.reportYear, driftSignals);

  const reportable = detailRows.filter((row) => row.decision === "REPORT_MONTH" && row.releaseEligible);
  const deferred = detailRows.filter((row) => row.decision === "DEFER_TO_NEXT_MONTH");
  const excludedReported = detailRows.filter((row) => row.decision === "EXCLUDE_ALREADY_REPORTED");
  const excludedNonTaxable = detailRows.filter((row) => row.decision === "EXCLUDE_NON_TAXABLE");
  const excludedTransferPricing = detailRows.filter((row) => row.decision === "EXCLUDE_TRANSFER_PRICING");
  const manualReview = detailRows.filter((row) => row.decision === "MANUAL_REVIEW" || row.exceptions.length > 0);

  const detailTotal = detailRows.reduce((sum, row) => sum + row.amount, 0);
  const summaryTotal = summaryBuckets.length ? summaryBuckets.reduce((sum, row) => sum + row.amount, 0) : null;
  const bucketChecks = compareBuckets(detailRows, summaryBuckets);
  const codes = params.controlCodesText.split(",").map((value) => value.trim()).filter(Boolean);

  const leftAgg = aggregateCodeSheet(params.controlWorkbook, params.revenueSheet, params.revenueHeaderRow, params.revenueMapping, codes, driftSignals, "Revenue");
  const rightAgg = aggregateCodeSheet(params.controlWorkbook, params.purchaseSheet, params.purchaseHeaderRow, params.purchaseMapping, codes, driftSignals, "Purchase");
  const codeChecks: CodeCheck[] = codes.map((code) => {
    const left = leftAgg.get(code) ?? { amount: 0, hits: 0 };
    const right = rightAgg.get(code) ?? { amount: 0, hits: 0 };
    return {
      code,
      leftAmount: left.hits > 0 ? left.amount : null,
      rightAmount: right.hits > 0 ? right.amount : null,
      leftHits: left.hits,
      rightHits: right.hits,
      matched: approxEqual(left.hits > 0 ? left.amount : null, right.hits > 0 ? right.amount : null)
    };
  });

  const totalMatched = summaryTotal !== null && approxEqual(detailTotal, summaryTotal);
  const anyBucketMismatch = bucketChecks.some((check) => !check.matched);
  const anyCodeMismatch = codeChecks.some((check) => !check.matched);
  if (!totalMatched) driftSignals.push({ severity: "error", code: "TOTAL_RECON_FAILED", message: "Detail total does not match detected summary total.", context: params.summarySheet });
  if (anyBucketMismatch) driftSignals.push({ severity: "error", code: "BUCKET_RECON_FAILED", message: "At least one category does not tie between detail and summary.", context: params.summarySheet });
  if (anyCodeMismatch) driftSignals.push({ severity: "error", code: "CONTROL_CODE_MISMATCH", message: "At least one control code amount does not tie between revenue and purchase sheets.", context: `${params.revenueSheet} / ${params.purchaseSheet}` });

  const criticalDrift = driftSignals.some((signal) => signal.severity === "error");
  const hardStop = !totalMatched || anyBucketMismatch || anyCodeMismatch || criticalDrift;
  const releaseReady = !hardStop && manualReview.length === 0 && reportable.length > 0;

  return {
    detailRows,
    reportable,
    deferred,
    excludedReported,
    excludedNonTaxable,
    excludedTransferPricing,
    manualReview,
    summaryBuckets,
    bucketChecks,
    codeChecks,
    detailTotal,
    summaryTotal,
    totalMatched,
    hardStop,
    releaseReady,
    driftSignals
  };
}