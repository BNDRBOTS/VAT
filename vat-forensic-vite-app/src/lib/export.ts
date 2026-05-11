import * as XLSX from "xlsx";
import type { AnalysisResult, DetailRecord, SummaryBucket } from "../types";

function buildRecords(records: DetailRecord[]): unknown[][] {
  return [[
    "Record ID","Sheet","Row","Description","Document Number","Posting Date","Report Date","Amount","Currency","Treatment","Note","Reference","Decision","Bucket","Exceptions","Trace"
  ], ...records.map((record) => [
    record.recordId, record.sheetName, record.rowNumber, record.description, record.documentNumber, record.postingDate, record.reportDate,
    record.amount, record.currency, record.treatment, record.note, record.reference, record.decision, record.routeBucket,
    record.exceptions.join(" | "), record.trace.join(" | ")
  ])];
}

function buildSummary(summaryBuckets: SummaryBucket[]): unknown[][] {
  return [["Category","Amount","Note","Sheet","Row","Trace"], ...summaryBuckets.map((bucket) => [
    bucket.category, bucket.amount, bucket.note, bucket.sheetName, bucket.rowNumber, bucket.trace.join(" | ")
  ])];
}

function buildBucketChecks(analysis: AnalysisResult): unknown[][] {
  return [
    ["Category","Detail Amount","Summary Amount","Delta","Matched"],
    ...analysis.bucketChecks.map((check) => [check.category, check.detailAmount, check.summaryAmount, check.delta, check.matched ? "Yes" : "No"]),
    [],
    ["Detail Total", analysis.detailTotal],
    ["Summary Total", analysis.summaryTotal ?? ""],
    ["Total Match", analysis.totalMatched ? "Yes" : "No"]
  ];
}

function buildCodeChecks(analysis: AnalysisResult): unknown[][] {
  return [["Code","Revenue Amount","Purchase Amount","Revenue Hits","Purchase Hits","Matched"], ...analysis.codeChecks.map((check) => [
    check.code, check.leftAmount ?? "", check.rightAmount ?? "", check.leftHits, check.rightHits, check.matched ? "Yes" : "No"
  ])];
}

function buildDriftSignals(analysis: AnalysisResult): unknown[][] {
  return [["Severity","Code","Message","Context"], ...analysis.driftSignals.map((signal) => [
    signal.severity, signal.code, signal.message, signal.context ?? ""
  ])];
}

export function exportAuditWorkbook(analysis: AnalysisResult, reportMonth: string, reportYear: string): void {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(buildRecords(analysis.reportable)), "Reportable");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(buildRecords(analysis.deferred)), "Deferred");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(buildRecords(analysis.excludedReported)), "Excluded Reported");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(buildRecords(analysis.excludedNonTaxable)), "Excluded NonTaxable");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(buildRecords(analysis.excludedTransferPricing)), "Excluded TransferPricing");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(buildRecords(analysis.manualReview)), "Manual Review");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(buildSummary(analysis.summaryBuckets)), "Detected Summary");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(buildBucketChecks(analysis)), "Reconciliation");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(buildCodeChecks(analysis)), "Code Controls");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(buildDriftSignals(analysis)), "Drift Signals");
  XLSX.writeFile(workbook, `vat_audit_${reportMonth.toLowerCase()}_${reportYear}.xlsx`);
}

export function exportReleaseWorkbook(analysis: AnalysisResult, reportMonth: string, reportYear: string): void {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(buildRecords(analysis.reportable)), `Release ${reportMonth}`);
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(buildBucketChecks(analysis)), "Reconciliation");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(buildCodeChecks(analysis)), "Code Controls");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(buildDriftSignals(analysis)), "Drift Signals");
  XLSX.writeFile(workbook, `vat_release_${reportMonth.toLowerCase()}_${reportYear}.xlsx`);
}