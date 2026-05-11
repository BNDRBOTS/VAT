import type * as XLSX from "xlsx";

export type Severity = "info" | "warn" | "error";

export type WorkbookBinding = {
  fileName: string;
  workbook: XLSX.WorkBook;
  sheetNames: string[];
};

export type DetectionCandidate = {
  name: string;
  score: number;
  matchedKeywords: string[];
};

export type DetailColumnMapping = {
  idCol: number | null;
  descriptionCol: number | null;
  docCol: number | null;
  postingDateCol: number | null;
  reportDateCol: number | null;
  amountCol: number | null;
  currencyCol: number | null;
  treatmentCol: number | null;
  noteCol: number | null;
  referenceCol: number | null;
};

export type LedgerColumnMapping = {
  codeCol: number | null;
  amountCol: number | null;
  descriptionCol: number | null;
};

export type SheetSelection = {
  detailSheet: string;
  summarySheet: string;
  revenueSheet: string;
  purchaseSheet: string;
};

export type RouteDecision =
  | "REPORT_MONTH"
  | "DEFER_TO_NEXT_MONTH"
  | "EXCLUDE_ALREADY_REPORTED"
  | "EXCLUDE_NON_TAXABLE"
  | "EXCLUDE_TRANSFER_PRICING"
  | "MANUAL_REVIEW";

export type DetailRecord = {
  rowNumber: number;
  sheetName: string;
  recordId: string;
  description: string;
  documentNumber: string;
  postingDate: string;
  reportDate: string;
  amount: number;
  currency: string;
  treatment: string;
  note: string;
  reference: string;
  decision: RouteDecision;
  routeBucket: string;
  releaseEligible: boolean;
  exceptions: string[];
  rawRow: unknown[];
  trace: string[];
};

export type SummaryBucket = {
  category: string;
  amount: number;
  note: string;
  rowNumber: number;
  sheetName: string;
  trace: string[];
};

export type BucketCheck = {
  category: string;
  detailAmount: number;
  summaryAmount: number;
  delta: number;
  matched: boolean;
};

export type CodeCheck = {
  code: string;
  leftAmount: number | null;
  rightAmount: number | null;
  leftHits: number;
  rightHits: number;
  matched: boolean;
};

export type DriftSignal = {
  severity: Severity;
  code: string;
  message: string;
  context?: string;
};

export type SheetSelectionProfile = {
  detailCandidates: DetectionCandidate[];
  summaryCandidates: DetectionCandidate[];
  revenueCandidates: DetectionCandidate[];
  purchaseCandidates: DetectionCandidate[];
};

export type AnalysisResult = {
  detailRows: DetailRecord[];
  reportable: DetailRecord[];
  deferred: DetailRecord[];
  excludedReported: DetailRecord[];
  excludedNonTaxable: DetailRecord[];
  excludedTransferPricing: DetailRecord[];
  manualReview: DetailRecord[];
  summaryBuckets: SummaryBucket[];
  bucketChecks: BucketCheck[];
  codeChecks: CodeCheck[];
  detailTotal: number;
  summaryTotal: number | null;
  totalMatched: boolean;
  hardStop: boolean;
  releaseReady: boolean;
  driftSignals: DriftSignal[];
};

export type PersistedSession = {
  reportMonth: string;
  reportYear: string;
  controlCodesText: string;
  sheetSelection: SheetSelection;
  detailHeaderRow: number;
  revenueHeaderRow: number;
  purchaseHeaderRow: number;
  detailMapping: DetailColumnMapping;
  revenueMapping: LedgerColumnMapping;
  purchaseMapping: LedgerColumnMapping;
};