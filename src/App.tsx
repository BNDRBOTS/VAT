import { useEffect, useMemo, useState } from "react";
import type {
  AnalysisResult,
  DetailColumnMapping,
  LedgerColumnMapping,
  PersistedSession,
  SheetSelection,
  WorkbookBinding
} from "./types";
import { DataTable } from "./components/DataTable";
import { MetricCard } from "./components/MetricCard";
import {
  AlertTriangleIcon,
  CheckCircleIcon,
  FileIcon,
  FolderUpIcon,
  LockIcon,
  RefreshIcon,
  ShieldIcon,
  UploadIcon
} from "./components/Icons";
import { analyzeWorkbooks } from "./lib/forensics";
import { exportAuditWorkbook, exportReleaseWorkbook } from "./lib/export";
import { loadSession, saveSession } from "./lib/storage";
import {
  MONTHS,
  autoDetectDetailMapping,
  autoDetectLedgerMapping,
  buildHeaderLabels,
  defaultSheetSelection,
  detectHeaderRow,
  getAOA,
  profileSheetSelection,
  readWorkbookFromFile,
  toColumnLetters,
  toMonthLabel
} from "./lib/workbook";

const RULES = [
  "Transf. pstg for target tax account",
  "Other",
  "Reported in June",
  "Reported in July",
  "Paid in September",
  "VAT Reclass",
  "Non-taxable - reverse",
  "Import services",
  "Transfer Pricing"
];

const emptySheetSelection: SheetSelection = {
  detailSheet: "",
  summarySheet: "",
  revenueSheet: "",
  purchaseSheet: ""
};

const emptyDetailMapping: DetailColumnMapping = {
  idCol: null,
  descriptionCol: null,
  docCol: null,
  postingDateCol: null,
  reportDateCol: null,
  amountCol: null,
  currencyCol: null,
  treatmentCol: null,
  noteCol: null,
  referenceCol: null
};

const emptyLedgerMapping: LedgerColumnMapping = {
  codeCol: null,
  amountCol: null,
  descriptionCol: null
};

function formatAmount(value: number | null | undefined): string {
  if (value === null || value === undefined || Number.isNaN(value)) return "";
  return new Intl.NumberFormat(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(value);
}

function DecisionPill({ label }: { label: string }) {
  const normalized = label.toLowerCase();
  const className = normalized.includes("manual")
    ? "pill pill-bad"
    : normalized.includes("defer")
    ? "pill pill-warn"
    : normalized.includes("exclude")
    ? "pill pill-muted"
    : "pill pill-good";
  return <span className={className}>{label}</span>;
}

function SectionCard({ title, subtitle, children }: { title: string; subtitle?: string; children: React.ReactNode }) {
  return (
    <section className="section-card">
      <div className="section-head">
        <div>
          <h2>{title}</h2>
          {subtitle ? <p>{subtitle}</p> : null}
        </div>
      </div>
      {children}
    </section>
  );
}

function MappingSelect({
  label,
  value,
  onChange,
  options
}: {
  label: string;
  value: number | null;
  onChange: (next: number | null) => void;
  options: string[];
}) {
  return (
    <label className="field">
      <span>{label}</span>
      <select value={value === null ? "" : String(value)} onChange={(event) => onChange(event.target.value === "" ? null : Number(event.target.value))}>
        <option value="">Unmapped</option>
        {options.map((option, index) => (
          <option key={`${label}-${index}`} value={index}>
            {toColumnLetters(index)} · {option}
          </option>
        ))}
      </select>
    </label>
  );
}

export default function App() {
  const restored = loadSession();
  const [primaryBinding, setPrimaryBinding] = useState<WorkbookBinding | null>(null);
  const [secondaryBinding, setSecondaryBinding] = useState<WorkbookBinding | null>(null);
  const [reportMonth, setReportMonth] = useState<string>(restored?.reportMonth ?? "August");
  const [reportYear, setReportYear] = useState<string>(restored?.reportYear ?? "2025");
  const [controlCodesText, setControlCodesText] = useState<string>(restored?.controlCodesText ?? "462,463,742");
  const [sheetSelection, setSheetSelection] = useState<SheetSelection>(restored?.sheetSelection ?? emptySheetSelection);
  const [detailHeaderRow, setDetailHeaderRow] = useState<number>(restored?.detailHeaderRow ?? 0);
  const [revenueHeaderRow, setRevenueHeaderRow] = useState<number>(restored?.revenueHeaderRow ?? 0);
  const [purchaseHeaderRow, setPurchaseHeaderRow] = useState<number>(restored?.purchaseHeaderRow ?? 0);
  const [detailMapping, setDetailMapping] = useState<DetailColumnMapping>(restored?.detailMapping ?? emptyDetailMapping);
  const [revenueMapping, setRevenueMapping] = useState<LedgerColumnMapping>(restored?.revenueMapping ?? emptyLedgerMapping);
  const [purchaseMapping, setPurchaseMapping] = useState<LedgerColumnMapping>(restored?.purchaseMapping ?? emptyLedgerMapping);
  const [analysis, setAnalysis] = useState<AnalysisResult | null>(null);
  const [isBusy, setIsBusy] = useState(false);
  const [errorMessage, setErrorMessage] = useState("");

  useEffect(() => {
    const session: PersistedSession = {
      reportMonth,
      reportYear,
      controlCodesText,
      sheetSelection,
      detailHeaderRow,
      revenueHeaderRow,
      purchaseHeaderRow,
      detailMapping,
      revenueMapping,
      purchaseMapping
    };
    saveSession(session);
  }, [
    reportMonth,
    reportYear,
    controlCodesText,
    sheetSelection,
    detailHeaderRow,
    revenueHeaderRow,
    purchaseHeaderRow,
    detailMapping,
    revenueMapping,
    purchaseMapping
  ]);

  const primarySheets = primaryBinding?.sheetNames ?? [];
  const controlWorkbook = secondaryBinding?.workbook ?? primaryBinding?.workbook ?? null;
  const controlSheets = secondaryBinding?.sheetNames ?? primaryBinding?.sheetNames ?? [];
  const selectionProfile = useMemo(() => profileSheetSelection(primarySheets), [primarySheets]);

  const detailAOA = useMemo(() => (primaryBinding && sheetSelection.detailSheet ? getAOA(primaryBinding.workbook, sheetSelection.detailSheet) : []), [primaryBinding, sheetSelection.detailSheet]);
  const revenueAOA = useMemo(() => (controlWorkbook && sheetSelection.revenueSheet ? getAOA(controlWorkbook, sheetSelection.revenueSheet) : []), [controlWorkbook, sheetSelection.revenueSheet]);
  const purchaseAOA = useMemo(() => (controlWorkbook && sheetSelection.purchaseSheet ? getAOA(controlWorkbook, sheetSelection.purchaseSheet) : []), [controlWorkbook, sheetSelection.purchaseSheet]);

  const detailHeaders = useMemo(() => buildHeaderLabels(detailAOA, detailHeaderRow), [detailAOA, detailHeaderRow]);
  const revenueHeaders = useMemo(() => buildHeaderLabels(revenueAOA, revenueHeaderRow), [revenueAOA, revenueHeaderRow]);
  const purchaseHeaders = useMemo(() => buildHeaderLabels(purchaseAOA, purchaseHeaderRow), [purchaseAOA, purchaseHeaderRow]);

  async function handlePrimaryUpload(file: File) {
    setIsBusy(true);
    setErrorMessage("");
    try {
      const binding = await readWorkbookFromFile(file);
      setPrimaryBinding(binding);
      const defaults = defaultSheetSelection(binding.sheetNames);
      setSheetSelection((current) => ({
        detailSheet: defaults.detailSheet || current.detailSheet,
        summarySheet: defaults.summarySheet || defaults.detailSheet || current.summarySheet,
        revenueSheet: secondaryBinding ? current.revenueSheet : defaults.revenueSheet || current.revenueSheet,
        purchaseSheet: secondaryBinding ? current.purchaseSheet : defaults.purchaseSheet || current.purchaseSheet
      }));

      const detailRows = getAOA(binding.workbook, defaults.detailSheet);
      const detailHeader = detectHeaderRow(detailRows);
      setDetailHeaderRow(detailHeader);
      setDetailMapping(autoDetectDetailMapping(buildHeaderLabels(detailRows, detailHeader)));

      const revenueRows = getAOA(binding.workbook, defaults.revenueSheet);
      const revenueHeader = detectHeaderRow(revenueRows);
      setRevenueHeaderRow(revenueHeader);
      setRevenueMapping(autoDetectLedgerMapping(buildHeaderLabels(revenueRows, revenueHeader)));

      const purchaseRows = getAOA(binding.workbook, defaults.purchaseSheet);
      const purchaseHeader = detectHeaderRow(purchaseRows);
      setPurchaseHeaderRow(purchaseHeader);
      setPurchaseMapping(autoDetectLedgerMapping(buildHeaderLabels(purchaseRows, purchaseHeader)));

      setAnalysis(null);
    } catch (error) {
      setErrorMessage(error instanceof Error ? error.message : "Unable to load primary workbook.");
    } finally {
      setIsBusy(false);
    }
  }

  async function handleSecondaryUpload(file: File) {
    setIsBusy(true);
    setErrorMessage("");
    try {
      const binding = await readWorkbookFromFile(file);
      setSecondaryBinding(binding);
      const defaults = defaultSheetSelection(binding.sheetNames);
      setSheetSelection((current) => ({
        ...current,
        revenueSheet: defaults.revenueSheet || current.revenueSheet,
        purchaseSheet: defaults.purchaseSheet || current.purchaseSheet
      }));

      const revenueRows = getAOA(binding.workbook, defaults.revenueSheet);
      const revenueHeader = detectHeaderRow(revenueRows);
      setRevenueHeaderRow(revenueHeader);
      setRevenueMapping(autoDetectLedgerMapping(buildHeaderLabels(revenueRows, revenueHeader)));

      const purchaseRows = getAOA(binding.workbook, defaults.purchaseSheet);
      const purchaseHeader = detectHeaderRow(purchaseRows);
      setPurchaseHeaderRow(purchaseHeader);
      setPurchaseMapping(autoDetectLedgerMapping(buildHeaderLabels(purchaseRows, purchaseHeader)));

      setAnalysis(null);
    } catch (error) {
      setErrorMessage(error instanceof Error ? error.message : "Unable to load reconciliation workbook.");
    } finally {
      setIsBusy(false);
    }
  }

  function rerunAutoDetection() {
    if (!primaryBinding) return;
    const defaults = defaultSheetSelection(primaryBinding.sheetNames);
    setSheetSelection((current) => ({
      detailSheet: defaults.detailSheet || current.detailSheet,
      summarySheet: defaults.summarySheet || current.summarySheet,
      revenueSheet: secondaryBinding ? current.revenueSheet : defaults.revenueSheet || current.revenueSheet,
      purchaseSheet: secondaryBinding ? current.purchaseSheet : defaults.purchaseSheet || current.purchaseSheet
    }));

    const detailRows = getAOA(primaryBinding.workbook, defaults.detailSheet);
    const detailHeader = detectHeaderRow(detailRows);
    setDetailHeaderRow(detailHeader);
    setDetailMapping(autoDetectDetailMapping(buildHeaderLabels(detailRows, detailHeader)));

    const workbookForControls = secondaryBinding?.workbook ?? primaryBinding.workbook;
    const revenueRows = getAOA(workbookForControls, sheetSelection.revenueSheet || defaults.revenueSheet);
    const revenueHeader = detectHeaderRow(revenueRows);
    setRevenueHeaderRow(revenueHeader);
    setRevenueMapping(autoDetectLedgerMapping(buildHeaderLabels(revenueRows, revenueHeader)));

    const purchaseRows = getAOA(workbookForControls, sheetSelection.purchaseSheet || defaults.purchaseSheet);
    const purchaseHeader = detectHeaderRow(purchaseRows);
    setPurchaseHeaderRow(purchaseHeader);
    setPurchaseMapping(autoDetectLedgerMapping(buildHeaderLabels(purchaseRows, purchaseHeader)));
  }

  function runAnalysis() {
    if (!primaryBinding) {
      setErrorMessage("Load the primary workbook first.");
      return;
    }
    if (!sheetSelection.detailSheet) {
      setErrorMessage("Select a detail sheet.");
      return;
    }
    setErrorMessage("");
    const result = analyzeWorkbooks({
      primaryWorkbook: primaryBinding.workbook,
      detailSheet: sheetSelection.detailSheet,
      summarySheet: sheetSelection.summarySheet || sheetSelection.detailSheet,
      controlWorkbook: controlWorkbook ?? primaryBinding.workbook,
      revenueSheet: sheetSelection.revenueSheet,
      purchaseSheet: sheetSelection.purchaseSheet,
      detailHeaderRow,
      revenueHeaderRow,
      purchaseHeaderRow,
      detailMapping,
      revenueMapping,
      purchaseMapping,
      reportMonth,
      reportYear,
      controlCodesText
    });
    setAnalysis(result);
  }

  const statusTone = !analysis ? "default" : analysis.releaseReady ? "good" : analysis.hardStop ? "bad" : "warn";

  return (
    <div className="app-shell">
      <header className="hero">
        <div className="hero-main">
          <div className="badge">Deployable forensic VAT workbench</div>
          <h1>Workbook-driven VAT reporting with drift handling and release gating</h1>
          <p>
            Upload the source workbook and optional reconciliation workbook. The parser auto-detects shifting sheets and headers, keeps summary/control rows out of detail extraction, routes rows into reporting buckets, reconciles totals and control codes, and blocks release until controls pass.
          </p>
          <div className="hero-grid">
            <div><strong>Graceful fallbacks</strong><span>Low-confidence mappings degrade to warnings and manual review instead of total failure.</span></div>
            <div><strong>Forensic trace</strong><span>Each extracted row keeps routing trace and exception markers for auditor review.</span></div>
            <div><strong>Railway-ready</strong><span>Vite build, static asset serving, config-as-code, and browser-side persistence included.</span></div>
          </div>
        </div>
        <aside className="hero-side">
          <div className="side-title"><ShieldIcon className="icon-sm" />Rule pack</div>
          <div className="rule-list">{RULES.map((rule) => <div key={rule} className="rule-chip">{rule}</div>)}</div>
        </aside>
      </header>

      <main className="layout">
        <div className="left-column">
          <SectionCard title="Workbook intake" subtitle="Primary workbook is required. Secondary workbook is optional for external revenue/purchase controls.">
            <div className="upload-grid">
              <div className="upload-box">
                <label htmlFor="primaryUpload" className="field">
                  <span>Primary VAT workbook</span>
                  <input id="primaryUpload" type="file" accept=".xlsx,.xls,.xlsm" onChange={(event) => {
                    const file = event.target.files?.[0];
                    if (file) void handlePrimaryUpload(file);
                  }} />
                </label>
                <button className="button button-secondary" type="button"><UploadIcon className="icon-sm" />Load primary</button>
                <div className="small-note">{primaryBinding ? `Loaded: ${primaryBinding.fileName}` : "No primary workbook loaded."}</div>
              </div>

              <div className="upload-box">
                <label htmlFor="secondaryUpload" className="field">
                  <span>Optional reconciliation workbook</span>
                  <input id="secondaryUpload" type="file" accept=".xlsx,.xls,.xlsm" onChange={(event) => {
                    const file = event.target.files?.[0];
                    if (file) void handleSecondaryUpload(file);
                  }} />
                </label>
                <button className="button button-secondary" type="button"><UploadIcon className="icon-sm" />Load secondary</button>
                <div className="small-note">{secondaryBinding ? `Loaded: ${secondaryBinding.fileName}` : "Control sheets will be read from the primary workbook unless a second file is loaded."}</div>
              </div>
            </div>

            <div className="field-grid three">
              <label className="field"><span>Report month</span><select value={reportMonth} onChange={(event) => setReportMonth(event.target.value)}>{MONTHS.map((month) => <option key={month} value={month}>{month}</option>)}</select></label>
              <label className="field"><span>Report year</span><input value={reportYear} onChange={(event) => setReportYear(event.target.value)} /></label>
              <label className="field"><span>Ledger control codes</span><input value={controlCodesText} onChange={(event) => setControlCodesText(event.target.value)} /></label>
            </div>

            <div className="button-row">
              <button type="button" className="button" onClick={rerunAutoDetection} disabled={!primaryBinding}><RefreshIcon className="icon-sm" />Refresh auto-detection</button>
              <button type="button" className="button button-primary" onClick={runAnalysis} disabled={!primaryBinding || isBusy}><FileIcon className="icon-sm" />Run full analysis</button>
            </div>

            {errorMessage ? <div className="callout callout-bad">{errorMessage}</div> : null}
          </SectionCard>

          <SectionCard title="Sheet bindings and drift-aware auto-detection" subtitle="Review the auto-detected targets. Candidate matches are shown from the current primary workbook.">
            <div className="field-grid two">
              <label className="field"><span>Detail sheet</span><select value={sheetSelection.detailSheet} onChange={(event) => setSheetSelection((current) => ({ ...current, detailSheet: event.target.value }))}>{primarySheets.map((sheet) => <option key={sheet} value={sheet}>{sheet}</option>)}</select></label>
              <label className="field"><span>Summary sheet</span><select value={sheetSelection.summarySheet} onChange={(event) => setSheetSelection((current) => ({ ...current, summarySheet: event.target.value }))}>{primarySheets.map((sheet) => <option key={sheet} value={sheet}>{sheet}</option>)}</select></label>
              <label className="field"><span>Revenue sheet</span><select value={sheetSelection.revenueSheet} onChange={(event) => setSheetSelection((current) => ({ ...current, revenueSheet: event.target.value }))}>{controlSheets.map((sheet) => <option key={sheet} value={sheet}>{sheet}</option>)}</select></label>
              <label className="field"><span>Purchase sheet</span><select value={sheetSelection.purchaseSheet} onChange={(event) => setSheetSelection((current) => ({ ...current, purchaseSheet: event.target.value }))}>{controlSheets.map((sheet) => <option key={sheet} value={sheet}>{sheet}</option>)}</select></label>
              <label className="field"><span>Detail header row index</span><input value={String(detailHeaderRow)} onChange={(event) => setDetailHeaderRow(Number(event.target.value || "0"))} /></label>
              <div className="field-grid two">
                <label className="field"><span>Revenue header row index</span><input value={String(revenueHeaderRow)} onChange={(event) => setRevenueHeaderRow(Number(event.target.value || "0"))} /></label>
                <label className="field"><span>Purchase header row index</span><input value={String(purchaseHeaderRow)} onChange={(event) => setPurchaseHeaderRow(Number(event.target.value || "0"))} /></label>
              </div>
            </div>

            <div className="candidate-grid">
              <div className="candidate-box">
                <h3>Detail candidates</h3>
                {selectionProfile.detailCandidates.slice(0, 4).map((candidate) => <div key={candidate.name} className="candidate-row"><span>{candidate.name}</span><strong>{candidate.score}</strong></div>)}
              </div>
              <div className="candidate-box">
                <h3>Summary candidates</h3>
                {selectionProfile.summaryCandidates.slice(0, 4).map((candidate) => <div key={candidate.name} className="candidate-row"><span>{candidate.name}</span><strong>{candidate.score}</strong></div>)}
              </div>
            </div>
          </SectionCard>

          <SectionCard title="Column maps" subtitle="Review and override the auto-detected column bindings when workbook copies drift.">
            <div className="mapping-columns">
              <div>
                <h3>Detail mapping</h3>
                <div className="field-grid two">
                  <MappingSelect label="Record ID" value={detailMapping.idCol} onChange={(next) => setDetailMapping((current) => ({ ...current, idCol: next }))} options={detailHeaders} />
                  <MappingSelect label="Description" value={detailMapping.descriptionCol} onChange={(next) => setDetailMapping((current) => ({ ...current, descriptionCol: next }))} options={detailHeaders} />
                  <MappingSelect label="Document number" value={detailMapping.docCol} onChange={(next) => setDetailMapping((current) => ({ ...current, docCol: next }))} options={detailHeaders} />
                  <MappingSelect label="Posting date" value={detailMapping.postingDateCol} onChange={(next) => setDetailMapping((current) => ({ ...current, postingDateCol: next }))} options={detailHeaders} />
                  <MappingSelect label="Report date" value={detailMapping.reportDateCol} onChange={(next) => setDetailMapping((current) => ({ ...current, reportDateCol: next }))} options={detailHeaders} />
                  <MappingSelect label="Amount" value={detailMapping.amountCol} onChange={(next) => setDetailMapping((current) => ({ ...current, amountCol: next }))} options={detailHeaders} />
                  <MappingSelect label="Currency" value={detailMapping.currencyCol} onChange={(next) => setDetailMapping((current) => ({ ...current, currencyCol: next }))} options={detailHeaders} />
                  <MappingSelect label="Treatment/status" value={detailMapping.treatmentCol} onChange={(next) => setDetailMapping((current) => ({ ...current, treatmentCol: next }))} options={detailHeaders} />
                  <MappingSelect label="Note/reason" value={detailMapping.noteCol} onChange={(next) => setDetailMapping((current) => ({ ...current, noteCol: next }))} options={detailHeaders} />
                  <MappingSelect label="Reference" value={detailMapping.referenceCol} onChange={(next) => setDetailMapping((current) => ({ ...current, referenceCol: next }))} options={detailHeaders} />
                </div>
              </div>

              <div>
                <h3>Ledger mapping</h3>
                <div className="field-grid two">
                  <MappingSelect label="Revenue code" value={revenueMapping.codeCol} onChange={(next) => setRevenueMapping((current) => ({ ...current, codeCol: next }))} options={revenueHeaders} />
                  <MappingSelect label="Revenue amount" value={revenueMapping.amountCol} onChange={(next) => setRevenueMapping((current) => ({ ...current, amountCol: next }))} options={revenueHeaders} />
                  <MappingSelect label="Revenue description" value={revenueMapping.descriptionCol} onChange={(next) => setRevenueMapping((current) => ({ ...current, descriptionCol: next }))} options={revenueHeaders} />
                  <MappingSelect label="Purchase code" value={purchaseMapping.codeCol} onChange={(next) => setPurchaseMapping((current) => ({ ...current, codeCol: next }))} options={purchaseHeaders} />
                  <MappingSelect label="Purchase amount" value={purchaseMapping.amountCol} onChange={(next) => setPurchaseMapping((current) => ({ ...current, amountCol: next }))} options={purchaseHeaders} />
                  <MappingSelect label="Purchase description" value={purchaseMapping.descriptionCol} onChange={(next) => setPurchaseMapping((current) => ({ ...current, descriptionCol: next }))} options={purchaseHeaders} />
                </div>
              </div>
            </div>
          </SectionCard>
        </div>

        <div className="right-column">
          <div className="metrics-grid">
            <MetricCard title="Primary workbook" value={primaryBinding ? "Loaded" : "Missing"} subtitle={primaryBinding?.fileName ?? "Upload required"} tone={primaryBinding ? "good" : "warn"} />
            <MetricCard title="Processing state" value={isBusy ? "Reading" : analysis ? "Analyzed" : "Idle"} subtitle={analysis ? "Results are available below" : "Run analysis after mapping"} tone={statusTone} />
            <MetricCard title="Report period" value={toMonthLabel(reportMonth, reportYear)} subtitle="Routing and defer logic use this period" />
            <MetricCard title="Control codes" value={controlCodesText.replace(/,/g, " · ")} subtitle="Revenue vs purchase parity check" />
          </div>

          {analysis ? (
            <>
              <div className={`gate-banner ${analysis.hardStop ? "gate-bad" : "gate-good"}`}>
                <div className="gate-copy">
                  <div className="gate-title">
                    {analysis.hardStop ? <AlertTriangleIcon className="icon-sm" /> : <CheckCircleIcon className="icon-sm" />}
                    {analysis.hardStop ? "Hard stop active" : analysis.releaseReady ? "Release unlocked" : "Controls passed, release still locked"}
                  </div>
                  <p>{analysis.hardStop ? "Release is blocked because one or more hard-stop checks failed or a critical drift signal was raised." : analysis.releaseReady ? "All hard-stop checks passed and no manual-review rows remain." : "Reconciliation passed, but manual-review rows still block release."}</p>
                </div>
                <div className="button-row">
                  <button className="button" type="button" onClick={() => exportAuditWorkbook(analysis, reportMonth, reportYear)}><FileIcon className="icon-sm" />Export audit workbook</button>
                  <button className="button button-primary" type="button" onClick={() => exportReleaseWorkbook(analysis, reportMonth, reportYear)} disabled={!analysis.releaseReady}><LockIcon className="icon-sm" />Export release workbook</button>
                </div>
              </div>

              <div className="metrics-grid">
                <MetricCard title="Detail total" value={formatAmount(analysis.detailTotal)} subtitle="Sum of parsed detail rows" tone="default" />
                <MetricCard title="Summary total" value={analysis.summaryTotal === null ? "Missing" : formatAmount(analysis.summaryTotal)} subtitle="Detected from summary block" tone={analysis.totalMatched ? "good" : "bad"} />
                <MetricCard title="Reportable total" value={formatAmount(analysis.reportable.reduce((sum, row) => sum + row.amount, 0))} subtitle={`${analysis.reportable.length} rows eligible`} tone="good" />
                <MetricCard title="Manual review" value={String(analysis.manualReview.length)} subtitle="Any unresolved row blocks release" tone={analysis.manualReview.length === 0 ? "good" : "warn"} />
              </div>

              <SectionCard title="Drift signals" subtitle="Warnings and errors raised by workbook drift handling and fallback logic.">
                <DataTable columns={["Severity", "Code", "Message", "Context"]} rows={analysis.driftSignals.map((signal) => [signal.severity, signal.code, signal.message, signal.context ?? ""])} maxRows={25} />
              </SectionCard>

              <SectionCard title={`Rows released to ${toMonthLabel(reportMonth, reportYear)}`}>
                <DataTable columns={["Record ID", "Amount", "Currency", "Bucket", "Treatment", "Reference"]} rows={analysis.reportable.map((row) => [row.recordId, formatAmount(row.amount), row.currency, row.routeBucket, row.treatment || row.note || row.description, row.reference])} />
              </SectionCard>

              <SectionCard title="Deferred rows">
                <DataTable columns={["Record ID", "Amount", "Currency", "Bucket", "Treatment", "Date"]} rows={analysis.deferred.map((row) => [row.recordId, formatAmount(row.amount), row.currency, row.routeBucket, row.treatment || row.note || row.description, row.postingDate || row.reportDate])} />
              </SectionCard>

              <SectionCard title="Excluded rows">
                <DataTable columns={["Record ID", "Amount", "Bucket", "Treatment"]} rows={[...analysis.excludedReported, ...analysis.excludedNonTaxable, ...analysis.excludedTransferPricing].map((row) => [row.recordId, formatAmount(row.amount), row.routeBucket, row.treatment || row.note || row.description])} />
              </SectionCard>

              <SectionCard title="Manual review blockers">
                <DataTable columns={["Record ID", "Amount", "Decision", "Exceptions", "Trace"]} rows={analysis.manualReview.map((row) => [row.recordId, formatAmount(row.amount), <DecisionPill key={`${row.recordId}-decision`} label={row.decision.replace(/_/g, " ")} />, row.exceptions.join(" | "), row.trace.join(" | ")])} />
              </SectionCard>

              <SectionCard title="Detected summary buckets">
                <DataTable columns={["Category", "Amount", "Note", "Sheet", "Row", "Trace"]} rows={analysis.summaryBuckets.map((bucket) => [bucket.category, formatAmount(bucket.amount), bucket.note, bucket.sheetName, bucket.rowNumber, bucket.trace.join(" | ")])} />
              </SectionCard>

              <SectionCard title="Bucket reconciliation">
                <DataTable columns={["Category", "Detail Amount", "Summary Amount", "Delta", "Matched"]} rows={analysis.bucketChecks.map((check) => [check.category, formatAmount(check.detailAmount), formatAmount(check.summaryAmount), formatAmount(check.delta), check.matched ? "Yes" : "No"])} />
              </SectionCard>

              <SectionCard title="Control code reconciliation">
                <DataTable columns={["Code", "Revenue Amount", "Purchase Amount", "Revenue Hits", "Purchase Hits", "Matched"]} rows={analysis.codeChecks.map((check) => [check.code, formatAmount(check.leftAmount), formatAmount(check.rightAmount), check.leftHits, check.rightHits, check.matched ? "Yes" : "No"])} />
              </SectionCard>

              <SectionCard title="Full parsed detail register">
                <DataTable columns={["Record ID", "Amount", "Currency", "Decision", "Bucket", "Exceptions", "Trace"]} rows={analysis.detailRows.map((row) => [row.recordId, formatAmount(row.amount), row.currency, <DecisionPill key={`${row.recordId}-full`} label={row.decision.replace(/_/g, " ")} />, row.routeBucket, row.exceptions.join(" | "), row.trace.join(" | ")])} maxRows={100} />
              </SectionCard>
            </>
          ) : (
            <div className="empty-panel">
              <FolderUpIcon className="icon-lg" />
              <h2>Ready for workbook analysis</h2>
              <p>Upload real workbooks, review the auto-detected sheet and column mappings, and run the analysis. The tool keeps operating through format drift by raising warnings, routing uncertain rows to manual review, and blocking release instead of failing silently.</p>
            </div>
          )}
        </div>
      </main>
    </div>
  );
}