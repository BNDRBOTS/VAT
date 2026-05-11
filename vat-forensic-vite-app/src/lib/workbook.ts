import * as XLSX from "xlsx";
import type {
  DetectionCandidate,
  DetailColumnMapping,
  LedgerColumnMapping,
  SheetSelection,
  SheetSelectionProfile,
  WorkbookBinding
} from "../types";

export const MONTHS = [
  "January","February","March","April","May","June","July","August","September","October","November","December"
] as const;

export function normalizeText(value: unknown): string {
  return String(value ?? "").replace(/\s+/g, " ").trim().toLowerCase();
}

export function displayText(value: unknown): string {
  return String(value ?? "").replace(/\s+/g, " ").trim();
}

export function toColumnLetters(index: number): string {
  let result = "";
  let n = index + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

export function monthNameToIndex(name: string): number {
  const idx = MONTHS.findIndex((value) => value.toLowerCase() === name.toLowerCase());
  return idx >= 0 ? idx : 7;
}

export function nextMonthName(name: string): string {
  const index = monthNameToIndex(name);
  return MONTHS[(index + 1) % 12];
}

export function toMonthLabel(monthName: string, year: string): string {
  return `${monthName} ${year}`.trim();
}

export function approxEqual(a: number | null, b: number | null, tolerance = 0.01): boolean {
  if (a === null || b === null) return false;
  return Math.abs(a - b) <= tolerance;
}

export function parseMoney(value: unknown): number | null {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const raw = String(value).trim();
  if (!raw) return null;
  const isNegative = /^\(.*\)$/.test(raw) || /-$/.test(raw) || /^-/.test(raw);
  const strippedAlpha = raw.replace(/[A-Za-z$€£¥₱₫₹₭₦₮₲₵₸₼₾₺₿]/g, " ");
  const cleaned = strippedAlpha.replace(/[^0-9,.\-]/g, "");
  if (!cleaned) return null;
  const unsigned = cleaned.replace(/^-/, "").replace(/-$/, "");
  const hasComma = unsigned.includes(",");
  const hasDot = unsigned.includes(".");
  let normalized = unsigned;
  if (hasComma && hasDot) {
    normalized = unsigned.lastIndexOf(",") > unsigned.lastIndexOf(".")
      ? unsigned.replace(/\./g, "").replace(",", ".")
      : unsigned.replace(/,/g, "");
  } else if (hasComma && !hasDot) {
    const parts = unsigned.split(",");
    normalized = parts.length === 2 && parts[1].length <= 2 ? unsigned.replace(",", ".") : unsigned.replace(/,/g, "");
  } else {
    normalized = unsigned.replace(/,/g, "");
  }
  const number = Number(normalized);
  if (!Number.isFinite(number)) return null;
  return isNegative ? -Math.abs(number) : number;
}

export function parseDateDisplay(value: unknown): string {
  if (value === null || value === undefined || value === "") return "";
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) return [String(parsed.y), String(parsed.m).padStart(2, "0"), String(parsed.d).padStart(2, "0")].join("-");
  }
  const raw = String(value).trim();
  if (!raw) return "";
  const iso = raw.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if (iso) return `${iso[1]}-${String(iso[2]).padStart(2, "0")}-${String(iso[3]).padStart(2, "0")}`;
  const slash = raw.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/);
  if (slash) {
    const first = Number(slash[1]);
    const second = Number(slash[2]);
    const year = slash[3].length === 2 ? `20${slash[3]}` : slash[3];
    const month = first > 12 ? second : first;
    const day = first > 12 ? first : second;
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
  }
  const date = new Date(raw);
  if (!Number.isNaN(date.getTime())) return date.toISOString().slice(0, 10);
  return raw;
}

export function extractCurrency(amountCell: unknown, currencyCell: unknown): string {
  const explicit = displayText(currencyCell);
  if (explicit) return explicit;
  const amountText = displayText(amountCell);
  const code = amountText.match(/\b([A-Z]{3})\b/);
  return code?.[1] ?? "";
}

export function rowHasUsefulCells(row: unknown[]): boolean {
  return row.some((cell) => displayText(cell) !== "");
}

export function getAOA(workbook: XLSX.WorkBook, sheetName: string): unknown[][] {
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return [];
  return XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: false,
    defval: "",
    blankrows: false
  }) as unknown[][];
}

function scoreHeaderRow(row: unknown[]): number {
  return row.reduce((score, cell) => {
    const text = displayText(cell);
    if (!text) return score;
    const numeric = parseMoney(text);
    const looksDate = /\d{4}-\d{2}-\d{2}/.test(parseDateDisplay(text)) || /\d{1,2}[/-]\d{1,2}[/-]\d{2,4}/.test(text);
    const headerHints = ["date", "amount", "description", "note", "status", "currency", "document", "invoice", "ref", "account", "code"];
    const hintBoost = headerHints.some((hint) => normalizeText(text).includes(hint)) ? 3 : 0;
    if (numeric !== null) return score - 1;
    if (looksDate) return score - 1;
    if (/[A-Za-z]/.test(text)) return score + 2 + hintBoost;
    return score;
  }, 0);
}

export function detectHeaderRow(aoa: unknown[][]): number {
  let bestIndex = 0;
  let bestScore = -Infinity;
  const scanLimit = Math.min(40, aoa.length);
  for (let i = 0; i < scanLimit; i += 1) {
    const row = aoa[i] ?? [];
    if (!rowHasUsefulCells(row)) continue;
    const score = scoreHeaderRow(row);
    if (score > bestScore) {
      bestScore = score;
      bestIndex = i;
    }
  }
  return bestIndex;
}

export function buildHeaderLabels(aoa: unknown[][], headerRow: number): string[] {
  const maxCols = Math.max(...aoa.slice(0, Math.max(headerRow + 3, 5)).map((row) => row.length), 0);
  const header = aoa[headerRow] ?? [];
  return Array.from({ length: maxCols }).map((_, index) => displayText(header[index]) || `Column ${toColumnLetters(index)}`);
}

function scoreColumn(header: string, synonyms: string[][]): number {
  const normalized = normalizeText(header);
  let best = -1;
  for (const group of synonyms) {
    let local = 0;
    let matchedAll = true;
    for (const word of group) {
      if (normalized === word) local += 6;
      else if (normalized.includes(word)) local += 3;
      else matchedAll = false;
    }
    if (matchedAll) best = Math.max(best, local);
  }
  return best;
}

export function detectColumn(headers: string[], synonymGroups: string[][]): number | null {
  let bestIndex: number | null = null;
  let bestScore = -1;
  headers.forEach((header, index) => {
    const score = scoreColumn(header, synonymGroups);
    if (score > bestScore) {
      bestScore = score;
      bestIndex = index;
    }
  });
  return bestScore > 0 ? bestIndex : null;
}

export function autoDetectDetailMapping(headers: string[]): DetailColumnMapping {
  return {
    idCol: detectColumn(headers, [["record"], ["line"], ["row"], ["entry"], ["id"]]),
    descriptionCol: detectColumn(headers, [["description"], ["text"], ["narrative"], ["memo"]]),
    docCol: detectColumn(headers, [["document"], ["doc"], ["invoice"], ["voucher"]]),
    postingDateCol: detectColumn(headers, [["posting","date"], ["book","date"], ["entry","date"], ["date"]]),
    reportDateCol: detectColumn(headers, [["report","date"], ["tax","date"], ["due","date"], ["vat","date"]]),
    amountCol: detectColumn(headers, [["amount"], ["gross"], ["tax"], ["value"], ["balance"]]),
    currencyCol: detectColumn(headers, [["currency"], ["curr"], ["ccy"]]),
    treatmentCol: detectColumn(headers, [["status"], ["treatment"], ["report"], ["classification"], ["comment"]]),
    noteCol: detectColumn(headers, [["note"], ["reason"], ["remarks"], ["comment"]]),
    referenceCol: detectColumn(headers, [["ref"], ["reference"], ["external"], ["po"], ["vendor"]])
  };
}

export function autoDetectLedgerMapping(headers: string[]): LedgerColumnMapping {
  return {
    codeCol: detectColumn(headers, [["code"], ["account"], ["gl"], ["line"]]),
    amountCol: detectColumn(headers, [["amount"], ["value"], ["balance"], ["total"]]),
    descriptionCol: detectColumn(headers, [["description"], ["text"], ["name"], ["narrative"]])
  };
}

function scoreSheetName(name: string, keywordSets: string[][]): DetectionCandidate {
  const normalized = normalizeText(name);
  let bestScore = 0;
  let matchedKeywords: string[] = [];
  for (const set of keywordSets) {
    let score = 0;
    const local: string[] = [];
    for (const keyword of set) {
      if (normalized.includes(keyword)) {
        score += keyword.length > 4 ? 4 : 3;
        local.push(keyword);
      } else {
        score = -1;
        break;
      }
    }
    if (score > bestScore) {
      bestScore = score;
      matchedKeywords = local;
    }
  }
  return { name, score: bestScore, matchedKeywords };
}

export function detectSheetCandidates(sheetNames: string[], keywordSets: string[][]): DetectionCandidate[] {
  return [...sheetNames].map((name) => scoreSheetName(name, keywordSets)).sort((a, b) => b.score - a.score || a.name.localeCompare(b.name));
}

export function profileSheetSelection(sheetNames: string[]): SheetSelectionProfile {
  return {
    detailCandidates: detectSheetCandidates(sheetNames, [["izettle"], ["0615"], ["detail"], ["fees"], ["vat"]]),
    summaryCandidates: detectSheetCandidates(sheetNames, [["0615"], ["summary"], ["vat","summary"], ["purchase","summary"]]),
    revenueCandidates: detectSheetCandidates(sheetNames, [["revenue","summary"], ["sales","summary"], ["revenue"], ["summary"]]),
    purchaseCandidates: detectSheetCandidates(sheetNames, [["purchase","summary"], ["purch","summary"], ["purchase"], ["summary"]])
  };
}

export function defaultSheetSelection(sheetNames: string[]): SheetSelection {
  const profile = profileSheetSelection(sheetNames);
  return {
    detailSheet: profile.detailCandidates[0]?.name ?? "",
    summarySheet: profile.summaryCandidates[0]?.name ?? "",
    revenueSheet: profile.revenueCandidates[0]?.name ?? "",
    purchaseSheet: profile.purchaseCandidates[0]?.name ?? ""
  };
}

export function readWorkbookFromFile(file: File): Promise<WorkbookBinding> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error(`Failed to read ${file.name}`));
    reader.onload = (event) => {
      try {
        const workbook = XLSX.read(event.target?.result, {
          type: "array",
          cellDates: false,
          cellStyles: false,
          sheetStubs: true
        });
        resolve({ fileName: file.name, workbook, sheetNames: workbook.SheetNames });
      } catch (error) {
        reject(error instanceof Error ? error : new Error("Workbook parse failed"));
      }
    };
    reader.readAsArrayBuffer(file);
  });
}

export function getCell(row: unknown[], idx: number | null): string {
  if (idx === null || idx < 0) return "";
  return displayText(row[idx]);
}

export function getCellMoney(row: unknown[], idx: number | null): number | null {
  if (idx === null || idx < 0) return null;
  return parseMoney(row[idx]);
}

export function dominantMoneyFromRow(row: unknown[]): number | null {
  const values = row.map(parseMoney).filter((value): value is number => value !== null && value !== 0);
  if (values.length === 1) return values[0];
  if (values.length < 2) return null;
  const sorted = [...values].sort((a, b) => Math.abs(b) - Math.abs(a));
  const largest = sorted[0];
  const second = sorted[1];
  return Math.abs(largest) >= Math.abs(second) * 1.5 ? largest : null;
}