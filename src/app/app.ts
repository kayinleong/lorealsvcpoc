import { App } from "@microsoft/teams.apps";
import { ChatPrompt } from "@microsoft/teams.ai";
import type { Message } from "@microsoft/teams.ai";
import { LocalStorage } from "@microsoft/teams.common";
import { OpenAIChatModel } from "@microsoft/teams.openai";
import { MessageActivity, TokenCredentials } from "@microsoft/teams.api";
import { ManagedIdentityCredential } from "@azure/identity";
import * as fs from "fs";
import * as path from "path";
import * as xlsx from "xlsx";
import AdmZip from "adm-zip";
import { XMLParser } from "fast-xml-parser";
import config from "../config";

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

const BRIEF_PATH = path.join(process.cwd(), ".data", "Brief.xlsx");
const ORIGINAL_FORMAT_PATH = path.join(process.cwd(), ".data", "Brief - Original Format.xlsx");
const PROMO_SHEET = "Promo-Voucher";
const EANS_SHEET = "Included-Excluded EANs- Voucher";

// Row indices (0-based) inside each sheet
const PROMO_HEADER_ROW = 7; // Row 8 in Excel — column names
const PROMO_DATA_ROW = 9; // Row 10 in Excel — first actual data row (row 9 is a description row)
const EANS_HEADER_ROW = 4; // Row 5 in Excel — column names
const EANS_DATA_ROW = 6; // Row 7 in Excel — first actual data row (row 6 is a description row)

// ---------------------------------------------------------------------------
// Shared model / storage / app
// ---------------------------------------------------------------------------

const storage = new LocalStorage();

function loadInstructions(): string {
  const filePath = path.join(__dirname, "instructions.txt");
  return fs.readFileSync(filePath, "utf-8").trim();
}

const instructions = loadInstructions();

const model = new OpenAIChatModel({
  model: config.openAIModelName,
  apiKey: config.openAIKey,
});

const createTokenFactory = () => {
  return async (
    scope: string | string[],
    tenantId?: string,
  ): Promise<string> => {
    const managedIdentityCredential = new ManagedIdentityCredential({
      clientId: process.env.CLIENT_ID,
    });
    const scopes = Array.isArray(scope) ? scope : [scope];
    const tokenResponse = await managedIdentityCredential.getToken(scopes, {
      tenantId,
    });
    return tokenResponse.token;
  };
};

const tokenCredentials: TokenCredentials = {
  clientId: process.env.CLIENT_ID || "",
  token: createTokenFactory(),
};

const credentialOptions =
  config.MicrosoftAppType === "UserAssignedMsi"
    ? { ...tokenCredentials }
    : undefined;

const app = new App({ ...credentialOptions, storage });

// ---------------------------------------------------------------------------
// Brief.xlsx helpers
// ---------------------------------------------------------------------------

/** Normalize Excel column headers by collapsing newlines. */
function cleanHeader(h: unknown): string {
  return String(h ?? "")
    .replace(/\r\n/g, " ")
    .replace(/\n/g, " ")
    .trim();
}

/** Extract plain text from an ExcelTS cell value (handles rich text, formulas, dates). */
function excelCellText(value: any): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "string") return value;
  if (typeof value === "number" || typeof value === "boolean")
    return String(value);
  if (value instanceof Date) return value.toISOString();
  const obj = value as unknown as Record<string, unknown>;
  if ("richText" in obj) {
    return (obj.richText as Array<{ text: string }>)
      .map((r) => r.text)
      .join("");
  }
  if ("result" in obj)
    return obj.result !== undefined ? String(obj.result) : "";
  return String(value);
}

/**
 * Read Brief.xlsx and return a tab-separated text snapshot of the two
 * relevant sheets so the LLM always has the latest data in context.
 */
function readBriefContext(): string {
  console.log(`[readBriefContext] Starting to read Brief.xlsx at ${BRIEF_PATH}`);

  if (!fs.existsSync(BRIEF_PATH)) {
    console.error(`[readBriefContext] Brief.xlsx not found at ${BRIEF_PATH}`);
    return "[BRIEF DATA]\nError: Brief.xlsx not found at .data/Brief.xlsx\n[/BRIEF DATA]";
  }

  const startTime = Date.now();
  const wb = xlsx.readFile(BRIEF_PATH);
  console.log(`[readBriefContext] Brief.xlsx loaded in ${Date.now() - startTime}ms`);

  const formatSheet = (
    sheetName: string,
    headerRow: number,
    dataRow: number,
    maxCols: number,
  ): string => {
    const ws = wb.Sheets[sheetName];
    if (!ws) {
      console.warn(`[readBriefContext] Sheet "${sheetName}" not found`);
      return `### ${sheetName}\n(sheet not found)\n`;
    }

    const rows = xlsx.utils.sheet_to_json<unknown[]>(ws, {
      header: 1,
      defval: "",
    }) as unknown[][];

    const headers = (rows[headerRow] as unknown[])
      .slice(0, maxCols)
      .map(cleanHeader);

    const dataRows = rows
      .slice(dataRow)
      .filter((row) =>
        (row as unknown[]).slice(0, maxCols).some((cell) => cell !== ""),
      )
      .map((row) =>
        (row as unknown[])
          .slice(0, maxCols)
          .map((cell) => (cell === "" ? "" : String(cell))),
      );

    const headerLine = headers.join("\t");
    const lines = dataRows.map((row) => row.join("\t"));
    const result = [
      `### ${sheetName} (${dataRows.length} rows)`,
      headerLine,
      ...lines,
    ].join("\n");

    console.log(`[readBriefContext] Formatted sheet "${sheetName}": ${dataRows.length} rows, ${result.length} characters`);
    return result;
  };

  const promoSection = formatSheet(
    PROMO_SHEET,
    PROMO_HEADER_ROW,
    PROMO_DATA_ROW,
    25,
  );
  const eansSection = formatSheet(
    EANS_SHEET,
    EANS_HEADER_ROW,
    EANS_DATA_ROW,
    11,
  );

  const result = ["[BRIEF DATA]", promoSection, "", eansSection, "[/BRIEF DATA]"].join(
    "\n",
  );

  console.log(`[readBriefContext] Total context size: ${result.length} characters (~${Math.ceil(result.length / 4)} tokens)`);
  console.log(`[readBriefContext] Completed in ${Date.now() - startTime}ms`);

  return result;
}

// ---------------------------------------------------------------------------
// Action parsing and application
// ---------------------------------------------------------------------------

interface BulkChange {
  campaign_name?: string;
  voucher_name?: string;
  field: string;
  value: string | number;
}

interface UpdateAction {
  operation: "update";
  tab: string;
  campaign_name?: string;
  voucher_name?: string;
  field?: string;          // optional — not present when changes[] is used
  value?: string | number; // optional — not present when changes[] is used
  changes?: BulkChange[];  // bulk mode: multiple field updates in one action
  updates?: BulkChange[];  // alias for changes — LLM sometimes returns "updates" instead
  rows?: BulkChange[];     // alias for changes — LLM sometimes returns "rows" instead
}

interface AddAction {
  operation: "add";
  tab: string;
  rows: Record<string, string | number | null>[];
}

type BriefAction = UpdateAction | AddAction;

/** Extract the JSON action block from LLM response text. */
function parseAction(text: string): BriefAction | null {
  const match = text.match(/<action>([\s\S]*?)<\/action>/i);
  if (!match) return null;
  try {
    return JSON.parse(match[1].trim()) as BriefAction;
  } catch {
    return null;
  }
}

/** Remove <action>...</action> block from display text. */
function stripAction(text: string): string {
  return text.replace(/<action>[\s\S]*?<\/action>/gi, "").trim();
}

/**
 * Helper to convert column number to Excel column letter (e.g., 1 -> A, 27 -> AA)
 */
function columnNumberToLetter(colNum: number): string {
  let result = '';
  while (colNum > 0) {
    colNum--; // Make it 0-based
    result = String.fromCharCode(65 + (colNum % 26)) + result;
    colNum = Math.floor(colNum / 26);
  }
  return result;
}

/** Escape text for safe insertion into XML content nodes. */
function escapeXmlText(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

/**
 * Update a single worksheet cell in raw XML text without reparsing the entire
 * worksheet. This avoids parser entity limits and preserves all existing styles.
 */
function updateWorksheetXmlCellValue(
  worksheetXML: string,
  cellRef: string,
  targetRow: number,
  newValue: string,
): { success: boolean; updatedXml?: string; error?: string } {
  const escapedCellRef = cellRef.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const cellRegex = new RegExp(
    `(<c\\b[^>]*\\br=\"${escapedCellRef}\"[^>]*>)([\\s\\S]*?)(</c>)`,
  );
  const cellMatch = worksheetXML.match(cellRegex);
  const encodedValue = escapeXmlText(newValue);

  if (cellMatch) {
    const cellOpenTag = cellMatch[1];
    const cellInnerXml = cellMatch[2];
    const cellCloseTag = cellMatch[3];

    let updatedInnerXml: string;
    if (/<v\b[^>]*>[\s\S]*?<\/v>/.test(cellInnerXml)) {
      updatedInnerXml = cellInnerXml.replace(/<v\b[^>]*>[\s\S]*?<\/v>/, `<v>${encodedValue}</v>`);
    } else {
      updatedInnerXml = `${cellInnerXml}<v>${encodedValue}</v>`;
    }

    const updatedXml = worksheetXML.replace(
      cellRegex,
      `${cellOpenTag}${updatedInnerXml}${cellCloseTag}`,
    );
    return { success: true, updatedXml };
  }

  // If the target cell does not exist, append it to the matching row.
  const rowRegex = new RegExp(
    `(<row\\b[^>]*\\br=\"${targetRow}\"[^>]*>)([\\s\\S]*?)(</row>)`,
  );
  const rowMatch = worksheetXML.match(rowRegex);

  if (!rowMatch) {
    return {
      success: false,
      error: `Could not locate row ${targetRow} in worksheet XML.`,
    };
  }

  const rowOpenTag = rowMatch[1];
  const rowInnerXml = rowMatch[2];
  const rowCloseTag = rowMatch[3];
  const newCellXml = `<c r=\"${cellRef}\"><v>${encodedValue}</v></c>`;

  const updatedXml = worksheetXML.replace(
    rowRegex,
    `${rowOpenTag}${rowInnerXml}${newCellXml}${rowCloseTag}`,
  );

  return { success: true, updatedXml };
}

/**
 * Surgical cell update - modifies ONLY the target cell value in the raw XML
 * without touching any formatting, styles, or other elements.
 * This preserves ALL Excel formatting perfectly.
 */
async function applyUpdateSurgical(action: UpdateAction): Promise<{
  success: boolean;
  oldValue?: unknown;
  error?: string;
}> {
  console.log(`[applyUpdateSurgical] Starting surgical update...`);

  if (!fs.existsSync(BRIEF_PATH)) {
    return { success: false, error: "Brief.xlsx not found." };
  }

  try {
    // Step 1: Open xlsx file as ZIP archive
    console.log(`[applyUpdateSurgical] Opening xlsx file as ZIP...`);
    const zip = new AdmZip(BRIEF_PATH);

    // Step 2: Get worksheet list to find the sheet ID
    const workbookXML = zip.readAsText("xl/workbook.xml");

    const parser = new XMLParser({
      ignoreAttributes: false,
      parseTagValue: false,
      parseAttributeValue: false,
      processEntities: false,
    });

    const workbook = parser.parse(workbookXML);

    let sheetId: string | null = null;
    const sheets = workbook.workbook?.sheets?.sheet;

    if (Array.isArray(sheets)) {
      for (const sheet of sheets) {
        if (sheet['@_name'] === action.tab) {
          sheetId = sheet['@_r:id'];
          break;
        }
      }
    } else if (sheets && sheets['@_name'] === action.tab) {
      sheetId = sheets['@_r:id'];
    }

    if (!sheetId) {
      return { success: false, error: `Sheet "${action.tab}" not found.` };
    }

    console.log(`[applyUpdateSurgical] Found sheet "${action.tab}" with ID: ${sheetId}`);

    // Step 3: Get relationships to find worksheet XML file
    const relsXML = zip.readAsText("xl/_rels/workbook.xml.rels");
    const rels = parser.parse(relsXML);

    let worksheetPath: string | null = null;
    const relationships = rels.Relationships?.Relationship;

    if (Array.isArray(relationships)) {
      for (const rel of relationships) {
        if (rel['@_Id'] === sheetId) {
          worksheetPath = "xl/" + rel['@_Target'];
          break;
        }
      }
    } else if (relationships && relationships['@_Id'] === sheetId) {
      worksheetPath = "xl/" + relationships['@_Target'];
    }

    if (!worksheetPath) {
      return { success: false, error: `Worksheet path not found for sheet "${action.tab}".` };
    }

    console.log(`[applyUpdateSurgical] Worksheet XML path: ${worksheetPath}`);

    // Step 4: Load worksheet XML as text and edit target cell directly
    const worksheetXML = zip.readAsText(worksheetPath);

    // Step 5: Find target cell coordinates using xlsx library for header/data analysis
    const tempWorkbook = xlsx.readFile(BRIEF_PATH);
    const tempWorksheet = tempWorkbook.Sheets[action.tab];

    if (!tempWorksheet) {
      return { success: false, error: `Temporary worksheet "${action.tab}" not found.` };
    }

    const rows = xlsx.utils.sheet_to_json<unknown[]>(tempWorksheet, {
      header: 1,
      defval: "",
    }) as unknown[][];

    const headerRowIdx = action.tab === PROMO_SHEET ? PROMO_HEADER_ROW : EANS_HEADER_ROW;
    const dataRowIdx = action.tab === PROMO_SHEET ? PROMO_DATA_ROW : EANS_DATA_ROW;

    // Find target column
    const headers = (rows[headerRowIdx] as unknown[]).map((h) =>
      cleanHeader(String(h ?? ""))
    );
    const fieldColIdx = headers.indexOf(action.field);

    if (fieldColIdx === -1) {
      return {
        success: false,
        error: `Field "${action.field}" not found in headers.`,
      };
    }

    // Find all matching rows
    const campaignColIdx = 1;
    const voucherColIdx = action.tab === PROMO_SHEET ? 11 : 1;

    const targetRowIdxs: number[] = [];
    for (let r = dataRowIdx; r < rows.length; r++) {
      const row = rows[r] as unknown[];
      const campaignVal = String(row[campaignColIdx] ?? "");
      const voucherVal = String(row[voucherColIdx] ?? "");

      if (voucherVal.trim() === "") continue;

      const campaignMatch =
        action.tab !== PROMO_SHEET ||
        !action.campaign_name ||
        campaignVal.toLowerCase().includes(action.campaign_name.toLowerCase());

      const voucherMatch =
        !action.voucher_name ||
        voucherVal.toLowerCase().includes(action.voucher_name.toLowerCase());

      if (campaignMatch && voucherMatch) {
        targetRowIdxs.push(r);
      }
    }

    if (targetRowIdxs.length === 0) {
      return {
        success: false,
        error: `Record not found.`,
      };
    }

    // Convert to Excel cell reference (e.g., "P10")
    const targetCol = columnNumberToLetter(fieldColIdx + 1);

    // Get old value from first match
    const oldValue = String(rows[targetRowIdxs[0]][fieldColIdx] ?? "");
    const normalizedUpdateValue = /Discount Percentage/i.test(action.field) && typeof action.value === "number"
      ? action.value / 100
      : action.value;
    console.log(`[applyUpdateSurgical] Old value: "${oldValue}", New value: "${normalizedUpdateValue}"`);

    // Step 6: Surgically update all matching cells in raw XML text
    let currentXml = worksheetXML;
    for (const targetRowIdx of targetRowIdxs) {
      const targetRow = targetRowIdx + 1;
      const cellRef = `${targetCol}${targetRow}`;
      console.log(`[applyUpdateSurgical] Target cell: ${cellRef} (Row ${targetRow}, Col ${targetCol})`);

      const xmlUpdateResult = updateWorksheetXmlCellValue(
        currentXml,
        cellRef,
        targetRow,
        String(normalizedUpdateValue),
      );

      if (!xmlUpdateResult.success || !xmlUpdateResult.updatedXml) {
        return {
          success: false,
          error: xmlUpdateResult.error ?? "Could not update target cell in worksheet XML.",
        };
      }

      console.log(`[applyUpdateSurgical] Updated cell ${cellRef} via raw XML patch`);
      currentXml = xmlUpdateResult.updatedXml;
    }

    // Step 7: Write modified XML back to ZIP
    const modifiedXML = currentXml;
    zip.updateFile(worksheetPath, Buffer.from(modifiedXML, 'utf8'));

    // Step 8: Save the modified ZIP as xlsx
    console.log(`[applyUpdateSurgical] Writing modified xlsx file...`);
    zip.writeZip(BRIEF_PATH);

    console.log(`[applyUpdateSurgical] ✓ Surgical update completed successfully`);
    return { success: true, oldValue };

  } catch (error) {
    console.error(`[applyUpdateSurgical] Error:`, error);
    return {
      success: false,
      error: `Surgical update failed: ${error instanceof Error ? error.message : String(error)}`
    };
  }
}

// ---------------------------------------------------------------------------
// applyAdd helpers
// ---------------------------------------------------------------------------

/** Convert "dd-mmm-yyyy" (or dd/mm/yyyy) to an Excel date serial number. */
function dateStringToExcelSerial(dateStr: string): number {
  const months: Record<string, number> = {
    jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5,
    jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11,
  };
  const m = dateStr.match(/^(\d{1,2})[\/\-\s]([a-zA-Z]+|\d{1,2})[\/\-\s](\d{4})$/);
  if (!m) return 0;
  const day = parseInt(m[1], 10);
  const monthKey = m[2].toLowerCase().slice(0, 3);
  const month = /^\d+$/.test(m[2]) ? parseInt(m[2], 10) - 1 : months[monthKey];
  const year = parseInt(m[3], 10);
  if (month === undefined || isNaN(month)) return 0;
  const date = new Date(year, month, day);
  // Excel epoch is 1899-12-30 (accounts for Excel's 1900 leap-year bug)
  const excelEpoch = new Date(1899, 11, 30);
  return Math.round((date.getTime() - excelEpoch.getTime()) / 86400000);
}

/** Convert "HH:mm:ss" to an Excel time fraction (0–1). */
function timeStringToExcelFraction(timeStr: string): number {
  const m = timeStr.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (!m) return 0;
  const h = parseInt(m[1], 10);
  const min = parseInt(m[2], 10);
  const sec = parseInt(m[3] ?? "0", 10);
  return (h * 3600 + min * 60 + sec) / 86400;
}

/**
 * Convert a user-supplied field value to the type Excel expects.
 * Date cells → serial number, time cells → fraction, others → unchanged.
 * "Time" is checked before "Date" so "Live Date (collection time)" → fraction.
 */
function processAddValue(
  fieldName: string,
  value: string | number | null
): string | number | null {
  if (value === null || value === undefined) return null;
  const isPercentField = /Discount Percentage/i.test(fieldName);
  if (typeof value === "number") {
    return isPercentField ? value / 100 : value;
  }
  const v = String(value).trim();
  if (v === "") return null;
  if (isPercentField) {
    const n = parseFloat(v);
    if (!isNaN(n)) return n / 100;
  }
  if (/time/i.test(fieldName)) {
    const frac = timeStringToExcelFraction(v);
    if (frac > 0 || v.startsWith("0:")) return frac;
  }
  if (/date/i.test(fieldName)) {
    const serial = dateStringToExcelSerial(v);
    if (serial > 0) return serial;
  }
  return v;
}

/**
 * Write a value into a specific cell inside a row's inner XML string.
 * - Preserves the cell's existing style (s="...") attribute.
 * - Removes any formula (<f>...</f>) so the written value is authoritative.
 * - Numbers → plain <v>; strings → inline string (t="inlineStr").
 */
function writeCellInRowXml(
  rowInnerXml: string,
  cellRef: string,
  value: string | number
): string {
  const escaped = escapeXmlText(String(value));
  const isNumber = typeof value === "number";

  const cellPattern = new RegExp(
    `<c\\s+r="${cellRef}"([^>]*)(?:/>|>[\\s\\S]*?</c>)`
  );
  const match = rowInnerXml.match(cellPattern);

  const sMatch = match?.[1]?.match(/\bs="([^"]+)"/);
  const styleAttr = sMatch ? ` s="${sMatch[1]}"` : "";

  const newCell = isNumber
    ? `<c r="${cellRef}"${styleAttr}><v>${escaped}</v></c>`
    : `<c r="${cellRef}"${styleAttr} t="inlineStr"><is><t>${escaped}</t></is></c>`;

  if (match) {
    return rowInnerXml.replace(cellPattern, newCell);
  }
  return rowInnerXml + newCell;
}

/**
 * Surgically write new data rows into an existing worksheet.
 *
 * The Promo-Voucher sheet has 966 pre-styled rows (with formulas and styles
 * already in place). New data must go into the first empty rows AFTER the last
 * row that has a non-empty Campaign Name — NOT appended after row 966.
 *
 * Strategy (no xlsx.writeFile — formatting is fully preserved):
 *  1. Open Brief.xlsx as a raw ZIP.
 *  2. Locate the worksheet XML via workbook.xml + its _rels.
 *  3. Use xlsx (read-only) to find the last non-empty data row and headers.
 *  4. For each new row, find the target row in XML and write values cell-by-cell,
 *     preserving styles and removing formula overrides where needed.
 *  5. Write only the changed worksheet entry back into the ZIP.
 */
async function applyAdd(action: AddAction): Promise<{
  success: boolean;
  rowsAdded?: number;
  error?: string;
}> {
  console.log(`[applyAdd] Starting surgical add: ${action.rows.length} row(s) to "${action.tab}"`);

  if (!fs.existsSync(BRIEF_PATH)) {
    return { success: false, error: "Brief.xlsx not found." };
  }

  try {
    // ── Step 1: locate the worksheet XML inside the ZIP ─────────────────────
    const zip = new AdmZip(BRIEF_PATH);

    const parser = new XMLParser({
      ignoreAttributes: false,
      parseTagValue: false,
      parseAttributeValue: false,
      processEntities: false,
    });

    const workbookXML = zip.readAsText("xl/workbook.xml");
    const workbook = parser.parse(workbookXML);

    let sheetId: string | null = null;
    const sheets = workbook.workbook?.sheets?.sheet;
    if (Array.isArray(sheets)) {
      for (const sheet of sheets) {
        if (sheet["@_name"] === action.tab) { sheetId = sheet["@_r:id"]; break; }
      }
    } else if (sheets?.["@_name"] === action.tab) {
      sheetId = sheets["@_r:id"];
    }
    if (!sheetId) return { success: false, error: `Sheet "${action.tab}" not found.` };

    const wbRelsXML = zip.readAsText("xl/_rels/workbook.xml.rels");
    const wbRels = parser.parse(wbRelsXML);

    let worksheetPath: string | null = null;
    const wbRelationships = wbRels.Relationships?.Relationship;
    if (Array.isArray(wbRelationships)) {
      for (const rel of wbRelationships) {
        if (rel["@_Id"] === sheetId) { worksheetPath = "xl/" + rel["@_Target"]; break; }
      }
    } else if (wbRelationships?.["@_Id"] === sheetId) {
      worksheetPath = "xl/" + wbRelationships["@_Target"];
    }
    if (!worksheetPath) return { success: false, error: `Worksheet path not found for "${action.tab}".` };

    // ── Step 2: use xlsx (read-only) to get headers + last data row ──────────
    const tempWb = xlsx.readFile(BRIEF_PATH);
    const tempWs = tempWb.Sheets[action.tab];
    if (!tempWs) return { success: false, error: `Worksheet "${action.tab}" not found.` };

    const sheetRows = xlsx.utils.sheet_to_json<unknown[]>(tempWs, { header: 1, defval: "" }) as unknown[][];
    const headerRowIdx = action.tab === PROMO_SHEET ? PROMO_HEADER_ROW : EANS_HEADER_ROW;
    const dataRowIdx  = action.tab === PROMO_SHEET ? PROMO_DATA_ROW  : EANS_DATA_ROW;
    const headers = (sheetRows[headerRowIdx] as unknown[]).map((h) => cleanHeader(String(h ?? "")));

    // Find the last row with a real Campaign Name / Voucher Name.
    // Skip placeholder rows (value contains "<" — e.g. "<Placeholder>") which
    // are template rows pre-filled at the bottom of the sheet.
    const keyColIdx = 1; // Campaign Name (Promo-Voucher) or Voucher Name (EANs)
    let lastDataExcelRow = headerRowIdx + 1; // fallback: just after header
    for (let r = sheetRows.length - 1; r >= dataRowIdx; r--) {
      const row = sheetRows[r] as unknown[];
      const val = String(row[keyColIdx] ?? "").trim();
      if (val !== "" && !val.includes("<")) {
        lastDataExcelRow = r + 1; // xlsx is 0-indexed; Excel rows are 1-indexed
        break;
      }
    }
    console.log(`[applyAdd] Last data Excel row: ${lastDataExcelRow}. New rows start at: ${lastDataExcelRow + 1}`);

    // ── Step 3: write values into existing rows, one row at a time ───────────
    let worksheetXML = zip.readAsText(worksheetPath);

    for (let i = 0; i < action.rows.length; i++) {
      const targetExcelRow = lastDataExcelRow + 1 + i;
      const fieldMap = action.rows[i];

      // Find the existing <row r="N" ...>...</row> element
      const rowPattern = new RegExp(
        `(<row\\b[^>]*\\br="${targetExcelRow}"[^>]*>)([\\s\\S]*?)(</row>)`
      );
      const rowMatch = worksheetXML.match(rowPattern);

      if (!rowMatch) {
        // Row not pre-defined — create a minimal new row
        const cells: string[] = [];
        for (const [field, rawValue] of Object.entries(fieldMap)) {
          const processed = processAddValue(field, rawValue);
          if (processed === null) continue;
          const colIdx = headers.indexOf(cleanHeader(field));
          if (colIdx === -1) continue;
          const cellRef = `${columnNumberToLetter(colIdx + 1)}${targetExcelRow}`;
          const escaped = escapeXmlText(String(processed));
          cells.push(
            typeof processed === "number"
              ? `<c r="${cellRef}"><v>${escaped}</v></c>`
              : `<c r="${cellRef}" t="inlineStr"><is><t>${escaped}</t></is></c>`
          );
        }
        worksheetXML = worksheetXML.replace(
          /<\/sheetData>/,
          `<row r="${targetExcelRow}">${cells.join("")}</row></sheetData>`
        );
        console.log(`[applyAdd] Created new row ${targetExcelRow}`);
        continue;
      }

      // Row exists — update each field's cell in-place
      let rowInner = rowMatch[2];
      for (const [field, rawValue] of Object.entries(fieldMap)) {
        const processed = processAddValue(field, rawValue);
        if (processed === null) continue;
        const colIdx = headers.indexOf(cleanHeader(field));
        if (colIdx === -1) {
          console.warn(`[applyAdd] Field "${field}" not in headers — skipped`);
          continue;
        }
        const cellRef = `${columnNumberToLetter(colIdx + 1)}${targetExcelRow}`;
        rowInner = writeCellInRowXml(rowInner, cellRef, processed);
      }

      worksheetXML = worksheetXML.replace(
        rowPattern,
        `${rowMatch[1]}${rowInner}${rowMatch[3]}`
      );
      console.log(`[applyAdd] Updated existing row ${targetExcelRow}`);
    }

    // ── Step 4: write back — only the worksheet entry in the ZIP changes ─────
    zip.updateFile(worksheetPath, Buffer.from(worksheetXML, "utf8"));
    zip.writeZip(BRIEF_PATH);

    console.log(`[applyAdd] ✓ Added ${action.rows.length} row(s) starting at Excel row ${lastDataExcelRow + 1}`);
    return { success: true, rowsAdded: action.rows.length };

  } catch (error) {
    console.error(`[applyAdd] Error:`, error);
    return {
      success: false,
      error: `Add failed: ${error instanceof Error ? error.message : String(error)}`,
    };
  }
}

/**
 * Apply a validated update action directly to Brief.xlsx.
 * Uses surgical XML update as primary method (preserves ALL formatting),
 * with xlsx library as fallback if surgical update fails.
 */
async function applyUpdate(action: UpdateAction): Promise<{
  success: boolean;
  oldValue?: unknown;
  error?: string;
}> {
  console.log(`[applyUpdate] Starting update:`, JSON.stringify(action, null, 2));

  if (!fs.existsSync(BRIEF_PATH)) {
    console.error(`[applyUpdate] Brief.xlsx not found at ${BRIEF_PATH}`);
    return { success: false, error: "Brief.xlsx not found." };
  }

  const startTime = Date.now();

  // Try surgical update first - preserves ALL formatting
  console.log(`[applyUpdate] Attempting surgical XML update (preserves all formatting)...`);

  try {
    const result = await applyUpdateSurgical(action);

    if (result.success) {
      console.log(`[applyUpdate] ✓ Surgical update completed in ${Date.now() - startTime}ms`);
      return result;
    } else {
      console.warn(`[applyUpdate] Surgical update failed: ${result.error}`);
      console.warn(`[applyUpdate] Falling back to xlsx library`);
    }
  } catch (error) {
    const errorMsg = error instanceof Error ? error.message : String(error);
    console.error(`[applyUpdate] Surgical update error: ${errorMsg}`);
    console.warn(`[applyUpdate] Falling back to xlsx library`);
  }

  // Fallback to xlsx library (fast but may lose some formatting)
  console.log(`[applyUpdate] Using xlsx library fallback...`);
  const xlsxStartTime = Date.now();

  const workbook = xlsx.readFile(BRIEF_PATH);
  console.log(`[applyUpdate] Workbook loaded with xlsx in ${Date.now() - xlsxStartTime}ms`);

  const worksheet = workbook.Sheets[action.tab];
  if (!worksheet) {
    console.error(`[applyUpdate] Sheet "${action.tab}" not found`);
    return { success: false, error: `Sheet "${action.tab}" not found.` };
  }

  // Convert sheet to array of arrays for easier manipulation
  const rows = xlsx.utils.sheet_to_json<unknown[]>(worksheet, {
    header: 1,
    defval: "",
  }) as unknown[][];

  const headerRowIdx = action.tab === PROMO_SHEET ? PROMO_HEADER_ROW : EANS_HEADER_ROW;
  const dataRowIdx = action.tab === PROMO_SHEET ? PROMO_DATA_ROW : EANS_DATA_ROW;

  // Find the target column
  const headers = (rows[headerRowIdx] as unknown[]).map((h) =>
    cleanHeader(String(h ?? ""))
  );
  const fieldColIdx = headers.indexOf(action.field);

  if (fieldColIdx === -1) {
    console.error(`[applyUpdate] Field "${action.field}" not found`);
    return {
      success: false,
      error: `Field "${action.field}" not found in "${action.tab}" headers.`,
    };
  }

  const campaignColIdx = 1;
  const voucherColIdx = action.tab === PROMO_SHEET ? 11 : 1;

  const targetRowIdxs: number[] = [];
  for (let r = dataRowIdx; r < rows.length; r++) {
    const row = rows[r] as unknown[];
    const campaignVal = String(row[campaignColIdx] ?? "");
    const voucherVal = String(row[voucherColIdx] ?? "");

    if (voucherVal.trim() === "") continue;

    const campaignMatch =
      action.tab !== PROMO_SHEET ||
      !action.campaign_name ||
      campaignVal.toLowerCase().includes(action.campaign_name.toLowerCase());

    const voucherMatch =
      !action.voucher_name ||
      voucherVal.toLowerCase().includes(action.voucher_name.toLowerCase());

    if (campaignMatch && voucherMatch) {
      console.log(`[applyUpdate] Found matching record at row index ${r}`);
      targetRowIdxs.push(r);
    }
  }

  if (targetRowIdxs.length === 0) {
    return {
      success: false,
      error: `Record not found — campaign: "${action.campaign_name}", voucher: "${action.voucher_name}".`,
    };
  }

  const firstTargetRow = rows[targetRowIdxs[0]] as unknown[];
  const oldValue = String(firstTargetRow[fieldColIdx] ?? "");

  const normalizedFallbackValue = /Discount Percentage/i.test(action.field) && typeof action.value === "number"
    ? action.value / 100
    : action.value;
  console.log(`[applyUpdate] Updating ${targetRowIdxs.length} matching row(s) - Old: "${oldValue}", New: "${normalizedFallbackValue}"`);
  for (const targetRowIdx of targetRowIdxs) {
    (rows[targetRowIdx] as unknown[])[fieldColIdx] = normalizedFallbackValue;
  }

  // Convert back to worksheet, copying metadata to preserve basic formatting
  const newWorksheet = xlsx.utils.aoa_to_sheet(rows);

  if (worksheet['!ref']) newWorksheet['!ref'] = worksheet['!ref'];
  if (worksheet['!cols']) newWorksheet['!cols'] = worksheet['!cols'];
  if (worksheet['!rows']) newWorksheet['!rows'] = worksheet['!rows'];
  if (worksheet['!merges']) newWorksheet['!merges'] = worksheet['!merges'];

  workbook.Sheets[action.tab] = newWorksheet;

  const writeStartTime = Date.now();
  xlsx.writeFile(workbook, BRIEF_PATH);
  console.log(`[applyUpdate] File written in ${Date.now() - writeStartTime}ms`);
  console.log(`[applyUpdate] ✓ xlsx fallback update completed in ${Date.now() - xlsxStartTime}ms`);

  return { success: true, oldValue };
}

// ---------------------------------------------------------------------------
// Post-operation formatting validation
// ---------------------------------------------------------------------------

/**
 * Extract a named XML element's full outer content from a string.
 * Returns the matched block or an empty string if not found.
 */
function extractXmlBlock(xml: string, tag: string): string {
  const m = xml.match(new RegExp(`<${tag}\\b[^>]*(?:/>|>[\\s\\S]*?</${tag}>)`));
  return m ? m[0] : "";
}

/**
 * Validate that Brief.xlsx formatting has not been corrupted by comparing
 * key structural XML sections against the original reference file.
 *
 * Checks:
 *  1. xl/styles.xml  — must be byte-identical (xlsx fallback rewrites this)
 *  2. Per-sheet: <cols> block  — column widths / default styles must be intact
 *  3. Per-sheet: <sheetFormatPr> — default row/col dimensions must be intact
 *
 * Returns a list of human-readable issues (empty = all good).
 */
async function validateBriefFormatting(): Promise<{
  valid: boolean;
  issues: string[];
}> {
  const issues: string[] = [];

  if (!fs.existsSync(ORIGINAL_FORMAT_PATH)) {
    console.warn(`[validateFormatting] Reference file not found at ${ORIGINAL_FORMAT_PATH} — skipping format check`);
    return { valid: true, issues: [] };
  }

  try {
    const currentZip = new AdmZip(BRIEF_PATH);
    const refZip = new AdmZip(ORIGINAL_FORMAT_PATH);

    // ── 1. styles.xml must be identical ─────────────────────────────────────
    const currentStyles = currentZip.readAsText("xl/styles.xml");
    const refStyles = refZip.readAsText("xl/styles.xml");
    if (currentStyles !== refStyles) {
      issues.push("xl/styles.xml differs from reference — cell styles may have been overwritten (xlsx fallback likely triggered)");
    }

    // ── 2. Per-sheet structural checks ──────────────────────────────────────
    const parser = new XMLParser({
      ignoreAttributes: false,
      parseTagValue: false,
      parseAttributeValue: false,
      processEntities: false,
    });

    const wbXml = currentZip.readAsText("xl/workbook.xml");
    const wbRelsXml = currentZip.readAsText("xl/_rels/workbook.xml.rels");
    const wb = parser.parse(wbXml);
    const wbRels = parser.parse(wbRelsXml);

    const sheets = wb.workbook?.sheets?.sheet;
    const sheetList: Array<{ name: string; rId: string }> = [];
    if (Array.isArray(sheets)) {
      for (const s of sheets) sheetList.push({ name: s["@_name"], rId: s["@_r:id"] });
    } else if (sheets) {
      sheetList.push({ name: sheets["@_name"], rId: sheets["@_r:id"] });
    }

    const relationships = wbRels.Relationships?.Relationship;
    const relMap: Record<string, string> = {};
    if (Array.isArray(relationships)) {
      for (const r of relationships) relMap[r["@_Id"]] = r["@_Target"];
    } else if (relationships) {
      relMap[relationships["@_Id"]] = relationships["@_Target"];
    }

    for (const { name, rId } of sheetList) {
      const target = relMap[rId];
      if (!target) continue;
      const wsPath = `xl/${target}`;

      let currentWsXml: string;
      let refWsXml: string;
      try {
        currentWsXml = currentZip.readAsText(wsPath);
        refWsXml = refZip.readAsText(wsPath);
      } catch {
        // Sheet path may differ in the ref file — skip gracefully
        continue;
      }

      const currentCols = extractXmlBlock(currentWsXml, "cols");
      const refCols = extractXmlBlock(refWsXml, "cols");
      if (currentCols !== refCols) {
        issues.push(`Sheet "${name}": <cols> block differs — column widths or styles may be corrupted`);
      }

      const currentFmtPr = extractXmlBlock(currentWsXml, "sheetFormatPr");
      const refFmtPr = extractXmlBlock(refWsXml, "sheetFormatPr");
      if (currentFmtPr !== refFmtPr) {
        issues.push(`Sheet "${name}": <sheetFormatPr> differs — default row/column dimensions may be altered`);
      }
    }

  } catch (error) {
    const msg = error instanceof Error ? error.message : String(error);
    console.error(`[validateFormatting] Unexpected error:`, error);
    issues.push(`Format validation failed with error: ${msg}`);
  }

  if (issues.length === 0) {
    console.log(`[validateFormatting] ✓ Formatting intact — all checks passed`);
  } else {
    console.warn(`[validateFormatting] ⚠️  ${issues.length} formatting issue(s) detected:`);
    for (const issue of issues) console.warn(`[validateFormatting]   • ${issue}`);
  }

  return { valid: issues.length === 0, issues };
}

// ---------------------------------------------------------------------------
// Message handler
// ---------------------------------------------------------------------------

/**
 * Attempt to send an activity, retrying once on transient network errors.
 * Returns true if delivered successfully, false otherwise.
 * Never throws — all errors are caught and logged.
 */
async function safeSend<T>(
  send: (activity: T) => Promise<unknown>,
  activity: T,
  retries = 1
): Promise<boolean> {
  for (let attempt = 0; attempt <= retries; attempt++) {
    try {
      console.log(`[safeSend] Attempt ${attempt + 1}/${retries + 1} - Sending message...`);
      await send(activity);
      console.log(`[safeSend] ✓ Message sent successfully`);
      return true;
    } catch (err: unknown) {
      const code = (err as NodeJS.ErrnoException)?.code;
      const message = err instanceof Error ? err.message : String(err);
      const isTransient =
        code === "ECONNRESET" || code === "ETIMEDOUT" || code === "ECONNREFUSED";

      console.error(`[safeSend] ✗ Send failed on attempt ${attempt + 1}:`, {
        code,
        message,
        isTransient,
        willRetry: isTransient && attempt < retries
      });

      if (isTransient && attempt < retries) {
        const delay = 1000;
        console.warn(
          `[safeSend] Transient error (${code}), retrying after ${delay}ms (${attempt + 1}/${retries})...`
        );
        await new Promise((r) => setTimeout(r, delay));
        continue;
      }

      console.error(
        `[safeSend] Failed to deliver message after ${attempt + 1} attempt(s). Full error:`,
        err
      );
      return false;
    }
  }
  return false;
}

app.on("message", async ({ send, activity }) => {
  const requestId = `${Date.now()}-${Math.random().toString(36).substring(7)}`;
  const conversationKey = `${activity.conversation.id}/${activity.from.id}`;
  const userText = activity.text ?? "";

  console.log(`[${requestId}] ===== NEW MESSAGE =====`);
  console.log(`[${requestId}] From: ${activity.from.name} (${activity.from.id})`);
  console.log(`[${requestId}] Conversation: ${activity.conversation.id}`);
  console.log(`[${requestId}] User message: "${userText}"`);

  const messages: Message[] = storage.get(conversationKey) ?? [];
  console.log(`[${requestId}] Conversation history: ${messages.length} messages`);

  try {
    // Prepend the latest Brief.xlsx snapshot to every user message so the LLM
    // always works from the current data without relying on stale history.
    console.log(`[${requestId}] Loading Brief context...`);
    const briefContext = readBriefContext();
    const userMessage = `${briefContext}\n\n${userText}`;

    console.log(`[${requestId}] Total message size: ${userMessage.length} characters (~${Math.ceil(userMessage.length / 4)} tokens)`);

    // Check if context is too large (OpenAI gpt-4o has 128k token limit)
    const estimatedTokens = Math.ceil(userMessage.length / 4);
    if (estimatedTokens > 100000) {
      console.warn(`[${requestId}] ⚠️  Context size (${estimatedTokens} tokens) is very large and may cause issues`);
    }

    console.log(`[${requestId}] Sending request to LLM (model: ${config.openAIModelName})...`);
    const llmStartTime = Date.now();

    const prompt = new ChatPrompt({ messages, instructions, model });

    // Call LLM with timeout protection (120 seconds)
    const LLM_TIMEOUT = 120000;
    const timeoutPromise = new Promise<never>((_, reject) => {
      setTimeout(() => reject(new Error(`LLM call timed out after ${LLM_TIMEOUT}ms`)), LLM_TIMEOUT);
    });

    const response = await Promise.race([
      prompt.send(userMessage),
      timeoutPromise
    ]);

    const llmDuration = Date.now() - llmStartTime;
    console.log(`[${requestId}] LLM response received in ${llmDuration}ms`);

    const rawText: string = response.content ?? "";
    console.log(`[${requestId}] LLM response length: ${rawText.length} characters`);
    console.log(`[${requestId}] LLM response preview: "${rawText.substring(0, 200)}${rawText.length > 200 ? '...' : ''}"`);

    // Strip Brief data from stored conversation history to keep it lean.
    // The last entry added by prompt.send() is the user message — replace
    // its content with just the original user text.
    const lastUserIdx = [...messages].map((m) => m.role).lastIndexOf("user");
    if (lastUserIdx !== -1) {
      (messages[lastUserIdx] as Message).content = userText;
    }

    // Check for an update action from the LLM
    const action = parseAction(rawText);
    if (action) {
      console.log(`[${requestId}] Action parsed:`, JSON.stringify(action, null, 2));
    } else {
      console.log(`[${requestId}] No action block found in response`);
    }

    let displayText = stripAction(rawText);

    // Ensure we always have a response to display
    if (!displayText || displayText.trim() === "") {
      console.warn(`[${requestId}] Response is empty after stripping action, using fallback message`);
      displayText = "I understand you're asking about the brief, but I need more information. Could you please specify:\n\n- What would you like to know? (e.g., show me the current value, list all campaigns)\n- What would you like to change? (e.g., update Min Spend to RM1000)\n\nFor example: \"Show me the Min Spend for PDS x Birthday Full Day\" or \"Update PDS x Birthday Full Day Min Spend to RM1000\"";
    }

    if (action?.operation === "update") {
      console.log(`[${requestId}] Applying update action...`);
      const updateAction = action as UpdateAction;
      const entries: BulkChange[] = updateAction.changes ?? updateAction.updates ?? updateAction.rows ?? [
        {
          campaign_name: updateAction.campaign_name,
          voucher_name: updateAction.voucher_name,
          field: updateAction.field!,
          value: updateAction.value!,
        },
      ];

      let allSucceeded = true;
      const oldValues: string[] = [];
      for (const entry of entries) {
        const singleAction: UpdateAction = {
          operation: "update",
          tab: updateAction.tab,
          campaign_name: entry.campaign_name,
          voucher_name: entry.voucher_name,
          field: entry.field,
          value: entry.value,
        };
        const result = await applyUpdate(singleAction);
        if (result.success) {
          if (result.oldValue !== undefined) oldValues.push(String(result.oldValue));
        } else {
          allSucceeded = false;
          console.error(`[${requestId}] ✗ Update failed for entry:`, entry, result.error);
        }
      }

      if (allSucceeded) {
        console.log(`[${requestId}] ✓ All updates applied. Old values: ${oldValues.join(", ")}`);
        const fmtCheck = await validateBriefFormatting();
        if (!fmtCheck.valid) {
          console.warn(`[${requestId}] ⚠️  Formatting issues after update: ${fmtCheck.issues.join("; ")}`);
        }
        displayText +=
          "\n\nBrief is updated: [Download the latest Brief.xlsx here](.data/Brief.xlsx)";
      } else {
        displayText +=
          "\n\n⚠️ One or more updates could not be saved. Please contact your admin.";
      }
    } else if (action?.operation === "add") {
      console.log(`[${requestId}] Applying add action...`);
      const result = await applyAdd(action as AddAction);
      if (result.success) {
        console.log(`[${requestId}] ✓ Add applied successfully. Rows added: ${result.rowsAdded}`);
        const fmtCheck = await validateBriefFormatting();
        if (!fmtCheck.valid) {
          console.warn(`[${requestId}] ⚠️  Formatting issues after add: ${fmtCheck.issues.join("; ")}`);
        }
        displayText +=
          `\n\n${result.rowsAdded} row(s) added. Brief is updated: [Download the latest Brief.xlsx here](.data/Brief.xlsx)`;
      } else {
        console.error(`[${requestId}] ✗ Add failed:`, result.error);
        displayText +=
          "\n\n⚠️ The brief could not be saved automatically. Please contact your admin.";
      }
    }

    console.log(`[${requestId}] Sending response to user...`);
    const sendSuccess = await safeSend(send, new MessageActivity(displayText).addAiGenerated().addFeedback());

    if (sendSuccess) {
      console.log(`[${requestId}] ✓ Response sent successfully`);
      storage.set(conversationKey, messages);
    } else {
      console.error(`[${requestId}] ✗ Failed to send response to user`);
    }

    console.log(`[${requestId}] ===== REQUEST COMPLETE =====`);
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    const errorStack = error instanceof Error ? error.stack : undefined;

    console.error(`[${requestId}] ✗✗✗ ERROR ✗✗✗`);
    console.error(`[${requestId}] Error type: ${error instanceof Error ? error.constructor.name : typeof error}`);
    console.error(`[${requestId}] Error message: ${errorMessage}`);
    if (errorStack) {
      console.error(`[${requestId}] Stack trace:`, errorStack);
    }
    console.error(`[${requestId}] User text: "${userText}"`);
    console.error(`[${requestId}] Full error:`, error);

    let userFriendlyMessage = "Sorry, I ran into an issue processing your request.";

    // Provide more specific error messages for common issues
    if (errorMessage.includes("timed out")) {
      userFriendlyMessage = "Sorry, the request took too long to process. The Brief data might be very large. Please try a more specific query or contact your admin.";
    } else if (errorMessage.includes("token") || errorMessage.includes("context_length")) {
      userFriendlyMessage = "Sorry, the Brief data is too large to process. Please contact your admin to optimize the data file.";
    } else if (errorMessage.includes("ECONNRESET") || errorMessage.includes("ETIMEDOUT")) {
      userFriendlyMessage = "Sorry, there was a network issue. Please try again in a moment.";
    }

    await safeSend(
      send,
      new MessageActivity(`${userFriendlyMessage}\n\nError ID: ${requestId}`)
    );

    console.error(`[${requestId}] ===== REQUEST FAILED =====`);
  }
});

app.on("message.submit.feedback", async ({ activity }) => {
  console.log("Feedback received:", JSON.stringify(activity.value));
});

export default app;
