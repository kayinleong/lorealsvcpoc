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
import ExcelJS from "exceljs";
import config from "../config";

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

const BRIEF_PATH = path.join(process.cwd(), ".data", "Brief.xlsx");
const PROMO_SHEET = "Promo-Voucher";
const EANS_SHEET = "Included-Excluded EANs- Voucher";

// Row indices (0-based) inside each sheet
const PROMO_HEADER_ROW = 7; // Row 8 in Excel — column names
const PROMO_DATA_ROW = 9;   // Row 10 in Excel — first actual data row (row 9 is a description row)
const EANS_HEADER_ROW = 4;  // Row 5 in Excel — column names
const EANS_DATA_ROW = 6;    // Row 7 in Excel — first actual data row (row 6 is a description row)

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
  model: config.azureOpenAIDeploymentName,
  apiKey: config.azureOpenAIKey,
  endpoint: config.azureOpenAIEndpoint,
  apiVersion: "2024-10-21",
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

/** Extract plain text from an ExcelJS cell value (handles rich text, formulas, dates). */
function excelCellText(value: ExcelJS.CellValue): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "string") return value;
  if (typeof value === "number" || typeof value === "boolean") return String(value);
  if (value instanceof Date) return value.toISOString();
  const obj = value as unknown as Record<string, unknown>;
  if ("richText" in obj) {
    return (obj.richText as Array<{ text: string }>).map((r) => r.text).join("");
  }
  if ("result" in obj) return obj.result !== undefined ? String(obj.result) : "";
  return String(value);
}

/**
 * Read Brief.xlsx and return a tab-separated text snapshot of the two
 * relevant sheets so the LLM always has the latest data in context.
 */
function readBriefContext(): string {
  if (!fs.existsSync(BRIEF_PATH)) {
    return "[BRIEF DATA]\nError: Brief.xlsx not found at .data/Brief.xlsx\n[/BRIEF DATA]";
  }

  const wb = xlsx.readFile(BRIEF_PATH);

  const formatSheet = (
    sheetName: string,
    headerRow: number,
    dataRow: number,
    maxCols: number,
  ): string => {
    const ws = wb.Sheets[sheetName];
    if (!ws) return `### ${sheetName}\n(sheet not found)\n`;

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
    return [
      `### ${sheetName} (${dataRows.length} rows)`,
      headerLine,
      ...lines,
    ].join("\n");
  };

  const promoSection = formatSheet(PROMO_SHEET, PROMO_HEADER_ROW, PROMO_DATA_ROW, 25);
  const eansSection = formatSheet(EANS_SHEET, EANS_HEADER_ROW, EANS_DATA_ROW, 11);

  return ["[BRIEF DATA]", promoSection, "", eansSection, "[/BRIEF DATA]"].join(
    "\n",
  );
}

// ---------------------------------------------------------------------------
// Action parsing and application
// ---------------------------------------------------------------------------

interface UpdateAction {
  operation: string;
  tab: string;
  campaign_name?: string;
  voucher_name?: string;
  field: string;
  value: string | number;
}

/** Extract the JSON action block from LLM response text. */
function parseAction(text: string): UpdateAction | null {
  const match = text.match(/<action>([\s\S]*?)<\/action>/i);
  if (!match) return null;
  try {
    return JSON.parse(match[1].trim()) as UpdateAction;
  } catch {
    return null;
  }
}

/** Remove <action>...</action> block from display text. */
function stripAction(text: string): string {
  return text.replace(/<action>[\s\S]*?<\/action>/gi, "").trim();
}

/**
 * Apply a validated update action directly to Brief.xlsx using ExcelJS,
 * which faithfully preserves all formatting, colors, and calculations.
 */
async function applyUpdate(action: UpdateAction): Promise<{
  success: boolean;
  oldValue?: unknown;
  error?: string;
}> {
  if (!fs.existsSync(BRIEF_PATH)) {
    return { success: false, error: "Brief.xlsx not found." };
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(BRIEF_PATH);

  const worksheet = workbook.getWorksheet(action.tab);
  if (!worksheet) {
    return { success: false, error: `Sheet "${action.tab}" not found.` };
  }

  // ExcelJS uses 1-based row numbers; our constants are 0-based row indices.
  const headerRowNum =
    (action.tab === PROMO_SHEET ? PROMO_HEADER_ROW : EANS_HEADER_ROW) + 1;
  const dataRowNum =
    (action.tab === PROMO_SHEET ? PROMO_DATA_ROW : EANS_DATA_ROW) + 1;

  // Find the target column by matching cleaned header text.
  const headerRow = worksheet.getRow(headerRowNum);
  let fieldColNum = -1;
  headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    if (cleanHeader(excelCellText(cell.value)) === action.field) {
      fieldColNum = colNumber;
    }
  });

  if (fieldColNum === -1) {
    return {
      success: false,
      error: `Field "${action.field}" not found in "${action.tab}" headers.`,
    };
  }

  // Column numbers are 1-based in ExcelJS.
  //   Promo-Voucher  — Campaign Name: col B (2), Voucher Name: col L (12)
  //   EANs-Voucher   — Voucher Name: col B (2) (no Campaign Name column)
  const campaignColNum = 2;
  const voucherColNum = action.tab === PROMO_SHEET ? 12 : 2;

  let targetRowNum = -1;
  for (let r = dataRowNum; r <= worksheet.rowCount; r++) {
    const row = worksheet.getRow(r);
    const campaignVal = excelCellText(row.getCell(campaignColNum).value);
    const voucherVal = excelCellText(row.getCell(voucherColNum).value);

    if (voucherVal.trim() === "") continue;

    const campaignMatch =
      action.tab !== PROMO_SHEET ||
      !action.campaign_name ||
      campaignVal.toLowerCase().includes(action.campaign_name.toLowerCase());

    const voucherMatch =
      !action.voucher_name ||
      voucherVal.toLowerCase().includes(action.voucher_name.toLowerCase());

    if (campaignMatch && voucherMatch) {
      targetRowNum = r;
      break;
    }
  }

  if (targetRowNum === -1) {
    return {
      success: false,
      error: `Record not found — campaign: "${action.campaign_name}", voucher: "${action.voucher_name}".`,
    };
  }

  const targetCell = worksheet.getRow(targetRowNum).getCell(fieldColNum);
  const oldValue = excelCellText(targetCell.value);

  // ExcelJS preserves all cell styles, fills, borders, and formulas in the
  // rest of the workbook — only this cell's value is changed.
  targetCell.value = action.value;

  await workbook.xlsx.writeFile(BRIEF_PATH);
  return { success: true, oldValue };
}

// ---------------------------------------------------------------------------
// Message handler
// ---------------------------------------------------------------------------

app.on("message", async ({ send, activity }) => {
  const conversationKey = `${activity.conversation.id}/${activity.from.id}`;
  const messages: Message[] = storage.get(conversationKey) ?? [];

  // Prepend the latest Brief.xlsx snapshot to every user message so the LLM
  // always works from the current data without relying on stale history.
  const briefContext = readBriefContext();
  const userText = activity.text ?? "";
  const userMessage = `${briefContext}\n\n${userText}`;

  try {
    const prompt = new ChatPrompt({ messages, instructions, model });
    const response = await prompt.send(userMessage);
    const rawText: string = response.content;

    // Strip Brief data from stored conversation history to keep it lean.
    // The last entry added by prompt.send() is the user message — replace
    // its content with just the original user text.
    const lastUserIdx = [...messages]
      .map((m) => m.role)
      .lastIndexOf("user");
    if (lastUserIdx !== -1) {
      (messages[lastUserIdx] as Message).content = userText;
    }

    // Check for an update action from the LLM
    const action = parseAction(rawText);
    let displayText = stripAction(rawText);

    if (action?.operation === "update") {
      const result = await applyUpdate(action);
      if (result.success) {
        displayText +=
          "\n\nBrief is updated: [Download the latest Brief.xlsx here](.data/Brief.xlsx)";
      } else {
        console.error("applyUpdate failed:", result.error);
        displayText +=
          "\n\n⚠️ The brief could not be saved automatically. Please contact your admin.";
      }
    }

    await send(
      new MessageActivity(displayText).addAiGenerated().addFeedback(),
    );
    storage.set(conversationKey, messages);
  } catch (error) {
    console.error("Brief agent error:", error);
    await send(
      "Sorry, I ran into an issue processing your request. Please try again or rephrase your message.",
    );
  }
});

app.on("message.submit.feedback", async ({ activity }) => {
  console.log("Feedback received:", JSON.stringify(activity.value));
});

export default app;
