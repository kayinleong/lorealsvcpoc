---
phase: quick
plan: 260327-ect
type: execute
wave: 1
depends_on: []
files_modified:
  - src/app/app.ts
  - src/app/instructions.txt
autonomous: true
requirements: []

must_haves:
  truths:
    - "writeCellInRowXml correctly rewrites self-closing cells without consuming adjacent cells"
    - "LLM uses the exact user-supplied voucher name for EANS tab adds — no partial matching"
    - "LLM does not hard-block EANS adds when voucher name is absent from Promo-Voucher tab"
    - "Brief.xlsx EANS tab contains a row for 'PDS' with Inclusion Group 'Exclude CBMO'"
  artifacts:
    - path: "src/app/app.ts"
      provides: "Fixed writeCellInRowXml regex (lazy quantifier)"
    - path: "src/app/instructions.txt"
      provides: "EANS add rules: literal voucher name, soft validation warning"
  key_links:
    - from: "src/app/app.ts writeCellInRowXml"
      to: "self-closing cell XML"
      via: "lazy ([^>]*?) regex"
      pattern: "\\[\\^>\\]\\*\\?"
---

<objective>
Fix three bugs affecting EANS tab add operations, then manually repair the data row that was lost due to these bugs.

Purpose: EANS tab adds were corrupting XML (regex greediness), hallucinating voucher names (LLM matched partial names), and blocking valid adds (hard validation rule). A "PDS" row is also missing from Brief.xlsx due to these failures.
Output: Fixed app.ts regex, updated instructions.txt rules, and "PDS / Exclude CBMO" row written to Brief.xlsx EANS tab.
</objective>

<execution_context>
@$HOME/.claude/get-shit-done/workflows/execute-plan.md
@$HOME/.claude/get-shit-done/templates/summary.md
</execution_context>

<context>
@.planning/STATE.md
</context>

<tasks>

<task type="auto">
  <name>Task 1: Fix writeCellInRowXml greedy regex and add missing EANS row</name>
  <files>src/app/app.ts</files>
  <action>
Two changes in src/app/app.ts:

**Change 1 — Fix regex (line ~597):**
In function `writeCellInRowXml`, change the cell pattern from:
```typescript
const cellPattern = new RegExp(
  `<c\\s+r="${cellRef}"([^>]*)(?:/>|>[\\s\\S]*?</c>)`
);
```
to:
```typescript
const cellPattern = new RegExp(
  `<c\\s+r="${cellRef}"([^>]*?)(?:/>|>[\\s\\S]*?</c>)`
);
```
The only change is `([^>]*)` → `([^>]*?)` (lazy quantifier). This prevents the greedy capture from consuming the `/` of a self-closing cell and then catastrophically eating through the next `</c>`.

**Change 2 — Add missing "PDS" row to Brief.xlsx:**
After the regex fix, write a one-off repair function at the bottom of the file (or inline in main) that:
1. Opens Brief.xlsx using AdmZip (same pattern as `applyAdd`)
2. Resolves the EANS worksheet XML path (same zip/workbook.xml.rels lookup)
3. Reads the worksheet XML as text
4. Finds the row after the last data row in the EANS tab. Current last data Excel row is 10 ("3.3 Tier 1"). Target: row 11.
5. Finds the existing `<row r="11" ...>...</row>` element using the row regex pattern already in `applyAdd`
6. Calls `writeCellInRowXml` (with the now-fixed regex) to:
   - Write `B11` = `"PDS"` (string, style s="1287" — same style as B10 and B7-B9)
   - Write `C11` = `"Exclude CBMO"` (string, style s="1288" — same style as C10)
7. Replaces the row XML in worksheetXML and writes back via `zip.updateFile` + `zip.writeZip`
8. Logs success

Run this repair function ONCE on startup (guarded by a check: if row 11 B column is already "PDS", skip). Alternatively, expose it as a separate exported function and call it from the top of the bot startup sequence.

The style attributes (s="1287" for Voucher Name, s="1288" for Inclusion Group) are confirmed from the existing rows in the EANS tab (B7-B9 use s="1287", C10 uses s="1288").

Do NOT use a separate script — add the repair inline or as a startup call in app.ts so it runs once when the server starts, then becomes a no-op on subsequent starts.
  </action>
  <verify>
    <automated>cd C:/Development/LorealSVCPOC && npx ts-node -e "
const AdmZip = require('adm-zip');
const zip = new AdmZip('./Brief.xlsx');
const xlsx = require('xlsx');
const wb = xlsx.readFile('./Brief.xlsx');
const ws = wb.Sheets['Included-Excluded EANs- Voucher'];
const rows = xlsx.utils.sheet_to_json(ws, { header: 1, defval: '' });
console.log('Row 11 (idx 10):', rows[10]);
console.log('Row 10 (idx 9):', rows[9]);
" 2>&1 | tail -5</automated>
  </verify>
  <done>Row 11 of EANS tab reads Voucher Name="PDS", Inclusion Group="Exclude CBMO". Row 10 and adjacent rows are intact. writeCellInRowXml regex contains `[^>]*?` (lazy).</done>
</task>

<task type="auto">
  <name>Task 2: Fix instructions.txt — EANS add voucher name and validation rules</name>
  <files>src/app/instructions.txt</files>
  <action>
Two additions to src/app/instructions.txt:

**Addition 1 — New EANS tab add section after the existing "PERFORMING ADDS" section:**
Append a new section titled `## PERFORMING ADDS — Included-Excluded EANs- Voucher Tab` immediately after the closing `---` of the current PERFORMING ADDS section (before RESPONSE FORMAT). Content:

```
## PERFORMING ADDS — Included-Excluded EANs- Voucher Tab

When the user requests to add a row to the EANS tab (Included-Excluded EANs- Voucher), follow these rules:

### Voucher Name — use EXACT literal value

The Voucher Name for an EANS add is provided LITERALLY by the user. Do NOT:
- Match or look up against existing Voucher Names in Promo-Voucher or EANS tabs
- Complete, substitute, or extend the name (e.g., if user says "PDS", write "PDS" — NOT "PDS CMBO")
- Correct apparent typos or abbreviations

Write exactly the string the user provided, character-for-character.

### Voucher Name validation — warning only (not a hard block)

The rule "Voucher Name must match a Voucher Name in Promo-Voucher" is a SOFT WARNING for EANS add operations.

- If the voucher name does NOT appear in Promo-Voucher, include this note in your response:
  "Note: 'PDS' was not found in the Promo-Voucher tab. The row has been added — verify the voucher name is correct."
- Do NOT block or reject the add. Proceed and emit the action block.

### Inclusion Group vs Exclusion Group

- If the user says "inclusion group [X]" or "add inclusion [X]" → set Inclusion Group = X, leave Exclusion Group blank.
- If the user says "exclusion group [X]" or "add exclusion [X]" → set Exclusion Group = X, leave Inclusion Group blank.
- Never fill both on the same row.

### Action block format for EANS adds

```
<action>
{
  "operation": "add",
  "tab": "Included-Excluded EANs- Voucher",
  "rows": [
    {
      "Voucher Name": "<exact string from user>",
      "Inclusion Group": "<value or empty string>",
      "Exclusion Group": "<value or empty string>"
    }
  ]
}
</action>
```
```

**Addition 2 — Update the existing EANS Voucher Name validation rule (line ~199-201):**

Find the existing rule:
```
1. Voucher Name
   - Required when any Promo-Voucher row has Storewide = "Specific SKUs"
   - Must match a Voucher Name that exists in the Promo-Voucher tab
```

Change to:
```
1. Voucher Name
   - Required when any Promo-Voucher row has Storewide = "Specific SKUs"
   - For UPDATE/READ operations: Must match a Voucher Name that exists in the Promo-Voucher tab
   - For ADD operations: Use the EXACT string provided by the user. If the name is not found in Promo-Voucher, emit a warning but do NOT block the add.
```
  </action>
  <verify>
    <automated>grep -n "EXACT\|literal\|warning only\|soft warning\|NOT.*block\|Inclusion Group.*empty\|Exclusion Group.*empty" C:/Development/LorealSVCPOC/src/app/instructions.txt | head -20</automated>
  </verify>
  <done>instructions.txt contains the new EANS add section with literal-name rule, soft-warning validation, and correct action block format. The existing Voucher Name rule is updated to distinguish ADD from UPDATE/READ.</done>
</task>

</tasks>

<verification>
1. `writeCellInRowXml` in app.ts uses `([^>]*?)` (lazy) — confirm with: `grep -n "\[^>\]\*" src/app/app.ts`
2. EANS tab row 11 has Voucher Name="PDS" and Inclusion Group="Exclude CBMO"
3. EANS tab row 10 ("3.3 Tier 1") is still structurally intact
4. instructions.txt contains PERFORMING ADDS — EANS section with literal-name and soft-warning rules
5. TypeScript compiles without errors: `npx tsc --noEmit`
</verification>

<success_criteria>
- Self-closing cells in EANS tab XML are no longer corrupted by writeCellInRowXml
- Future LLM EANS adds use the exact user-supplied voucher name
- LLM does not block EANS adds when voucher is absent from Promo-Voucher (emits warning instead)
- Brief.xlsx EANS tab has row 11: PDS / Inclusion Group: Exclude CBMO
</success_criteria>

<output>
After completion, create `.planning/quick/260327-ect-fix-eans-tab-add-bugs-writecellinrowxml-/260327-ect-SUMMARY.md`
</output>
