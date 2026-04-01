---
phase: quick
plan: 260401-oyg
type: execute
wave: 1
depends_on: []
files_modified:
  - .data/Brief.xlsx
autonomous: true
requirements: []

must_haves:
  truths:
    - "EANS tab row 65 has Product ID (col J) = 3854794414"
    - "EANS tab row 65 has Voucher Name (col B) = 'PDS CBMO'"
    - "EANS tab row 65 has Inclusion Group (col C) = 'BPC'"
    - "No new row is added — only the existing row 65 is updated"
    - "All other rows in EANS tab are intact and unmodified"
  artifacts:
    - path: ".data/Brief.xlsx"
      provides: "Updated Brief with row 65 fully populated in EANS tab"
  key_links:
    - from: "Price Change - non Flash row 74"
      to: "EANS row 65 col J"
      via: "Product ID 3854794414 looked up by EAN 9557534776907"
---

<objective>
Patch EANS tab row 65 in Brief.xlsx to fill in the three missing cells: Product ID (col J = 3854794414), Voucher Name (col B = "PDS CBMO"), and Inclusion Group (col C = "BPC").

Purpose: Row 65 was previously added for EAN 9557534776907 (L'Oreal Paris Elseve Extraordinary Oil Hair Treatment Set - Pink 100ml x 2) but is incomplete — the Product ID and left-side voucher columns were not written. This patch completes the row.
Output: .data/Brief.xlsx with row 65 fully populated, no other changes.
</objective>

<context>
@.planning/STATE.md
@.data/Brief.xlsx
</context>

<tasks>

<task type="auto">
  <name>Task 1: Patch EANS tab row 65 with Product ID, Voucher Name, and Inclusion Group</name>
  <files>.data/Brief.xlsx</files>
  <action>
Write a Python script using openpyxl to patch the existing row 65 in the EANS tab:

**Step 1 — Verify preconditions:**
Open .data/Brief.xlsx with data_only=True. In sheet "Price Change - non Flash", confirm row 74 col B (column 2) = 9557534776907 and col D (column 4) = 3854794414. If either value is absent, abort with a clear error.

In sheet "Included-Excluded EANs- Voucher", confirm row 65 col H (column 8) = 9557534776907 and col G (column 7) = "BPC". If not matching, abort — this is not the expected row.

**Step 2 — Patch row 65:**
Reopen .data/Brief.xlsx with data_only=False (to preserve formulas). Access sheet "Included-Excluded EANs- Voucher".

Use ws.cell(row=65, column=N).value = X to set only these three cells:
- col B (column 2): "PDS CBMO"   ← Voucher Name (if currently empty)
- col C (column 3): "BPC"        ← Inclusion Group (if currently empty)
- col J (column 10): 3854794414  ← Product ID (integer)

Do NOT write to any other column in row 65. Do NOT modify any other row.

**Step 3 — Save:**
Save the workbook back to .data/Brief.xlsx.
  </action>
  <verify>
    <automated>python3 -c "
import openpyxl
wb = openpyxl.load_workbook('.data/Brief.xlsx', data_only=True)

# Verify source in Price Change - non Flash
ws_nf = wb['Price Change - non Flash']
assert ws_nf.cell(row=74, column=2).value == 9557534776907, f'B74={ws_nf.cell(row=74,column=2).value}'
assert ws_nf.cell(row=74, column=4).value == 3854794414, f'D74={ws_nf.cell(row=74,column=4).value}'

# Verify patched row 65 in EANS tab
ws = wb['Included-Excluded EANs- Voucher']
assert ws.cell(row=65, column=2).value == 'PDS CBMO', f'B65={ws.cell(row=65,column=2).value}'
assert ws.cell(row=65, column=3).value == 'BPC', f'C65={ws.cell(row=65,column=3).value}'
assert ws.cell(row=65, column=7).value == 'BPC', f'G65={ws.cell(row=65,column=7).value}'
assert ws.cell(row=65, column=8).value == 9557534776907, f'H65={ws.cell(row=65,column=8).value}'
assert ws.cell(row=65, column=10).value == 3854794414, f'J65={ws.cell(row=65,column=10).value}'

# Verify surrounding rows intact
assert ws.cell(row=64, column=7).value == 'Exclude CBMO', f'G64={ws.cell(row=64,column=7).value}'

print('All assertions passed: row 65 fully populated, surrounding rows intact')
"</automated>
  </verify>
  <done>EANS tab row 65 has B="PDS CBMO", C="BPC", G="BPC", H=9557534776907, J=3854794414. No other rows or sheets modified.</done>
</task>

</tasks>

<verification>
Run the automated verify command above. All assertions must pass — confirming the Product ID is set, the voucher columns are filled, and no adjacent rows were disturbed.
</verification>

<success_criteria>
- EANS tab row 65: col B = "PDS CBMO", col C = "BPC", col G = "BPC", col H = 9557534776907, col J = 3854794414
- No new row added — row count in EANS tab unchanged
- All existing EANS rows (7-64, 66+) intact
- No other sheets modified
</success_criteria>

<output>
After completion, create `.planning/quick/260401-oyg-add-l-oreal-paris-elseve-extraordinary-o/260401-oyg-SUMMARY.md`
</output>
