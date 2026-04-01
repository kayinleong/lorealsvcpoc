---
phase: quick
plan: 260401-nfs
type: execute
wave: 1
depends_on: []
files_modified:
  - .data/Brief.xlsx
autonomous: true
requirements: []

must_haves:
  truths:
    - "EAN 9557534776907 exists in Price Change - non Flash sheet (confirmed row 74 col B)"
    - "SKU Group Library has a new 'BPC' group entry with EAN 9557534776907"
    - "Voucher 'PDS CBMO' has Inclusion Group 'BPC' in the EANS tab"
    - "All existing EANS rows (rows 7-64) are intact and unmodified"
  artifacts:
    - path: ".data/Brief.xlsx"
      provides: "Updated Brief with BPC group and PDS CBMO inclusion row"
  key_links:
    - from: "EANS row (new)"
      to: "Col B = 'PDS CBMO', Col C = 'BPC'"
      via: "Voucher Name + Inclusion Group"
    - from: "EANS row (new)"
      to: "Col G = 'BPC', Col H = 9557534776907"
      via: "SKU Group Library Group Name + EAN"
---

<objective>
Add L'Oreal Paris Elseve Extraordinary Oil Hair Treatment Set - Pink (100ml x 2) (EAN: 9557534776907) to a new SKU group "BPC" in the EANS tab, and add voucher "PDS CBMO" with inclusion group "BPC".

Purpose: Set up the BPC inclusion group for the PDS CBMO voucher so the product is included in the correct promotional scope.
Output: Updated .data/Brief.xlsx with one new row in Included-Excluded EANs- Voucher tab.
</objective>

<context>
@.planning/STATE.md
@.data/Brief.xlsx
</context>

<tasks>

<task type="auto">
  <name>Task 1: Verify EAN and add BPC group + PDS CBMO inclusion row to EANS tab</name>
  <files>.data/Brief.xlsx</files>
  <action>
Write a Python script using openpyxl to perform all three steps:

**Step 1 — Verify EAN in Price Change - non Flash:**
Open .data/Brief.xlsx with data_only=True. In sheet "Price Change - non Flash", scan all rows for a cell whose value equals 9557534776907. Confirm it exists and that the product name in the same row contains "Elseve Extraordinary Oil" or similar. If not found, abort with an error message.

**Step 2 — Add new row to Included-Excluded EANs- Voucher tab:**
Reopen .data/Brief.xlsx with data_only=False (to preserve formulas). Access sheet "Included-Excluded EANs- Voucher".

Current last data row is 64. Append a new row at row 65 with:
- Col B (index 2, column 2): "PDS CBMO"    ← Voucher Name
- Col C (index 3, column 3): "BPC"          ← Inclusion Group
- Col D (index 4, column 4): ""             ← Exclusion Group (leave empty)
- Col G (index 7, column 7): "BPC"          ← Group Name (SKU Group Library)
- Col H (index 8, column 8): 9557534776907  ← SKU EAN (as integer)

All other columns in the new row: leave empty (do not write). Do NOT copy styles from adjacent rows — openpyxl will write plain values.

Use ws.cell(row=65, column=N).value = X to set each cell individually.

**Step 3 — Save:**
Save the workbook back to .data/Brief.xlsx.

Do NOT modify any other rows, sheets, or columns.
  </action>
  <verify>
    <automated>python3 -c "
import openpyxl
wb = openpyxl.load_workbook('.data/Brief.xlsx', data_only=True)

# Verify EAN in Price Change - non Flash
ws_nf = wb['Price Change - non Flash']
found_ean = False
for row in ws_nf.iter_rows(values_only=True):
    if any(c == 9557534776907 for c in row if c is not None):
        found_ean = True
        break
assert found_ean, 'EAN 9557534776907 not found in Price Change - non Flash'

# Verify EANS row 65
ws = wb['Included-Excluded EANs- Voucher']
assert ws.cell(row=65, column=2).value == 'PDS CBMO', f'B65={ws.cell(row=65,column=2).value}'
assert ws.cell(row=65, column=3).value == 'BPC', f'C65={ws.cell(row=65,column=3).value}'
assert ws.cell(row=65, column=7).value == 'BPC', f'G65={ws.cell(row=65,column=7).value}'
assert ws.cell(row=65, column=8).value == 9557534776907, f'H65={ws.cell(row=65,column=8).value}'

# Verify row 64 intact
assert ws.cell(row=64, column=7).value == 'Exclude CBMO', f'G64={ws.cell(row=64,column=7).value}'

print('All assertions passed: EAN verified, row 65 correct, row 64 intact')
"</automated>
  </verify>
  <done>Row 65 of Included-Excluded EANs- Voucher tab has Voucher Name="PDS CBMO", Inclusion Group="BPC", Group Name="BPC", SKU EAN=9557534776907. Existing rows 7-64 are unchanged.</done>
</task>

</tasks>

<verification>
Run the automated verify command above. All four assertions must pass.
</verification>

<success_criteria>
- EAN 9557534776907 confirmed present in "Price Change - non Flash" sheet
- EANS tab row 65: B="PDS CBMO", C="BPC", G="BPC", H=9557534776907
- All existing EANS rows (7-64) intact
- No other sheets or rows modified
</success_criteria>

<output>
After completion, create `.planning/quick/260401-nfs-add-l-oreal-paris-elseve-extraordinary-o/260401-nfs-SUMMARY.md`
</output>
