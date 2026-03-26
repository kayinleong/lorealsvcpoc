---
phase: quick
plan: 260326-ess
type: execute
wave: 1
depends_on: []
files_modified: [".data/Brief.xlsx"]
autonomous: true
requirements: ["Update PDS x Birthday Full Day Min Spend to RM1000"]
must_haves:
  truths:
    - "Promo-Flexi Combo-GWP rows 21-23 column M all show 1000"
    - "Promo-Voucher row 16 remains unchanged at 1000"
    - "All other data, formulas, and formatting preserved"
  artifacts:
    - path: ".data/Brief.xlsx"
      provides: "Updated promotional brief"
  key_links: []
---

<objective>
Update the Min Spend values for "PDS x Birthday Full Day" campaign in the Promo-Flexi Combo-GWP sheet of Brief.xlsx from 120/150 to 1000.

Purpose: Correct the minimum spend threshold for the PDS x Birthday Full Day campaign.
Output: Updated .data/Brief.xlsx
</objective>

<context>
@.data/Brief.xlsx
</context>

<tasks>

<task type="auto">
  <name>Task 1: Update Min Spend values in Promo-Flexi Combo-GWP sheet</name>
  <files>.data/Brief.xlsx</files>
  <action>
Write a Python script using openpyxl to:
1. Open .data/Brief.xlsx with data_only=False (preserve formulas)
2. Access the "Promo-Flexi Combo-GWP" sheet
3. Update cell M21 to 1000 (currently 120)
4. Update cell M22 to 1000 (currently 120)
5. Update cell M23 to 1000 (currently 150)
6. Verify Promo-Voucher sheet row 16 column N is already 1000 (read-only check, do not modify)
7. Save the workbook

Do NOT touch any other cells, sheets, or rows. Use openpyxl to preserve all existing formulas and formatting.
  </action>
  <verify>
    <automated>python -c "import openpyxl; wb=openpyxl.load_workbook('.data/Brief.xlsx'); ws=wb['Promo-Flexi Combo-GWP']; assert ws['M21'].value==1000, f'M21={ws[\"M21\"].value}'; assert ws['M22'].value==1000, f'M22={ws[\"M22\"].value}'; assert ws['M23'].value==1000, f'M23={ws[\"M23\"].value}'; pv=wb['Promo-Voucher']; assert pv['N16'].value==1000, f'N16={pv[\"N16\"].value}'; print('All values correct')"</automated>
  </verify>
  <done>Promo-Flexi Combo-GWP rows 21-23 column M all equal 1000. Promo-Voucher row 16 column N unchanged at 1000.</done>
</task>

</tasks>

<verification>
Run the automated verify command to confirm all three cells updated and Promo-Voucher unchanged.
</verification>

<success_criteria>
- Brief.xlsx Promo-Flexi Combo-GWP M21=1000, M22=1000, M23=1000
- Brief.xlsx Promo-Voucher N16=1000 (unchanged)
- No other data modified
</success_criteria>

<output>
After completion, create `.planning/quick/260326-ess-update-pds-x-birthday-full-day-min-spend/260326-ess-SUMMARY.md`
</output>
