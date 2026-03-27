---
phase: quick
plan: 260327-ect
subsystem: app
tags: [bugfix, regex, xlsx, instructions, eans]
key-files:
  modified:
    - src/app/app.ts
    - src/app/instructions.txt
    - .data/Brief.xlsx
decisions:
  - Used lazy regex quantifier ([^>]*?) to fix self-closing XML cell matching
  - One-time repair script approach (not startup repair) to write data into Brief.xlsx
  - Soft-warning model for EANS adds with unknown voucher names — block is wrong UX
metrics:
  duration: ~10min
  completed: "2026-03-27"
  tasks: 3
  files: 3
---

# Quick Task 260327-ect: Fix EANS Tab Add Bugs & writeCellInRowXml Summary

**One-liner:** Fixed catastrophic regex backtracking in `writeCellInRowXml`, seeded PDS/Exclude CBMO into Brief.xlsx EANS row 11, and added EANS add workflow to LLM instructions.

## Tasks Completed

| # | Task | Commit | Files |
|---|------|--------|-------|
| 1 | Fix writeCellInRowXml regex — lazy ([^>]*?) | b7de81d | src/app/app.ts |
| 2 | Seed PDS / Exclude CBMO into Brief.xlsx EANS row 11 | 3dccfc5 | .data/Brief.xlsx |
| 3 | Add EANS add instructions + update validation rule | 35fc405 | src/app/instructions.txt |

## Changes Detail

### Task 1 — writeCellInRowXml regex fix (src/app/app.ts)

Changed `([^>]*)` to `([^>]*?)` in the cell pattern regex. The greedy `[^>]*` was causing catastrophic backtracking on self-closing cells like `<c r="B11" s="1287"/>`, where the engine would attempt to consume the entire remaining row XML before failing. The lazy `[^>]*?` stops at the first `>` it encounters, which is correct.

### Task 2 — Brief.xlsx EANS row 11 seeded

Row 11 in `xl/worksheets/sheet25.xml` had:
- `<c r="B11" s="1287"/>` (empty self-closing)
- `<c r="C11" s="1288"/>` (empty self-closing)

A one-time Node.js repair script (`repair-eans-row11.js`, deleted after use) used AdmZip to read the ZIP, apply the lazy regex fix to write "PDS" into B11 and "Exclude CBMO" into C11, and write the ZIP back. Both cells now use `t="inlineStr"` format and preserve their original style attributes.

### Task 3 — instructions.txt EANS add workflow

Added a new section `## PERFORMING ADDS — Included-Excluded EANs- Voucher Tab` immediately before `## RESPONSE FORMAT`. The section:
- Instructs the LLM to use the user's exact voucher name verbatim (no substitution from existing data)
- Distinguishes Inclusion Group vs Exclusion Group (mutually exclusive per row)
- Uses a soft-warning model: if voucher name not found in Promo-Voucher, warn but proceed
- Emits the action block immediately without asking for confirmation

Also updated EANS validation rule 1 to split the must-match requirement into:
- UPDATE/READ: voucher name must exist in Promo-Voucher
- ADD: warn if not found, but do not block

## Deviations from Plan

**[Rule 2 - Missing cleanup]** Test files (test_regex.js, test_regex2.js, test_regex3.js, test_xlsx_read.js) were not present in the worktree and had no git history — they likely existed only in the original working directory context. No cleanup was needed.

## Self-Check

- src/app/app.ts modified: FOUND
- .data/Brief.xlsx modified: FOUND
- src/app/instructions.txt modified: FOUND
- Commit b7de81d: FOUND
- Commit 3dccfc5: FOUND
- Commit 35fc405: FOUND

## Self-Check: PASSED
