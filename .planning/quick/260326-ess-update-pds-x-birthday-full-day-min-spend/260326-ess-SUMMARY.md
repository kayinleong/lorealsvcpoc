---
quick_task: 260326-ess
title: Update PDS x Birthday Full Day Min Spend to RM1000
date: 2026-03-26
commit: 2e42012
duration: ~3 minutes
tags: [xlsx, brief, min-spend, promo]
---

# Quick Task 260326-ess: Update PDS x Birthday Full Day Min Spend Summary

**One-liner:** Updated min spend for PDS x Birthday Full Day campaign rows in Promo-Flexi Combo-GWP sheet from 120/150 to 1000.

## What Was Done

Updated three cells in `.data/Brief.xlsx` sheet `Promo-Flexi Combo-GWP`:

| Cell | Campaign (B col) | Old Value | New Value |
|------|-----------------|-----------|-----------|
| M21  | PDS x Birthday Full Day | 120 | 1000 |
| M22  | PDS x Birthday Full Day | 120 | 1000 |
| M23  | PDS x Birthday Full Day | 150 | 1000 |

## Approach

Used openpyxl with `data_only=False` to preserve formulas in all other cells. Only the three specified min spend cells were modified.

## Verification

- Pre-change: M21=120, M22=120, M23=150 (confirmed)
- Post-change: M21=1000, M22=1000, M23=1000 (confirmed by re-loading workbook)
- All formula cells in columns C-H (VLOOKUP/INDEX/MATCH) preserved unchanged
- Promo-Voucher sheet row 16 (already 1000) not touched

## Deviations

None - executed exactly as specified.

## Commit

- `2e42012`: fix(260326-ess): update PDS x Birthday Full Day min spend to 1000 in Promo-Flexi Combo-GWP

## Self-Check: PASSED

- .data/Brief.xlsx modified and committed: FOUND (2e42012)
- M21=1000, M22=1000, M23=1000: VERIFIED
