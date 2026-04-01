# Quick Task 260401-lq1 — Summary

**Task:** Fix prompt: Update Campaign BAU Full Day all tier Min Spend to RM1k
**Date:** 2026-04-01
**Commit:** 4ad6ea8

## What was done

Updated `src/app/instructions.txt` with two changes:

1. **VOUCHER IDENTIFIER section** — added rule: if user says "all", "all tiers",
   "all vouchers", "semua tier", or omits the tier → omit `voucher_name` from the
   action block entirely so the backend updates all matching rows.

2. **PERFORMING UPDATES action template** — added IMPORTANT note: when targeting
   all tiers, omit `voucher_name` from the action block (do not include the key at
   all). The system will then update ALL rows matching the campaign name.

## Root cause

`applyUpdateSurgical` in `app.ts` already handles the "all rows" case correctly:
when `action.voucher_name` is falsy, `voucherMatch` is always true, so every row
matching `campaign_name` is updated. The instructions never told the LLM to omit
`voucher_name` for all-tier requests, so it always filled it in — preventing the
match-all behavior.

## Files changed

- `src/app/instructions.txt` — 3 lines added
