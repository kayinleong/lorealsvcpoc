# L'Oreal Brief Management Agent

## What This Is

A Microsoft Teams AI bot that enables L'Oreal business users to read and update the promotional brief (Brief.xlsx) using natural language commands. Users type instructions like "Update BAU Full Day Tier 2 Min Spend to RM1000" and the agent parses the intent, validates business rules, applies the change to Excel, and returns a confirmation with a download link.

## Core Value

Business users can update the promotional brief directly from Teams using plain language — no Excel knowledge, no file uploads, no manual cell-hunting.

## Requirements

### Validated

- Read voucher attributes from Brief.xlsx (Promo-Voucher and Included-Excluded EANs tabs)
- Update voucher attributes via natural language (field value changes)
- Validate all business rules before applying any change (Min Spend, Discount Type consistency, Amount Capped, Voucher Budget/Quantity exclusivity, Live Date, Limit per Customer, Voucher Code rules, Storewide/SKU rules)
- Return structured confirmation message after successful update
- Handle ambiguous requests (multiple matches) by listing options and asking for confirmation
- Preserve Excel formatting, colors, and formulas on write (ExcelJS)
- Normalize input formats: strip RM/$, accept dd/mm/yyyy and natural language dates, accept percentages as numbers

### Active

- [ ] Add new voucher row to Promo-Voucher tab
- [ ] Delete/remove a voucher row from Promo-Voucher tab
- [ ] Read/write Brief.xlsx from SharePoint instead of local file system (Microsoft Graph API)
- [ ] Return working SharePoint download link after updates (instead of local path)
- [ ] Audit trail: structured change confirmation message in Teams chat that documents who changed what and when

### Out of Scope

- Multi-user permission tiers (read-only vs read-write) — not requested
- Real-time collaboration / conflict detection — single-user bot flow
- Modifying any other Excel tabs beyond Promo-Voucher and Included-Excluded EANs — per instructions
- Modifying auto-calculated columns (Voucher Status, Shop SKU ID, Product ID, Variation ID) — enforced by instructions

## Context

- **Platform:** Microsoft Teams via @microsoft/teams.apps + @microsoft/teams.ai SDK
- **AI Backend:** Azure OpenAI (deployment name, key, endpoint from env vars)
- **Auth:** Managed Identity (`ManagedIdentityCredential`) — already in place, suitable for Graph API calls
- **Excel I/O:** ExcelJS (write, preserves formatting) + xlsx (read, fast parsing for context snapshot)
- **Brief.xlsx structure:** Headers on row 8, data from row 10 (Promo-Voucher); headers row 5, data from row 7 (Included-Excluded EANs)
- **Current file location:** `.data/Brief.xlsx` (local) — to be migrated to SharePoint
- **SharePoint site/library:** Not yet determined — needs to be configured during implementation
- **Audit approach:** Teams conversation history serves as the log; bot will output structured change summaries per update

## Constraints

- **Tech stack:** TypeScript + Node 20/22 — no stack changes
- **Auth:** Must use Managed Identity for Graph API (no service principal secrets in code)
- **Excel:** Must preserve all cell formatting and formulas on write (ExcelJS required)
- **SharePoint:** Site URL and document library to be provided before SharePoint phase begins

## Key Decisions

| Decision | Rationale | Outcome |
|----------|-----------|---------|
| Use ExcelJS for writes (not xlsx) | xlsx library doesn't preserve formatting; ExcelJS does | — Pending |
| Audit via Teams chat messages | Zero extra infrastructure; Teams already records history | — Pending |
| Managed Identity for Graph API | Already used for bot auth; no new secrets needed | — Pending |
| Add/Delete before SharePoint migration | Simpler to test locally first, then swap file source | — Pending |

---
*Last updated: 2026-03-20 after initialization*
