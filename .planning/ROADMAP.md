# Roadmap: L'Oreal Brief Management Agent — Extension Milestone

## Overview

The bot already ships Read and Update. This milestone adds the remaining three capabilities: adding rows, deleting rows, a structured audit trail in Teams chat, and then migrating all file I/O from local disk to SharePoint via Microsoft Graph API. Phase 1 works entirely against the existing local file so every new operation is testable in isolation. Phase 2 swaps the file transport layer once the full operation set is proven.

## Phases

**Phase Numbering:**
- Integer phases (1, 2, 3): Planned milestone work
- Decimal phases (2.1, 2.2): Urgent insertions (marked with INSERTED)

Decimal phases appear between their surrounding integers in numeric order.

- [ ] **Phase 1: Row Operations + Audit Trail** - Add/delete voucher rows and emit structured audit output in Teams chat (local file)
- [ ] **Phase 2: SharePoint Integration** - Swap local file I/O for Microsoft Graph API reads and writes, return SharePoint web URL after every change

## Phase Details

### Phase 1: Row Operations + Audit Trail
**Goal**: Users can add and delete voucher rows via natural language, and every successful change produces a structured audit record in the Teams conversation
**Depends on**: Nothing (extends existing bot in place)
**Requirements**: ROW-01, ROW-02, ROW-03, ROW-04, AUD-01, AUD-02, AUD-03
**Success Criteria** (what must be TRUE):
  1. User can say "Add a new voucher row" with the required fields and the row appears in Promo-Voucher with all validation rules enforced
  2. User can say "Delete voucher [name]" and the row is removed from Promo-Voucher without a confirmation prompt
  3. After a successful add or delete, the bot posts a table in Teams showing Campaign, Voucher, Operation, Changed by, and Time
  4. After any successful update (existing or new), the bot posts a table showing Campaign, Voucher, Field, Before, After, Changed by (Teams display name), and UTC Time
  5. Attempted adds that violate validation rules (Min Spend, Discount Type, Budget/Quantity exclusivity, Voucher Code) are rejected with an explanation before any write occurs
**Plans**: TBD

Plans:
- [ ] 01-01: Implement add-row operation (action block, ExcelJS write, validation)
- [ ] 01-02: Implement delete-row operation (action block, ExcelJS splice)
- [ ] 01-03: Implement audit trail output (structured table, user identity, UTC timestamp)

### Phase 2: SharePoint Integration
**Goal**: The bot reads and writes Brief.xlsx from SharePoint via Microsoft Graph API, and returns the SharePoint web URL after every successful change
**Depends on**: Phase 1
**Requirements**: SPO-01, SPO-02, SPO-03, SPO-04
**Success Criteria** (what must be TRUE):
  1. The bot reads Brief.xlsx from SharePoint on every message (no local file required)
  2. After any successful add, update, or delete, the bot writes the modified workbook back to SharePoint preserving all formatting and formulas
  3. After any successful change, the bot returns a clickable SharePoint web URL so the user can open the file in Excel Online
  4. SharePoint site URL and document library path are set via environment variables with no hardcoded values in source
**Plans**: TBD

Plans:
- [ ] 02-01: Implement Graph API read (download file bytes via Managed Identity, feed to ExcelJS/xlsx)
- [ ] 02-02: Implement Graph API write (upload modified workbook bytes back to SharePoint)
- [ ] 02-03: Wire SharePoint web URL into post-change response; add env var configuration

## Progress

**Execution Order:**
Phases execute in numeric order: 1 → 2

| Phase | Plans Complete | Status | Completed |
|-------|----------------|--------|-----------|
| 1. Row Operations + Audit Trail | 0/3 | Not started | - |
| 2. SharePoint Integration | 0/3 | Not started | - |
