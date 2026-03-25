# Requirements: L'Oreal Brief Management Agent

**Defined:** 2026-03-20
**Core Value:** Business users can update the promotional brief directly from Teams using plain language

## v1 Requirements

Requirements for this extension milestone. Validated capabilities are already shipped.

### Row Operations

<!-- Promo-Voucher insert (add row + validation) is already shipped and validated. ROW-01 and ROW-02
     have been updated to reflect the remaining net-new EANs insert work that Phase 1 delivers. -->
- [ ] **ROW-01**: User can add a new Voucher Dashboard row to the EANs tab (Included-Excluded EANs- Voucher) via natural language, providing Voucher Name (must match an existing Promo-Voucher entry) and either Inclusions Group or Exclusion Group
- [ ] **ROW-02**: New EANs row addition enforces mutual-exclusivity validation (Inclusions Group XOR Exclusion Group) and Voucher Name cross-check against Promo-Voucher; invalid inserts are rejected with an explanation
- [ ] **ROW-03**: User can delete a voucher row from Promo-Voucher via natural language (immediate, no confirmation prompt)
- [ ] **ROW-04**: After successful add or delete, bot outputs a table-format audit summary (Campaign | Voucher | Operation | Changed by | Time)
- [ ] **ROW-05**: User can update any column on an existing EANs Voucher Dashboard row (Voucher Name, Inclusions Group, or Exclusion Group) via natural language; when a voucher has both an inclusions row and an exclusion row, the group type stated in the request disambiguates which row is targeted; when no matching row is found, the bot offers to insert instead and requires explicit user confirmation before any write occurs

### SharePoint Integration

- [ ] **SPO-01**: Bot reads Brief.xlsx from SharePoint via Microsoft Graph API instead of local file system
- [ ] **SPO-02**: Bot writes updated Brief.xlsx back to SharePoint via Microsoft Graph API, preserving all formatting and formulas
- [ ] **SPO-03**: After any successful change (add/update/delete), bot returns the SharePoint web URL so the user can open the file in Excel Online
- [ ] **SPO-04**: SharePoint site URL and document library path are configurable via environment variables (no hardcoding)

### Audit Trail

- [ ] **AUD-01**: Every successful change (add, update, delete) outputs a structured table in Teams chat: Campaign | Voucher | Field | Before | After | Changed by | Time
- [ ] **AUD-02**: "Changed by" is derived from the Teams user identity (display name from activity)
- [ ] **AUD-03**: "Time" is the UTC timestamp of the change

## v2 Requirements

### Permissions

- **PERM-01**: Different Teams users have read-only vs read-write access based on their AAD group membership

### Bulk Operations

- **BULK-01**: User can import multiple rows from a pasted table or CSV snippet
- **BULK-02**: User can request a filtered export (e.g., "show me all BAU campaigns")

## Out of Scope

| Feature | Reason |
|---------|--------|
| Real-time collaboration / conflict detection | Single-user update flow; complexity not justified for v1 |
| Modifying other Excel tabs | Scope limited to Promo-Voucher and Included-Excluded EANs per product decision |
| Modifying auto-calculated columns | Enforced by agent instructions; preserves Excel formula integrity |
| Direct file download URL | User requested SharePoint web URL (open in browser) instead |
| Promo-Voucher insert (add row) | Already shipped and validated — not a remaining requirement |

## Traceability

| Requirement | Phase | Status |
|-------------|-------|--------|
| ROW-01 | Phase 1 | Pending |
| ROW-02 | Phase 1 | Pending |
| ROW-03 | Phase 1 | Deferred (per 01-CONTEXT.md) |
| ROW-04 | Phase 1 | Deferred (per 01-CONTEXT.md) |
| ROW-05 | Phase 1 | Pending |
| AUD-01 | Phase 1 | Deferred (per 01-CONTEXT.md) |
| AUD-02 | Phase 1 | Deferred (per 01-CONTEXT.md) |
| AUD-03 | Phase 1 | Deferred (per 01-CONTEXT.md) |
| SPO-01 | Phase 2 | Pending |
| SPO-02 | Phase 2 | Pending |
| SPO-03 | Phase 2 | Pending |
| SPO-04 | Phase 2 | Pending |

**Coverage:**
- v1 requirements: 12 total
- Mapped to phases: 12
- Unmapped: 0

---
*Requirements defined: 2026-03-20*
*Last updated: 2026-03-25 — ROW-01/ROW-02 redefined to EANs insert scope (Promo-Voucher insert already shipped); ROW-03/ROW-04/AUD-01-03 marked Deferred per phase 1 context session; ROW-05 added for EANs update with group-type disambiguation and no-match offer flow*
