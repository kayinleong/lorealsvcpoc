# Requirements: L'Oreal Brief Management Agent

**Defined:** 2026-03-20
**Core Value:** Business users can update the promotional brief directly from Teams using plain language

## v1 Requirements

Requirements for this extension milestone. Validated capabilities are already shipped.

### Row Operations

- [ ] **ROW-01**: User can add a new voucher row to Promo-Voucher via natural language, providing Campaign Name, Voucher Name, Tier, Start/End Dates, Voucher Type, Discount Type, and relevant discount fields
- [ ] **ROW-02**: New row addition enforces all existing validation rules (Min Spend, Discount Type consistency, Amount Capped, Voucher Budget/Quantity exclusivity, Voucher Code rules)
- [ ] **ROW-03**: User can delete a voucher row from Promo-Voucher via natural language (immediate, no confirmation prompt)
- [ ] **ROW-04**: After successful add or delete, bot outputs a table-format audit summary (Campaign | Voucher | Operation | Changed by | Time)

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

## Traceability

| Requirement | Phase | Status |
|-------------|-------|--------|
| ROW-01 | Phase 1 | Pending |
| ROW-02 | Phase 1 | Pending |
| ROW-03 | Phase 1 | Pending |
| ROW-04 | Phase 1 | Pending |
| AUD-01 | Phase 1 | Pending |
| AUD-02 | Phase 1 | Pending |
| AUD-03 | Phase 1 | Pending |
| SPO-01 | Phase 2 | Pending |
| SPO-02 | Phase 2 | Pending |
| SPO-03 | Phase 2 | Pending |
| SPO-04 | Phase 2 | Pending |

**Coverage:**
- v1 requirements: 11 total
- Mapped to phases: 11
- Unmapped: 0

---
*Requirements defined: 2026-03-20*
*Last updated: 2026-03-20 — traceability finalized after roadmap creation*
