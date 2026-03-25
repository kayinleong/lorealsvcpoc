# Phase 1: Row Operations + Audit Trail - Context

**Gathered:** 2026-03-25
**Status:** Ready for planning

<domain>
## Phase Boundary

Insert and update rows in the Included-Excluded EANs- Voucher tab's Voucher Dashboard table (Voucher Name | Inclusions Group | Exclusion Group) via natural language.

**Already shipped (out of scope for this phase):**
- Promo-Voucher: insert new row, update existing row, Malay language context — these are complete and working.

**What this phase delivers:**
- EANs tab: Insert a new Voucher Dashboard row (Voucher Name + either Inclusions Group or Exclusion Group)
- EANs tab: Update any column on an existing Voucher Dashboard row

</domain>

<decisions>
## Implementation Decisions

### Row Identification for Update
- Target row is identified by Voucher Name + group type (Inclusion vs Exclusion)
- When a user says "update inclusions group for [voucher] to [value]", the bot targets the row that has the given Voucher Name AND has an Inclusions Group value (i.e., the Inclusions-type row for that voucher)
- This resolves ambiguity when a voucher has both an Inclusions row and an Exclusion row

### Insert Validation
- Mutual exclusivity enforced: one row must have either Inclusions Group OR Exclusion Group — not both. Reject inserts where both are provided; prompt user to split into two rows.
- Voucher Name must match an existing Voucher Name in the Promo-Voucher tab. Reject inserts with an unrecognised voucher name.

### Update Scope
- User can update any of the 3 columns: Voucher Name, Inclusions Group, or Exclusion Group
- No restriction on switching group type via update (user can change Inclusions Group value, Exclusion Group value, or even the Voucher Name reference)

### No-Match Behavior on Update
- If no matching row is found (by Voucher Name + group type), the bot offers to insert instead: "No existing [inclusions/exclusion] row found for [voucher]. Would you like to add it?"
- User must confirm before an insert is triggered from a failed update attempt

### Claude's Discretion
- Exact phrasing of the "offer to insert" prompt
- How to handle case-insensitive / partial matching on Voucher Name for the EANs tab (follow same pattern as Promo-Voucher tab matching)

</decisions>

<code_context>
## Existing Code Insights

### Reusable Assets
- `UpdateAction` interface (`src/app/app.ts:192`): `{operation, tab, campaign_name?, voucher_name?, field, value}` — reusable for EANs updates; `campaign_name` is optional so it works without a campaign
- `AddAction` interface (`src/app/app.ts:201`): `{operation, tab, rows: Record<string, string|number|null>[]}` — reusable for EANs inserts
- `applyUpdate` / `applyUpdateSurgical` (`src/app/app.ts:315, 759`): surgical XML cell update; currently finds row by voucher_name match on col 1 — needs enhancement to also match by group type (which column is populated)
- `applyAdd` (`src/app/app.ts:605`): surgical row insert into pre-styled rows; works for any tab including EANs
- `parseAction` / `stripAction` (`src/app/app.ts:210, 221`): no changes needed — already generic

### Established Patterns
- Action blocks: LLM emits `<action>{...}</action>` JSON; code parses and executes — same pattern applies for EANs operations
- Row targeting: existing code matches rows by `voucher_name` substring on col 1 (EANs) — this is the right column, but multi-row disambiguation by group type needs to be added
- EANs tab constants: `EANS_HEADER_ROW = 4` (row 5 in Excel), `EANS_DATA_ROW = 6` (row 7 in Excel), max 11 cols

### Integration Points
- `instructions.txt`: needs new sections for EANs tab insert and update — action block format, validation rules, row targeting by group type
- `applyUpdate`/`applyUpdateSurgical`: add group-type disambiguation when `action.tab === EANS_SHEET` — find row where `voucher_name` matches AND the target group-type column (col 2 or col 3) is non-empty
- Message handler (`src/app/app.ts:1013`): no changes needed — already routes by `action.operation`

</code_context>

<specifics>
## Specific Ideas

- "Voucher Dashboard" is the user's name for the cols 1-3 view of the Included-Excluded EANs- Voucher tab (Voucher Name | Inclusions Group | Exclusion Group)
- The same tab also has cols 6-7 (Group Name, SKU EAN) — those are out of scope for this phase
- Row targeting must handle the case where one voucher has both an Inclusions-type row and an Exclusion-type row; the group type in the user's request disambiguates which row to act on

</specifics>

<deferred>
## Deferred Ideas

- Delete row from EANs Voucher Dashboard — not requested for this phase
- Audit trail / structured change confirmation table — originally in Phase 1 roadmap but not requested; defer until needed
- Group Name and SKU EAN column operations (cols 6-7) — separate capability, own phase
- Promo-Voucher delete row — not requested for this phase

</deferred>

---

*Phase: 01-row-operations-audit-trail*
*Context gathered: 2026-03-25*
