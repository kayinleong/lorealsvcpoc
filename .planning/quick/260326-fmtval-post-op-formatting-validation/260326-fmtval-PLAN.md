# Quick Task: 260326-fmtval — Post-operation formatting validation

## Goal
After every successful update or add row operation on Brief.xlsx, validate that the file's formatting has not been corrupted by comparing key XML structures against the reference file `Brief - Original Format.xlsx`.

## Implementation

### Changes to `src/app/app.ts`
1. Added `ORIGINAL_FORMAT_PATH` constant pointing to `.data/Brief - Original Format.xlsx`
2. Added `validateBriefFormatting()` async function that:
   - Compares `xl/styles.xml` byte-for-byte (detects xlsx fallback rewrite)
   - Compares `<cols>` block per sheet (column widths / default styles)
   - Compares `<sheetFormatPr>` per sheet (default row/col dimensions)
   - Logs `✓ Formatting intact` or `⚠️ N formatting issue(s)` with details
3. Called `validateBriefFormatting()` after `applyUpdate` success (all entries)
4. Called `validateBriefFormatting()` after `applyAdd` success

## Validation issues detected → logged to console only (non-blocking)
Issues are surfaced in server logs so the operator can investigate without interrupting the user experience.
