---
phase: quick
plan: 260324-kym
subsystem: message-handler
tags: [error-handling, resilience, networking, econnreset]
dependency_graph:
  requires: []
  provides: [resilient-message-handler]
  affects: [src/app/app.ts]
tech_stack:
  added: []
  patterns: [retry-on-transient-error, safe-fallback-send, structured-error-logging]
key_files:
  created: []
  modified:
    - src/app/app.ts
decisions:
  - "safeSend never throws — it catches all errors internally and returns false, making a nested try/catch in the catch block unnecessary"
  - "One retry with 1-second delay chosen as the right balance — enough to survive a brief network blip without hanging the bot"
  - "safeSend is generic (typed on T) to handle both MessageActivity objects and plain strings"
metrics:
  duration: ~5m
  completed: 2026-03-24
---

# Quick Fix 260324-kym: Fix ECONNRESET Error in Message Handler Summary

**One-liner:** Added `safeSend` helper with one retry on transient network errors (ECONNRESET/ETIMEDOUT/ECONNREFUSED) and replaced all bare `send()` calls in the message handler to prevent unhandled rejections.

## What Was Done

The bot was crashing with an unhandled rejection when processing certain inputs (e.g. "Update PDS x Birthday Full Day Min Spend to RM1000"). The root cause: the primary message processing would fail (ECONNRESET from the Teams channel), and then the catch block's fallback `send()` call would also fail with ECONNRESET — producing a second unhandled rejection that crashed the bot.

### Changes to `src/app/app.ts`

**1. Added `safeSend` helper (lines 301-329)**

A generic async function that wraps any `send()` call with:
- Up to 1 retry on transient errors (ECONNRESET, ETIMEDOUT, ECONNREFUSED) with a 1-second delay between attempts
- `console.warn` logging on retry attempts
- `console.error` logging on final failure
- Returns `true` on success, `false` on failure — never throws

**2. Primary send path (line 370)**

Replaced:
```typescript
await send(new MessageActivity(displayText).addAiGenerated().addFeedback());
```
With:
```typescript
await safeSend(send, new MessageActivity(displayText).addAiGenerated().addFeedback());
```

**3. Catch block fallback send (lines 374-378)**

Replaced:
```typescript
try {
  await send("Sorry...");
} catch (sendError) { ... }
```
With:
```typescript
await safeSend(send, "Sorry, I ran into an issue...");
```
`safeSend` itself never throws, so the explicit nested try/catch is not needed.

**4. Structured error logging (line 373)**

Replaced:
```typescript
console.error("Brief agent error:", error);
```
With:
```typescript
console.error("Brief agent error:", { conversationKey, userText: userText.substring(0, 200), error });
```

## Verification

- `npx tsc --noEmit` passes with no errors
- No bare `await send(` calls remain outside `safeSend`
- Business logic (parsing, validation, `applyUpdate`) unchanged

## Commit

`7f55a5e` — fix: add safeSend helper to handle ECONNRESET in message handler

## Deviations from Plan

None — plan executed exactly as written. The plan showed a nested try/catch around `safeSend` in the catch block as one option; the implemented version omits it because `safeSend` is already guaranteed not to throw, which is functionally equivalent and cleaner.

## Self-Check: PASSED

- `src/app/app.ts` modified and committed: 7f55a5e
- TypeScript compiles: confirmed (no output = no errors)
- All bare `send(` calls replaced: confirmed
