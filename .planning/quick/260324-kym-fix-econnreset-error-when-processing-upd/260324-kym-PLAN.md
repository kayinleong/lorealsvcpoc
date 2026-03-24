---
phase: quick
plan: 260324-kym
type: execute
wave: 1
depends_on: []
files_modified:
  - src/app/app.ts
autonomous: true
requirements: []
must_haves:
  truths:
    - "Bot does not crash with unhandled rejection when ECONNRESET occurs"
    - "Error in catch-block send() is safely caught and logged"
    - "Transient network errors on send() are retried once before giving up"
    - "All errors during message processing are logged with enough context for debugging"
  artifacts:
    - path: "src/app/app.ts"
      provides: "Resilient message handler with nested error handling and retry"
  key_links:
    - from: "app.on('message') catch block"
      to: "send()"
      via: "try/catch wrapping the fallback send"
      pattern: "catch.*send.*catch"
---

<objective>
Fix ECONNRESET error when processing "Update PDS x Birthday Full Day Min Spend to RM1000" input.

Purpose: The bot crashes with an unhandled ECONNRESET when processing this input. The error occurs in the message handler's catch block — when the primary processing fails and the catch block tries to send an error message back to the user, that second `send()` also fails (connection already reset), producing an unhandled rejection that crashes the bot.

Output: Resilient message handler in src/app/app.ts that gracefully handles send failures.
</objective>

<execution_context>
@C:/Users/ka.yin.leong/.claude/get-shit-done/workflows/execute-plan.md
@C:/Users/ka.yin.leong/.claude/get-shit-done/templates/summary.md
</execution_context>

<context>
@src/app/app.ts
</context>

<tasks>

<task type="auto">
  <name>Task 1: Add resilient error handling with retry and safe fallback send</name>
  <files>src/app/app.ts</files>
  <action>
Refactor the message handler in src/app/app.ts (lines 296-343) with the following changes:

1. **Extract a `safeSend` helper** above the message handler:
   ```typescript
   async function safeSend(
     send: (activity: unknown) => Promise<unknown>,
     activity: unknown,
     retries = 1
   ): Promise<boolean> {
     for (let attempt = 0; attempt <= retries; attempt++) {
       try {
         await send(activity);
         return true;
       } catch (err: unknown) {
         const code = (err as NodeJS.ErrnoException)?.code;
         const isTransient = code === 'ECONNRESET' || code === 'ETIMEDOUT' || code === 'ECONNREFUSED';
         if (isTransient && attempt < retries) {
           console.warn(`safeSend: transient error (${code}), retrying (${attempt + 1}/${retries})...`);
           await new Promise(r => setTimeout(r, 1000));
           continue;
         }
         console.error(`safeSend: failed to deliver message after ${attempt + 1} attempt(s):`, err);
         return false;
       }
     }
     return false;
   }
   ```

2. **Replace the `await send(...)` call on line 335** (the success path) with:
   ```typescript
   await safeSend(send, new MessageActivity(displayText).addAiGenerated().addFeedback());
   ```

3. **Wrap the catch block's error send in a nested try/catch** to prevent unhandled rejections:
   Replace lines 337-342:
   ```typescript
   } catch (error) {
     console.error("Brief agent error:", error);
     try {
       await safeSend(
         send,
         "Sorry, I ran into an issue processing your request. Please try again or rephrase your message."
       );
     } catch (sendError) {
       console.error("Failed to send error message to user:", sendError);
     }
   }
   ```

4. **Add structured error logging** — in the catch block, log the user's input text and conversation key for debugging:
   ```typescript
   console.error("Brief agent error:", { conversationKey, userText: userText.substring(0, 200), error });
   ```

Key points:
- The `safeSend` helper retries once on transient network errors (ECONNRESET, ETIMEDOUT, ECONNREFUSED) with a 1-second delay
- The catch block's fallback send is itself wrapped in try/catch so it can never produce an unhandled rejection
- Error logs include context (conversation key, truncated user text) for debugging
- Do NOT change any business logic — only the error handling and send patterns
  </action>
  <verify>
    <automated>cd C:/Development/LorealSVCPOC && npx tsc --noEmit 2>&1 | head -20</automated>
  </verify>
  <done>
    - safeSend helper exists with retry logic for transient errors
    - Primary send() on line 335 uses safeSend
    - Catch block's fallback send is wrapped in nested try/catch via safeSend
    - Error logging includes conversationKey and truncated userText
    - TypeScript compiles without errors
    - No changes to business logic (parsing, validation, applyUpdate)
  </done>
</task>

</tasks>

<verification>
- `npx tsc --noEmit` passes with no errors
- Manual review: search for bare `await send(` — none should remain outside safeSend
- The catch block at the end of the message handler has a nested try/catch around its send call
</verification>

<success_criteria>
- Bot no longer crashes with unhandled ECONNRESET when the fallback error message send fails
- Transient network errors get one retry before giving up
- All error paths are logged with sufficient context for debugging
- No regression in normal message processing flow
</success_criteria>

<output>
After completion, create `.planning/quick/260324-kym-fix-econnreset-error-when-processing-upd/260324-kym-SUMMARY.md`
</output>
