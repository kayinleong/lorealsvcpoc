# Codebase Concerns

**Analysis Date:** 2026-03-20

## Tech Debt

**Synchronous File I/O in Startup Path:**
- Issue: `fs.readFileSync()` is used to load instructions during module initialization, blocking the event loop at startup
- Files: `src/app/app.ts:15-17`
- Impact: Application startup is delayed by file system latency; affects total time-to-ready in containerized environments and slow storage backends
- Fix approach: Convert to async file loading with promise-based initialization or load instructions lazily on first message

**Unvalidated Environment Variables:**
- Issue: Configuration values are read directly from `process.env` with minimal validation. Missing required variables fall back to defaults or undefined
- Files: `src/config.ts`, `src/app/app.ts:29, 42`
- Impact: Silent failures when required credentials are missing; app may start but fail on first interaction; harder to debug production issues
- Fix approach: Add startup validation that throws explicit errors with clear messages for missing critical env vars (CLIENT_ID, AZURE_OPENAI_API_KEY, etc.)

**Hardcoded API Version:**
- Issue: Azure OpenAI API version `2024-10-21` is hardcoded in the message handler
- Files: `src/app/app.ts:71`
- Impact: Code will break when Microsoft deprecates this API version; requires code changes and redeployment to update
- Fix approach: Move API version to environment variable with default fallback

**Generic Error Handling with Limited Context:**
- Issue: Catch block in message handler logs error then sends generic user message without error details or recovery mechanism
- Files: `src/app/app.ts:93-99`
- Impact: Hard to diagnose real issues in production; users see generic error message without guidance; debugging requires checking server logs
- Fix approach: Implement structured error logging with error codes; provide specific error guidance to users based on error type (auth errors, rate limits, etc.)

## Known Bugs

**Conversation History Not Persisted:**
- Symptoms: `storage.set()` is called AFTER streaming in single chat but conversation history may not contain the latest user message or response
- Files: `src/app/app.ts:60-61, 92`
- Trigger: User sends message, receives response via streaming; history updated after response completes
- Issue: The `messages` array is initialized from storage but the current activity.text is never added before sending to prompt; response is not added either
- Workaround: Current messages are retrieved from storage on next interaction but current exchange is not preserved
- Fix approach: Explicitly add user message and response to messages array before storage.set()

**Missing Fallback for ManagedIdentityCredential:**
- Symptoms: App will crash during token creation if ManagedIdentityCredential.getToken() fails or times out
- Files: `src/app/app.ts:28-36`
- Trigger: Attempting to send first message when managed identity endpoint is unreachable (e.g., no VM identity assigned, network timeout)
- Workaround: None - process dies
- Fix approach: Add try-catch with fallback credential strategy or timeout with explicit error messaging

## Security Considerations

**Sensitive Credentials in Process Memory:**
- Risk: API keys (AZURE_OPENAI_API_KEY, CLIENT_SECRET) are stored in plain text in memory throughout app lifetime
- Files: `src/config.ts`, `src/app/app.ts:69`
- Current mitigation: Relies on operating system memory protection; credentials only accessible to process owner
- Recommendations:
  - Consider using Azure Key Vault client library to fetch secrets on-demand rather than loading at startup
  - Implement credential rotation strategy
  - Never log or print config object in debug output

**No Authentication on API Endpoints:**
- Risk: If deployed, bot endpoints would be exposed to internet without request validation
- Files: App framework (external) - no custom endpoints currently visible
- Current mitigation: Microsoft Teams SDK handles authentication at framework level
- Recommendations:
  - Document that all endpoints require valid Teams request signing
  - Add custom request validation middleware if extending with additional endpoints
  - Regularly audit Bot Framework security bulletins

**Insufficient Validation of Incoming Messages:**
- Risk: User activity.text is passed directly to OpenAI without length checks or content validation
- Files: `src/app/app.ts:78, 84`
- Current mitigation: Azure OpenAI API enforces token limits
- Recommendations:
  - Add client-side message length validation
  - Implement rate limiting per user/conversation
  - Add checks for suspicious content patterns

**No HTTPS Enforcement Visible:**
- Risk: If communications aren't over HTTPS, credentials and messages could be intercepted
- Files: App configuration (external/deployment)
- Current mitigation: Microsoft Teams SDK enforces HTTPS in production; local dev uses HTTP
- Recommendations: Document HTTPS-only deployment requirement; implement HSTS headers if exposing custom endpoints

## Performance Bottlenecks

**Synchronous Instructions File Load on Each Module Init:**
- Problem: `fs.readFileSync()` blocks during module import, before app even starts handling requests
- Files: `src/app/app.ts:15-21`
- Cause: Instructions loaded at module evaluation time, not in async context
- Improvement path:
  1. Move loadInstructions() to async app initialization
  2. Cache loaded instructions in memory
  3. Add file watcher for hot-reload in development

**Full Conversation History Sent on Every Message:**
- Problem: Entire conversation history is loaded from storage and sent to OpenAI API on every user message
- Files: `src/app/app.ts:60-65, 78, 84`
- Cause: No pagination, truncation, or summarization of old messages
- Improvement path:
  1. Implement sliding window - keep last N messages only
  2. Add message pruning - remove old messages beyond token limit
  3. Consider summarization for very long conversations
  4. Monitor API costs per conversation

**No Caching of OpenAI Models:**
- Problem: New OpenAIChatModel instance created on every incoming message
- Files: `src/app/app.ts:67-72`
- Cause: Model configuration instantiated in request handler rather than application initialization
- Improvement path:
  1. Create model instance once at app startup
  2. Reuse for all requests
  3. Add connection pooling if using HTTP-based client

**LocalStorage Implementation Performance Unknown:**
- Problem: `new LocalStorage()` performance characteristics not documented; implementation may be in-memory only
- Files: `src/app/app.ts:12`
- Cause: External dependency; unclear if it persists across restarts or is memory-only
- Improvement path: Investigate actual implementation; if memory-only, add persistent storage layer (database, file-based, etc.)

## Fragile Areas

**Configuration Module:**
- Files: `src/config.ts`
- Why fragile:
  - No schema validation - any missing env var silently becomes undefined
  - No type safety for config values - all are `string | undefined`
  - Exports plain object with no accessor methods
  - If consuming code tries to use undefined values, errors surface far from source
- Safe modification:
  - Add validation function that throws on missing required vars
  - Use const assertion and typed exports
  - Add getConfig() function with null checks
- Test coverage: No tests exist for missing environment variables

**App Initialization:**
- Files: `src/app/app.ts`
- Why fragile:
  - Module-level side effects (instruction loading, storage creation, app instantiation)
  - Errors in setup propagate to module import, not caught until app startup
  - ManagedIdentityCredential created lazily in token factory, errors only surface on first message
  - TokenCredentials object includes async function but typed as synchronous
- Safe modification:
  - Move all setup to explicit async init function called from index.ts
  - Add startup health checks before marking app ready
  - Test failure scenarios explicitly
- Test coverage: Zero test coverage

**Message Handler:**
- Files: `src/app/app.ts:58-100`
- Why fragile:
  - Complex branching logic for group vs individual chat
  - Storage interaction not atomic - history may diverge if request fails between get and set
  - Feedback handler not integrated with main handler - separate codepath with minimal logging
  - Stream emission logic unclear - stream.emit() behavior undocumented
- Safe modification:
  - Add explicit transaction/checkpoint mechanism for storage
  - Add detailed logging at each step
  - Test both chat types with error injection
- Test coverage: Zero

## Scaling Limits

**In-Memory Conversation Storage:**
- Current capacity: Single LocalStorage instance; unbounded by application code
- Limit: Will consume increasing memory as conversations accumulate; eventual OOM if process runs long enough
- Scaling path:
  1. Implement time-based eviction (TTL on stored conversations)
  2. Switch to persistent backend (database, Redis, blob storage)
  3. Implement conversation cleanup on schedule
  4. Add memory metrics monitoring

**Single Process Constraint:**
- Current capacity: Single Node.js process handles all users
- Limit: CPU-bound streaming responses will block other users; OpenAI API calls serialize
- Scaling path:
  1. Deploy multiple instances behind load balancer
  2. Implement distributed session store for conversation history
  3. Add message queue for async processing
  4. Consider serverless deployment model

**API Rate Limits:**
- Current capacity: No explicit rate limiting on incoming messages
- Limit: Single user can flood with requests; no protection against Azure OpenAI quota limits
- Scaling path:
  1. Add per-user rate limiting (e.g., 10 requests/minute)
  2. Implement per-conversation rate limiting
  3. Add queue with backpressure when Azure OpenAI quota approached
  4. Implement graceful degradation

## Dependencies at Risk

**@microsoft/teams.* Packages (2.0.0):**
- Risk: Locked to major version 2.0.0; no minor version flexibility indicated in package.json
- Impact: Security patches for framework won't auto-install; framework bugs unfixable without major version bump
- Migration plan:
  1. Update package.json to allow patch updates: `^2.0.0`
  2. Monitor Microsoft Teams SDK GitHub for security advisories
  3. Plan quarterly updates minimum

**@azure/identity (^4.11.1):**
- Risk: Version uses caret range (allows minor/patch); but ManagedIdentityCredential API may change in future major versions
- Impact: Azure authentication could break on major upgrade
- Migration plan:
  1. Add version pinning tests that fail on major upgrade
  2. Subscribe to Azure SDK release notes
  3. Plan compatibility validation before upgrading majors

**typescript (~5.8.3):**
- Risk: Tilde range allows patch updates only; new TypeScript releases may introduce breaking check changes
- Impact: Build could break when switching to new minor version
- Migration plan:
  1. Consider updating to `^5.8.3` if strict mode can handle minor updates
  2. Add CI step to test with next minor version
  3. Monitor TypeScript release notes for breaking changes

**Missing Production Dependencies:**
- Risk: No logging framework, no monitoring SDK, no error tracking
- Impact: Issues in production have no structured visibility
- Migration plan:
  1. Add `winston` or `pino` for structured logging
  2. Add Application Insights SDK for Azure monitoring
  3. Add Sentry or similar for error tracking
  4. Implement health check endpoint

## Missing Critical Features

**No Test Suite:**
- Problem: `npm test` exits with error message; zero test coverage
- Blocks: Cannot safely refactor or upgrade dependencies; no regression protection
- Must add: Unit tests for message handling, config validation, error cases; integration tests for Teams SDK interaction

**No Logging Infrastructure:**
- Problem: Only `console.log/error` used; no structured logging, no log levels, no context
- Blocks: Production debugging nearly impossible; cannot correlate requests across services
- Must add: Winston or Pino logger with request correlation IDs, structured fields, appropriate log levels

**No Health/Status Endpoints:**
- Problem: No way to check app readiness in production
- Blocks: Load balancers cannot verify app health; unclear if deployment succeeded
- Must add: GET /health endpoint that checks dependencies (storage, Azure OpenAI connectivity)

**No Configuration Validation:**
- Problem: Missing required environment variables fail silently on first request
- Blocks: Deployments can succeed but app non-functional; hard to debug
- Must add: Startup validation that throws before app starts listening

**No Conversation Persistence Strategy:**
- Problem: Conversation history only in memory; lost on process restart
- Blocks: Users lose conversation context; cannot implement conversation search/history
- Must add: Persistent storage selection (database, blob storage, etc.) and implementation

## Test Coverage Gaps

**Untested Startup Path:**
- What's not tested: Application initialization, config loading, app startup sequence
- Files: `src/index.ts`, `src/app/app.ts:1-21`, `src/config.ts`
- Risk: Startup failures only caught in deployed environment (e.g., missing config file, invalid env vars)
- Priority: High

**Untested Message Handler:**
- What's not tested: Core feature - actual message handling and AI responses
- Files: `src/app/app.ts:58-100`
- Risk: Changes to message flow, storage interaction, or streaming could break silently
- Priority: Critical

**Untested Error Scenarios:**
- What's not tested: Network failures, Azure OpenAI timeouts, malformed responses, credential failures
- Files: `src/app/app.ts:93-99`
- Risk: Error handling is generic; real failure modes unknown until production incident
- Priority: High

**Untested Token/Credential Flow:**
- What's not tested: ManagedIdentityCredential token acquisition, token refresh, token expiration
- Files: `src/app/app.ts:23-44`
- Risk: Auth failures only surface at runtime; no validation that token factory works
- Priority: High

**Untested Group vs Individual Chat Differences:**
- What's not tested: Branching logic for `activity.conversation.isGroup`
- Files: `src/app/app.ts:75-91`
- Risk: One code path untested; group chat responses may have different behavior than 1:1
- Priority: Medium

**Untested Feedback Handler:**
- What's not tested: Feedback collection and storage
- Files: `src/app/app.ts:102-105`
- Risk: Feedback processing may be silently failing
- Priority: Low

---

*Concerns audit: 2026-03-20*
