# Coding Conventions

**Analysis Date:** 2026-03-20

## Naming Patterns

**Files:**
- Lowercase with dot separators for file names: `config.ts`, `app.ts`, `index.ts`
- Descriptive names indicating purpose: `instructions.txt` for prompt definitions
- Use camelCase for multi-word descriptors

**Functions:**
- camelCase for function names: `loadInstructions()`, `createTokenFactory()`
- Descriptive verb-based names indicating action
- Arrow functions for factory patterns: `const createTokenFactory = () => { ... }`
- Standard function declarations for named functions

**Variables:**
- camelCase for all variable and constant names: `storage`, `conversationKey`, `instructions`, `tokenResponse`
- Descriptive names indicating content or purpose
- Const for immutable values (preferred): `const storage = new LocalStorage()`
- Const for factory functions and initialized objects

**Types:**
- PascalCase for class/type names (imported from external libraries): `App`, `ChatPrompt`, `LocalStorage`, `OpenAIChatModel`, `MessageActivity`, `TokenCredentials`, `ManagedIdentityCredential`
- Interface names follow PascalCase convention
- Object literals use camelCase properties: `MicrosoftAppId`, `azureOpenAIKey`, `azureOpenAIEndpoint`

## Code Style

**Formatting:**
- No explicit formatter configured (no .prettierrc or .eslintrc files present)
- Two-space indentation observed in codebase
- Template literals used for dynamic strings: `` `\nAgent started, app listening to` ``
- String concatenation with `+` operator: `"Your feedback is " + JSON.stringify(activity.value)`

**Linting:**
- No linting tool configured in project
- TypeScript compilation without ESLint configuration
- Relies on TypeScript type checking for code quality

## Import Organization

**Order:**
1. Third-party npm packages (external frameworks): `import { App } from "@microsoft/teams.apps"`
2. Additional third-party packages: `import { LocalStorage } from "@microsoft/teams.common"`
3. Node built-in modules: `import * as fs from "fs"` and `import * as path from "path"`
4. Local application imports: `import config from "../config"`

**Path Aliases:**
- No path aliases configured in tsconfig.json
- Relative imports used throughout: `"../config"`, `"./app/app"`
- Namespace imports for built-in modules: `import * as fs`, `import * as path`

## Error Handling

**Patterns:**
- Try-catch blocks for async operations in event handlers
- Generic error catch blocks that log to console: `catch (error) { console.error(error); }`
- User-friendly error messages sent to client: `"The agent encountered an error or bug."`
- Fallback error responses guide users to source code fixes
- Error context preserved via console logging before responding to user

**Example from `src/app/app.ts`:**
```typescript
try {
  const prompt = new ChatPrompt({...});
  // ... operation code
} catch (error) {
  console.error(error);
  await send("The agent encountered an error or bug.");
  await send("To continue to run this agent, please fix the agent source code.");
}
```

## Logging

**Framework:** console (native Node.js)

**Patterns:**
- `console.log()` for informational messages: startup messages, feedback logs
- `console.error()` for error conditions: exception logging
- Direct logging without formatting library
- String concatenation for log messages: `"Your feedback is " + JSON.stringify(activity.value)`
- Template literals for complex messages: `` `\nAgent started, app listening to` ``

**Example usage:**
```typescript
console.log(`\nAgent started, app listening to`, process.env.PORT || process.env.port || 3978);
console.error(error);
console.log("Your feedback is " + JSON.stringify(activity.value));
```

## Comments

**When to Comment:**
- Inline comments before complex logic blocks (limited use)
- High-level comments explaining intent rather than implementation
- Comments placed on separate lines before code blocks

**Example from `src/app/app.ts`:**
```typescript
// Create storage for conversation history
const storage = new LocalStorage();

// Load instructions from file on initialization
function loadInstructions(): string {

// Load instructions once at startup
const instructions = loadInstructions();

// Get conversation history
const conversationKey = ...

// Configure authentication using TokenCredentials
const tokenCredentials: TokenCredentials = {
```

**JSDoc/TSDoc:**
- Not used in current codebase
- No function documentation strings present
- Implicit type hints via TypeScript annotations

## Function Design

**Size:**
- Small focused functions: `loadInstructions()` is 4 lines
- Factory functions for creating complex objects: `createTokenFactory()` is 14 lines
- Inline event handlers for specific operations

**Parameters:**
- Named parameters for clarity
- Type annotations on all parameters: `scope: string | string[]`, `tenantId?: string`
- Optional parameters marked with `?`: `tenantId?: string`

**Return Values:**
- Explicit return type annotations: `: string`, `: Promise<string>`
- Async functions return Promises: `async (...): Promise<string>`
- Storage methods return data or undefined: `storage.get(conversationKey) || []`

## Module Design

**Exports:**
- Default exports for main modules: `export default app` (app.ts, config.ts)
- Single responsibility modules
- Configuration object exported as default export

**Barrel Files:**
- Not used in this codebase
- Direct imports from individual files preferred

**File Organization:**
- `src/index.ts` - Application entry point, async IIFE for startup
- `src/config.ts` - Configuration constants from environment
- `src/app/app.ts` - Main application logic with event handlers

---

*Convention analysis: 2026-03-20*
