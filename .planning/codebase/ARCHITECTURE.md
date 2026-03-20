# Architecture

**Analysis Date:** 2026-03-20

## Pattern Overview

**Overall:** Agent-based Bot Framework architecture with event-driven message handling

**Key Characteristics:**
- Serverless bot application deployed on Azure App Service
- Event-driven message processing via Microsoft Teams Bot Framework SDK
- OpenAI integration for conversational AI responses
- Conversation history storage with local in-memory storage
- Managed identity authentication using Azure identity credentials

## Layers

**Presentation/Entry Point Layer:**
- Purpose: Expose the bot as a web service to Microsoft Teams
- Location: `src/index.ts`
- Contains: Application server bootstrap and startup logic
- Depends on: Express/Teams.apps framework, App configuration
- Used by: Teams client, bot framework

**Application/Core Logic Layer:**
- Purpose: Define bot behavior and message handling
- Location: `src/app/app.ts`
- Contains: Message event handlers, prompt orchestration, feedback handlers
- Depends on: Teams.ai ChatPrompt, Teams.apps App, OpenAI models, Storage
- Used by: Entry point layer, event handlers

**Configuration Layer:**
- Purpose: Centralize environment-based settings
- Location: `src/config.ts`
- Contains: Microsoft App credentials, Azure OpenAI endpoint settings
- Depends on: Environment variables
- Used by: Application layer, authentication

**Supporting Assets:**
- Purpose: AI model instructions and system prompts
- Location: `src/app/instructions.txt`
- Contains: System instructions for AI agent behavior
- Used by: ChatPrompt during message processing

## Data Flow

**Message Reception and Processing:**

1. User sends message in Teams (group chat or 1:1)
2. Teams Bot Framework routes message to POST `/api/messages` endpoint
3. `app.on("message")` handler receives MessageActivity
4. Application retrieves conversation history from storage using key `${conversationId}/${userId}`
5. ChatPrompt orchestrates:
   - Loads system instructions from `src/app/instructions.txt`
   - Creates OpenAI request with conversation history + new message
   - Sends request to Azure OpenAI deployment
6. Response handling differs by chat type:
   - **Group Chat:** Sends full response via `send()` with AI indicator
   - **1:1 Chat:** Streams response chunks via `stream.emit()` with AI indicator and feedback buttons
7. Conversation history updated in storage
8. Error handling sends user-friendly error messages

**Feedback Flow:**

1. User clicks feedback button on AI-generated message
2. `app.on("message.submit.feedback")` handler invokes
3. Feedback value logged (placeholder for custom feedback processing)

**State Management:**

- **Storage:** `LocalStorage` instance (in-memory key-value store)
- **Conversation Key:** `${conversation.id}/${from.id}` pattern
- **Stored Data:** Array of message objects (conversation history)
- **Persistence:** In-process memory only (lost on restart)

## Key Abstractions

**App (Teams.apps.App):**
- Purpose: Core bot application container managing Teams integration
- Examples: `src/app/app.ts` (instantiation and export)
- Pattern: Singleton instance with event-based message handling
- Configuration: Created with storage and credentials

**ChatPrompt (Teams.ai.ChatPrompt):**
- Purpose: Orchestrate conversation with AI model
- Pattern: Instantiated per message with conversation history and instructions
- Behavior: Supports streaming (1:1) and direct send (group)
- Configuration: Model instance passed with API key, endpoint, version

**OpenAIChatModel (Teams.openai.OpenAIChatModel):**
- Purpose: Interface to Azure OpenAI
- Pattern: Configured with endpoint, API key, deployment name, API version
- Created fresh per message to ensure credential freshness

**MessageActivity (Teams.api.MessageActivity):**
- Purpose: Represent outgoing message with metadata
- Pattern: Chainable builder pattern (`.addAiGenerated().addFeedback()`)
- Contains: Message content, AI indicator, feedback enabled

**TokenFactory:**
- Purpose: Generate Azure identity tokens for managed identity authentication
- Pattern: Closure-based factory in `createTokenFactory()`
- Behavior: Uses ManagedIdentityCredential to acquire tokens with specified scopes
- Authentication: Works with UserAssignedMSI when BOT_TYPE is "UserAssignedMsi"

## Entry Points

**Application Start:**
- Location: `src/index.ts`
- Triggers: `npm run dev` (development) or `npm start` (production)
- Responsibilities:
  - Import app configuration
  - Start app listening on port (from PORT, port, or default 3978)
  - Log startup success

**Message Event Handler:**
- Location: `src/app/app.ts` line 58
- Triggers: Incoming message activity from Teams
- Responsibilities:
  - Retrieve conversation history
  - Create ChatPrompt with instructions and model
  - Route response based on chat context (group vs 1:1)
  - Store updated conversation history
  - Handle and log errors

**Feedback Event Handler:**
- Location: `src/app/app.ts` line 102
- Triggers: User feedback submission
- Responsibilities: Log feedback value (extensible for custom processing)

## Error Handling

**Strategy:** Try-catch with user-facing error messages and console logging

**Patterns:**

- **Prompt/Model Errors:** Caught in message handler, sends two user messages:
  1. "The agent encountered an error or bug."
  2. "To continue to run this agent, please fix the agent source code."
- **Token Acquisition Errors:** Not explicitly caught (propagates as unhandled rejection)
- **Storage Errors:** Not explicitly caught (would cause message handler to fail)
- **Logging:** All errors logged to console with `console.error()`

## Cross-Cutting Concerns

**Logging:**
- Tool: Node.js console
- Pattern: `console.log()` for startup and feedback, `console.error()` for exceptions
- Locations: Entry point startup message, feedback logging, error handling

**Validation:**
- Environment variables validated at runtime via `process.env` access
- Configuration loads without validation; missing values cause runtime errors
- No input validation on message content (passed directly to OpenAI)

**Authentication:**
- Local Development: ClientSecret-based (from `.localConfigs` file)
- Azure Production: Managed Identity (UserAssignedMSI)
- Token Factory: Dynamically acquires tokens per request
- Bot Framework: Validates incoming requests using MicrosoftAppId and MicrosoftAppPassword

**Conversation Storage:**
- Pattern: Key-value store with composite key (conversationId/userId)
- Scope: Per-conversation per-user (isolated chat histories)
- Lifecycle: Lost on application restart
- No encryption or database persistence

---

*Architecture analysis: 2026-03-20*
