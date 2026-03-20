# Technology Stack

**Analysis Date:** 2026-03-20

## Languages

**Primary:**
- TypeScript 5.8.3 - All source code in `src/` directory
- JavaScript - Generated from TypeScript compilation to `lib/`

**Secondary:**
- YAML - Deployment and configuration files (`m365agents.yml`, `m365agents.local.yml`, `m365agents.playground.yml`)
- XML - IIS configuration (`web.config`)
- JSON - Package manifests, app manifests, and ARM template parameters

## Runtime

**Environment:**
- Node.js 20 or 22 (defined in `package.json` engines field)

**Package Manager:**
- npm (uses `package-lock.json` lockfile present)

## Frameworks

**Core:**
- @microsoft/teams.apps 2.0.0 - Microsoft 365 Agents Toolkit application framework (`src/app/app.ts`)
- @microsoft/teams.ai 2.0.0 - AI and chat prompt handling (`src/app/app.ts`)
- @microsoft/teams.openai 2.0.0 - OpenAI integration for chat models (`src/app/app.ts`)

**Utilities:**
- @microsoft/teams.common 2.0.0 - Common utilities including LocalStorage (`src/app/app.ts`)
- @microsoft/teams.api - API types and abstractions like MessageActivity and TokenCredentials (`src/app/app.ts`)

**Build/Dev:**
- TypeScript ~5.8.3 - Language and transpilation
- ts-node 10.4.0 - TypeScript execution for development
- nodemon 3.1.7 - Development watch mode (`package.json` dev script)
- tsc (TypeScript Compiler) - Builds to `lib/` directory with source maps
- shx 0.4.0 - Cross-platform shell commands for build tasks (copying instructions)
- env-cmd 10.1.0 - Environment variable management for local development

## Key Dependencies

**Critical:**
- @azure/identity 4.11.1 - Azure authentication using ManagedIdentityCredential for token generation (`src/app/app.ts`)
- Node built-ins: fs, path - File system and path utilities for loading instructions.txt

**Infrastructure:**
- No database drivers or ORMs
- No HTTP client libraries (embedded in framework)
- No logging frameworks (uses console)

## Configuration

**Environment:**
- Environment variables loaded via:
  - `env-cmd` for development (`.localConfigs` file in `package.json` dev scripts)
  - Environment files in `env/` directory:
    - `.env.dev` - Development environment
    - `.env.local` - Local development
    - `.env.playground` - Playground environment
    - `.env.*.user` - User-specific overrides (not committed)

**Key environment variables required:**
- `CLIENT_ID` - Microsoft Entra application client ID (from `config.ts`)
- `CLIENT_SECRET` - Microsoft Entra application client secret (from `config.ts`)
- `BOT_TYPE` - Bot type (e.g., "UserAssignedMsi") (from `config.ts`)
- `TENANT_ID` - Azure tenant ID (from `config.ts`)
- `AZURE_OPENAI_API_KEY` - Azure OpenAI API key (from `config.ts`)
- `AZURE_OPENAI_ENDPOINT` - Azure OpenAI endpoint URL (from `config.ts`)
- `AZURE_OPENAI_DEPLOYMENT_NAME` - OpenAI deployment name (from `config.ts`)
- `PORT` - Application port (default 3978 in `src/index.ts`)

**Build:**
- `tsconfig.json` - TypeScript compilation config
  - Target: ESNext
  - Module: NodeNext
  - Output: `lib/` directory with source maps enabled
  - Module resolution: NodeNext

**Deployment:**
- Teams Framework YAML manifests:
  - `m365agents.yml` - Production deployment manifest
  - `m365agents.local.yml` - Local development manifest
  - `m365agents.playground.yml` - Playground environment manifest
- `appPackage/manifest.json` - Teams app manifest (Teams schema v1.25)
- `infra/azure.bicep` - Azure infrastructure as code (Bicep)
- `infra/azure.parameters.json` - ARM template parameters
- `web.config` - IIS configuration for web.config handler routing to `lib/src/index.js`

## Platform Requirements

**Development:**
- Node.js 20 or 22
- npm for dependency management
- TypeScript compiler
- Azure CLI (for local testing with Managed Identity)
- Teams Toolkit / Microsoft 365 development tools (for app provisioning)

**Production:**
- Azure App Service (B1 SKU specified in `infra/azure.parameters.json`)
- IIS with iisnode module (configured in `web.config`)
- Azure OpenAI Service deployment
- Microsoft Entra ID application registration
- Teams Bot Framework registration

**Deployment Pipeline:**
- Teams Toolkit CLI for provisioning and publishing
- ARM deployment for Azure resources
- Zip deployment to Azure App Service (via `azureAppService/zipDeploy`)

---

*Stack analysis: 2026-03-20*
