# Stylistic

A Microsoft Word add-in that analyzes Spanish-language documents and proposes editorial suggestions through Word's native **Track Changes** system. Users accept or reject each suggestion directly from the Review tab — no new interface to learn.

## Why Stylistic?

Writers and editors working in Spanish frequently encounter redundant phrases, filler words, and unnecessarily complex expressions. Stylistic leverages AI to detect these patterns and inserts corrections as tracked changes, integrating seamlessly into the existing Word review workflow.

**Key idea:** the add-in writes suggestions *as the document author would*, using Track Changes. The user reviews them with the same tools they already know.

## Architecture

Stylistic is a Word add-in with a hexagonal frontend boundary. The add-in owns
Word-host orchestration, auth state, taskpane presentation, and document
mutation. A separate Mastra backend owns authentication and AI-powered analysis:

```
┌─────────────────────────────────────────────────┐
│  Word Add-in                                    │
│                                                 │
│  taskpane.ts — composition root                 │
│  ├── adapters/word — Office.js boundary         │
│  ├── adapters/auth — Better Auth + Dialog API   │
│  ├── adapters/mastra — @mastra/client-js        │
│  ├── domain — pure ports, types, workflows      │
│  └── taskpane/components — React + Fluent UI    │
├─────────────────────────────────────────────────┤
│  Mastra Backend                                 │
│  ├── Better Auth / Google OAuth                 │
│  └── stylistic-workflow — AI text analysis      │
└─────────────────────────────────────────────────┘
```

| Layer | Module | Responsibility | Imports Office.js? |
|---|---|---|---|
| Composition | `src/taskpane/taskpane.ts` | Wires ports/adapters and exposes UI handlers | No |
| Presentation | `src/taskpane/**` | React shell, Fluent UI, Zustand presentation state | No |
| Auth adapters | `src/adapters/auth/**` | Better Auth client, Office Dialog flow, OfficeRuntime storage | Uses Office Dialog/OfficeRuntime only |
| Word adapters | `src/adapters/word/**` | Document read/write, Track Changes, comments, cleanup | Yes |
| Backend adapters | `src/adapters/mastra/**` | Workflow execution and feedback through authenticated Mastra clients | No |
| Domain | `src/domain/**` | Ports, pure types, pipeline, review workflow | No |
| Infrastructure | `src/infrastructure/config.ts` | URLs, workflow IDs, auth constants, retry policy | No |

See [docs/architecture.md](docs/architecture.md) for a detailed walkthrough.

## Features

- **AI-powered analysis** — detects redundancy, filler words, style issues, and more via a Mastra workflow
- **Authenticated backend access** — users sign in with Google before protected Mastra workflow calls
- **Native Track Changes** — suggestions appear as real Word revisions (strikethrough + underline), reviewable from the Review tab
- **Justification comments** — each tracked change includes a Word comment with the editorial category and reason
- **Comment cleanup** — a "Limpiar comentarios resueltos" button removes orphaned comments after the user accepts/rejects tracked changes
- **Large document support** — handles documents with 200,000+ words through paragraph-boundary chunking
- **Severity levels** — suggestions include severity (high/medium/low) for prioritization
- **Reliability first** — retry logic with exponential backoff, partial success reporting, preserve-and-restore tracking mode
- **Non-destructive** — preserves the document's original tracking mode after analysis
- **Simple UI** — one-click analysis with a profile dropdown, progress bar, and results panel
- **Extensible** — new analysis rules require only backend prompt changes, no frontend modifications

## Quick Start

### Prerequisites

- [Node.js](https://nodejs.org/) v18+
- [Bun](https://bun.sh/) for dependency and script execution
- Microsoft Word (desktop or Word Online)
- A running [Mastra](https://mastra.ai/) server with Better Auth configured and the `stylistic-workflow` registered

### Setup

```bash
cd stylistic-addon
bun install
```

### Configure Backend

The add-in expects a Mastra server at `http://localhost:4111` with:

- Better Auth mounted at `/auth/*`,
- Google OAuth configured,
- `https://localhost:3000` listed in Better Auth trusted origins,
- the `stylistic-workflow` registered and protected by bearer auth.

See [docs/api-contract.md](docs/api-contract.md) for workflow requirements and the backend [`docs/auth.md`](../stylistics-backend/docs/auth.md) for auth setup.

### Run

```bash
bun run start
```

This starts the dev server on `https://localhost:3000` and sideloads the add-in into Word desktop. In Word, go to **Home** tab and click **Show Task Pane**.

### Try It

1. Click **"Continuar con Google"** and complete login in the Office Dialog.
2. Select an analysis profile from the dropdown (General, Formal, or Académico).
3. Click **"Analizar y sugerir"**.
4. Watch the progress bar as chunks are analyzed and suggestions are applied.
5. Open the **Review** tab to accept or reject tracked changes.
6. Click **"Limpiar comentarios resueltos"** to remove comments from resolved changes.

## Project Structure

```
stylistic-addon/
├── src/
│   ├── commands/
│   │   ├── commands.ts          # Ribbon command handler
│   │   └── commands.html
│   ├── domain/
│   │   ├── ports.ts             # Hexagonal ports, including auth/session ports
│   │   ├── auth/                # Auth session contracts
│   │   └── pipeline/            # Analysis pipeline handlers/state
│   ├── adapters/
│   │   ├── auth/                # Better Auth, Office Dialog, OfficeRuntime storage
│   │   ├── mastra/              # Authenticated Mastra workflow clients
│   │   └── word/                # Office.js document boundary
│   ├── infrastructure/
│   │   ├── config.ts            # Mastra/Auth URLs, workflow IDs, retry policy
│   │   └── chunker.ts           # Paragraph-boundary text splitting
│   └── taskpane/
│       ├── taskpane.ts          # Composition root and top-level handlers
│       ├── index.tsx            # React + Fluent UI bootstrap
│       ├── auth-dialog.html     # Office Dialog OAuth entry page
│       ├── auth-dialog.ts       # OAuth dialog bridge runtime
│       └── components/          # Fluent UI taskpane components
├── assets/                      # Add-in icons (16, 32, 64, 80, 128px)
├── manifest.xml                 # Office add-in manifest (WordApi 1.6)
├── webpack.config.js            # Webpack config (dev + production)
├── tsconfig.json                # TypeScript configuration
├── package.json                 # Dependencies and scripts
└── docs/                        # Extended documentation
    ├── architecture.md          # System design, data flow, decisions
    ├── api-contract.md          # Backend requirements (Mastra workflow I/O)
    ├── adding-rules.md          # How to extend analysis rules
    └── troubleshooting.md       # Common issues and solutions
```

## Available Scripts

| Command | Description |
|---|---|
| `bun run start` | Start dev server + sideload add-in in Word desktop |
| `bun run stop` | Stop the dev server and unload the add-in |
| `bun run dev-server` | Start only the webpack dev server |
| `bun run watch` | Continuous development rebuild on file changes |
| `bun run typecheck` | Run TypeScript without emitting files |
| `bun run lint` | Run Office add-in lint checks |
| `bun run validate` | Run lint, architecture rails, filename, complexity, React rails, and typecheck |

## Requirements

| Requirement | Minimum |
|---|---|
| WordApi | **1.6** |
| Office.js | 1.1 |
| Node.js | 18+ |
| Word Desktop | 2019+ / Microsoft 365 |
| Word Online | Supported (no version constraint) |
| Mastra Server | `@mastra/client-js` v1.7.1 compatible, Better Auth enabled |

The add-in requires **WordApi 1.6** for `changeTrackingMode` and tracked change range comparison support. Word Online has full support regardless of client version.

## Dependencies

| Package | Purpose |
|---|---|
| `@mastra/client-js` | Mastra workflow client SDK for backend communication |
| `better-auth` | Client-side auth proxy for Google sign-in/session calls |
| `@fluentui/react-components` | Taskpane UI components |
| `zustand` | React-owned taskpane presentation state |
| `core-js` | Polyfills for older browser environments |
| `regenerator-runtime` | Async/await support for transpiled code |

## Documentation

| Document | Description |
|---|---|
| [Architecture](docs/architecture.md) | System design, data flow, and design decisions |
| [API Contract](docs/api-contract.md) | Backend requirements (Mastra workflow input/output) |
| [Troubleshooting](docs/troubleshooting.md) | Common auth, backend, and Word-host issues |
| [Adding Rules](docs/adding-rules.md) | How to extend editorial analysis rules |

## Known Limitations

- **Author attribution** — tracked changes appear under the current user's name, not the add-in's. This is a Word API limitation.
- **Co-authoring** — the add-in should not be used while multiple authors are actively editing the same document.
- **DRM/RMS protected documents** — protected documents cannot be modified by add-ins.
- **Backend required** — the add-in requires a running Mastra server with Better Auth and the stylistic workflow deployed.
- **Auth required** — analysis and feedback require a valid Better Auth bearer session. Logout clears local OfficeRuntime storage.
- **Language** — currently sends `"es"` (Spanish) as the language code. Multi-language support requires a language selector in the UI.

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/improvement`)
3. Read [docs/api-contract.md](docs/api-contract.md) if modifying backend integration
4. Run `bun run validate` before committing
5. Open a Pull Request

## License

[MIT](LICENSE)
