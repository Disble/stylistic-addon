# Stylistic

A Microsoft Word add-in that analyzes Spanish-language documents and proposes editorial suggestions through Word's native **Track Changes** system. Users accept or reject each suggestion directly from the Review tab — no new interface to learn.

## Why Stylistic?

Writers and editors working in Spanish frequently encounter redundant phrases, filler words, and unnecessarily complex expressions. Stylistic leverages AI to detect these patterns and inserts corrections as tracked changes, integrating seamlessly into the existing Word review workflow.

**Key idea:** the add-in writes suggestions *as the document author would*, using Track Changes. The user reviews them with the same tools they already know.

## Architecture

Stylistic is a **frontend-only** Word add-in that communicates with a **Mastra backend** for AI-powered text analysis:

```
┌─────────────────────────────────────────────────┐
│  Word Add-in (Frontend)                         │
│                                                 │
│  taskpane.ts — UI orchestrator                  │
│  ├── wordApi.ts    — Read document, apply       │
│  │                   Track Changes via OOXML    │
│  ├── mastraClient.ts — Workflow execution       │
│  │                     via @mastra/client-js     │
│  └── chunker.ts    — Split large texts at       │
│                      paragraph boundaries       │
├─────────────────────────────────────────────────┤
│  Mastra Backend (separate application)          │
│  └── editorial-workflow — AI text analysis      │
└─────────────────────────────────────────────────┘
```

| Layer | Module | Responsibility | Imports Office.js? |
|---|---|---|---|
| UI | `taskpane.ts` | Event handling, progress, rendering | No (delegates) |
| API | `wordApi.ts` | Document read/write, OOXML tracked changes, comment cleanup | Yes (only module) |
| Backend | `mastraClient.ts` | Workflow execution, retry logic | No |
| Chunker | `chunker.ts` | Text splitting at paragraph boundaries | No |
| Config | `config.ts` | Internal constants and defaults | No |
| Types | `types.ts` | Shared interfaces | No |

See [docs/architecture.md](docs/architecture.md) for a detailed walkthrough.

## Features

- **AI-powered analysis** — detects redundancy, filler words, style issues, and more via a Mastra workflow
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
- Microsoft Word (desktop or Word Online)
- A running [Mastra](https://mastra.ai/) server with the `editorial-workflow` deployed

### Setup

```bash
cd stylistic-addon
npm install
```

### Configure Backend

The add-in expects a Mastra server at `http://localhost:4111` with the `editorial-workflow` registered. See [docs/api-contract.md](docs/api-contract.md) for the complete backend requirements.

### Run

```bash
npm start
```

This starts the dev server on `https://localhost:3000` and sideloads the add-in into Word desktop. In Word, go to **Home** tab and click **Show Task Pane**.

### Try It

1. Select an analysis profile from the dropdown (General, Formal, or Académico).
2. Click **"Analizar y sugerir"**.
3. Watch the progress bar as chunks are analyzed and suggestions are applied.
4. Open the **Review** tab to accept or reject tracked changes.
5. Click **"Limpiar comentarios resueltos"** to remove comments from resolved changes.

## Project Structure

```
stylistic-addon/
├── src/
│   ├── commands/
│   │   ├── commands.ts          # Ribbon command handler
│   │   └── commands.html
│   ├── lib/
│   │   ├── types.ts             # Shared interfaces (Suggestion, WorkflowInput, etc.)
│   │   ├── config.ts            # Constants (Mastra URL, retry policy)
│   │   ├── wordApi.ts           # Office.js abstraction (OOXML tracked changes + cleanup)
│   │   ├── mastraClient.ts      # Mastra workflow client with retry logic
│   │   └── chunker.ts           # Paragraph-boundary text splitting
│   └── taskpane/
│       ├── taskpane.ts          # UI orchestrator (multi-phase analysis pipeline)
│       ├── taskpane.html        # Task pane markup
│       └── taskpane.css         # Styles (Fluent UI compatible)
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
| `npm start` | Start dev server + sideload add-in in Word desktop |
| `npm stop` | Stop the dev server and unload the add-in |
| `npm run build` | Production build to `dist/` |
| `npm run build:dev` | Development build with source maps |
| `npm run watch` | Continuous rebuild on file changes |
| `npm run validate` | Validate `manifest.xml` against Office schema |
| `npm run lint` | Run ESLint on source files |
| `npm run lint:fix` | Auto-fix linting issues |

## Requirements

| Requirement | Minimum |
|---|---|
| WordApi | **1.6** |
| Office.js | 1.1 |
| Node.js | 18+ |
| Word Desktop | 2019+ / Microsoft 365 |
| Word Online | Supported (no version constraint) |
| Mastra Server | `@mastra/client-js` v1.7.1 compatible |

The add-in requires **WordApi 1.6** for `changeTrackingMode` and tracked change range comparison support. Word Online has full support regardless of client version.

## Dependencies

| Package | Purpose |
|---|---|
| `@mastra/client-js` | Mastra workflow client SDK for backend communication |
| `core-js` | Polyfills for older browser environments |
| `regenerator-runtime` | Async/await support for transpiled code |

## Documentation

| Document | Description |
|---|---|
| [Architecture](docs/architecture.md) | System design, data flow, and design decisions |
| [API Contract](docs/api-contract.md) | Backend requirements (Mastra workflow input/output) |
| [Adding Rules](docs/adding-rules.md) | How to extend editorial analysis rules |
| [Troubleshooting](docs/troubleshooting.md) | Common issues and their solutions |

## Known Limitations

- **Author attribution** — tracked changes appear under the current user's name, not the add-in's. This is a Word API limitation.
- **Co-authoring** — the add-in should not be used while multiple authors are actively editing the same document.
- **DRM/RMS protected documents** — protected documents cannot be modified by add-ins.
- **Backend required** — the add-in requires a running Mastra server with the editorial workflow deployed.
- **Language** — currently sends `"es"` (Spanish) as the language code. Multi-language support requires a language selector in the UI.

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/improvement`)
3. Read [docs/api-contract.md](docs/api-contract.md) if modifying backend integration
4. Run `npm run lint` and `npm run build` before committing
5. Open a Pull Request

## License

[MIT](LICENSE)
