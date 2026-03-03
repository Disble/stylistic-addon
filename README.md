# Stylistic

A Microsoft Word add-in that analyzes Spanish-language documents and proposes editorial suggestions through Word's native **Track Changes** system. Users accept or reject each suggestion directly from the Review tab — no new interface to learn.

## Why Stylistic?

Writers and editors working in Spanish frequently encounter redundant phrases, filler words, and unnecessarily complex expressions. Stylistic detects these patterns and inserts corrections as tracked changes, integrating seamlessly into the existing Word review workflow.

**Key idea:** the add-in writes suggestions *as the document author would*, using Track Changes. The user reviews them with the same tools they already know.

## Features

- **Redundancy detection** — flags pleonasms like *"completamente necesario"*, *"subir arriba"*, *"periodo de tiempo"*
- **Filler word removal** — identifies weakening phrases like *"básicamente"*, *"obviamente"*, *"en realidad"*
- **Word choice simplification** — suggests *"usar"* over *"utilizar"*, *"para"* over *"con el objetivo de"*, and more
- **Native Track Changes** — suggestions appear as real Word revisions (strikethrough + underline), reviewable from the Review tab
- **Non-destructive** — preserves the document's original tracking mode after analysis (preserve-and-restore pattern)
- **Extensible rule engine** — adding new editorial rules requires only a pattern declaration, no framework code

## Quick Start

### Prerequisites

- [Node.js](https://nodejs.org/) v18+
- Microsoft Word (desktop or Word Online)
- A code editor

### Setup

```bash
cd stylistic-addon
npm install
```

### Run

```bash
npm start
```

This starts the dev server on `https://localhost:3000` and sideloads the add-in into Word desktop. In Word, go to **Home** tab and click **Show Task Pane**.

### Try it

Paste this text into a Word document:

> Básicamente, es completamente necesario utilizar este periodo de tiempo con el objetivo de realizar la tarea. Obviamente, el resultado final será muy único.

Click **"Analizar y sugerir"**. The add-in will insert tracked changes for each detected issue. Open the **Review** tab to accept or reject them.

## Architecture

Stylistic follows a strict three-layer separation of concerns:

```
┌─────────────────────────────────────────┐
│  taskpane.ts — UI orchestration         │
│  Binds events, delegates, renders       │
├─────────────────────────────────────────┤
│  wordApi.ts — Word API abstraction      │  analyzer.ts — Business logic
│  All Office.js calls live here          │  Pure functions, zero Office.js
│  Preserve-and-restore tracking mode     │  Pattern-based rule engine
├─────────────────────────────────────────┤
│  types.ts — Shared interfaces (Suggestion, InsertionResult)               │
└─────────────────────────────────────────┘
```

| Layer | Module | Responsibility | Imports Office.js? |
|---|---|---|---|
| UI | `taskpane.ts` | Event handling, rendering, loading states | No (delegates) |
| API | `wordApi.ts` | Document read/write, track changes | Yes (only module) |
| Logic | `analyzer.ts` | Text analysis, suggestion generation | No |
| Types | `types.ts` | Shared interfaces | No |

See [docs/architecture.md](docs/architecture.md) for a detailed walkthrough.

## Project Structure

```
stylistic-addon/
├── src/
│   ├── commands/
│   │   ├── commands.ts          # Ribbon command handler
│   │   └── commands.html
│   ├── lib/
│   │   ├── types.ts             # Suggestion, InsertionResult interfaces
│   │   ├── wordApi.ts           # Office.js abstraction layer
│   │   └── analyzer.ts          # Rule-based text analysis engine
│   └── taskpane/
│       ├── taskpane.ts          # UI orchestrator
│       ├── taskpane.html        # Task pane markup
│       └── taskpane.css         # Styles (Fluent UI compatible)
├── assets/                      # Add-in icons (16, 32, 64, 80, 128px)
├── manifest.xml                 # Office add-in manifest (WordApi 1.6)
├── webpack.config.js            # Webpack config (dev + production)
├── tsconfig.json                # TypeScript configuration
├── package.json                 # Dependencies and scripts
└── docs/                        # Extended documentation
    ├── architecture.md
    ├── adding-rules.md
    └── troubleshooting.md
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

The add-in requires **WordApi 1.6** for `changeTrackingMode` support. Word Online has full support regardless of client version.

## Documentation

| Document | Description |
|---|---|
| [Architecture](docs/architecture.md) | Layered design, data flow, and design decisions |
| [Adding Rules](docs/adding-rules.md) | Step-by-step guide to creating new editorial rules |
| [Troubleshooting](docs/troubleshooting.md) | Common issues and their solutions |

## Known Limitations

- **Author attribution** — tracked changes appear under the current user's name, not the add-in's. This is a Word API limitation and an accepted non-requirement.
- **Co-authoring** — the add-in should not be used while multiple authors are actively editing the same document.
- **DRM/RMS protected documents** — protected documents cannot be modified by add-ins.
- **Performance** — documents with 50+ suggestions may experience slower processing. Future versions will batch in groups of 10-15.
- **Language** — currently supports Spanish text analysis only.

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/new-rule`)
3. Read [docs/adding-rules.md](docs/adding-rules.md) if adding editorial rules
4. Run `npm run lint` and `npm run build` before committing
5. Open a Pull Request

## License

[MIT](LICENSE)
