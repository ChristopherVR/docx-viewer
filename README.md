# Word document editor

A browser Word editor with one ProseMirror-backed `<docx-editor>` and thin React, Vue, Angular, Svelte and vanilla adapters. Rendering, selection, commands, history and styling live in the web component, so bindings share behavior.

This is an early implementation, not yet Microsoft Word feature or pagination parity.

## Packages and shared code

Install the umbrella package for the editor and framework entry points:

```sh
npm install @christophervr/docx-viewer
```

It installs the internal Word packages together. Import `@christophervr/docx-viewer` for the framework-neutral editor API or select a dedicated `@christophervr/docx-viewer/react`, `/vue`, `/angular`, `/svelte`, `/solid`, or `/vanilla` entry. The existing `@christophervr/docx-*` packages remain available for consumers who prefer explicit dependencies. Frameworks remain optional peers; install the framework used by your app.

- `@christophervr/docx-core`: modern DOCX model, import and preserving serialization. Its `/embedded` entry point is the planned PowerPoint integration once this package is published and the consumer migration is ready. PowerPoint's existing embedded-DOCX adapter remains in use for now.
- `@christophervr/docx-legacy`: legacy Word DOC adapter using `@christophervr/ole2`.
- `@christophervr/docx-document`: format detection and loading.
- `@christophervr/docx-web-component`: shared WYSIWYG editor.
- `@christophervr/docx-bindings`: framework lifecycle and event adapters.

`ole2` contains CFB/OLE2 and legacy binary codecs only. Modern DOCX belongs in this repository. The current release scope is `ole2`; do not publish or tag the Word packages yet. Package manifests use versioned npm dependencies; sibling checkouts are not required once the relevant packages have been published and consumer migrations are complete.

## Development

```sh
bun install
bun run demo
```

The default demo uses vanilla. Append `?framework=react`, `vue`, `angular`, `svelte` or `solid` to exercise another real adapter. It opens DOCX and DOC, saves in the imported format, and can export supported visible content as a new DOCX.

```sh
bun run typecheck
bun run test
bun run check:shared
bun run build:packages
bun run pack:smoke
bunx playwright install chromium
bun run test:browser
bun install --cwd docs
bun run --cwd docs docs:build
```

Browser test environment variables: `PLAYWRIGHT_CHROMIUM_EXECUTABLE` uses an explicit Chromium executable. If Playwright's pinned Chromium build is not installed, the config falls back to the newest `chromium_headless_shell-*` or `chromium-*` build under `PLAYWRIGHT_BROWSERS_PATH`. `PLAYWRIGHT_PORT` (default 4180) selects the preview-server port, which avoids reusing a stale server from another checkout.

The Pages workflow builds VitePress documentation and all six demos using the PowerPoint site's route structure. Deployment is enabled for public repositories. For a private repository, a GitHub plan supporting private-repository Pages and the `ENABLE_PRIVATE_PAGES=true` repository variable are required. Publishing instructions are in [the release guide](docs/releasing.md).

## Support today

| Capability                                                          | Status                                                                                    |
| ------------------------------------------------------------------- | ----------------------------------------------------------------------------------------- |
| DOCX paragraphs, direct bold/italic/underline/font/size/color       | Import, render, edit and export                                                           |
| Alignment and simple tables                                         | Import, edit, table insertion and export                                                  |
| Word-style ribbon, undo/redo, read-only, zoom                       | One shared web component                                                                  |
| Page size and margins                                               | Editable page surface; continuous content without automatic pagination                    |
| Untouched DOCX/DOC                                                  | Original bytes returned on no-op save                                                     |
| DOCX package parts                                                  | Preserved on supported edits; unsafe edits around unsupported inline content are rejected |
| Word 97-2003 DOC                                                    | Main-body text import and constrained existing-paragraph edits                            |
| Styles, numbering, fields, tracked changes, headers/footers, images | Not rendered with Word fidelity yet                                                       |
| Password-protected files                                            | Unsupported                                                                               |

DOC formatting and structural changes are rejected when saving DOC. **Export DOCX** creates a new document from supported visible content and cannot recover omitted layout or content. Assigning a replacement `documentModel` starts a new document; the save API preserves a loaded session.

## API

```ts
import { createDocument } from '@christophervr/docx-viewer/core';
import { mountEditor } from '@christophervr/docx-viewer';

const editor = mountEditor(container, {
	documentModel: createDocument(),
	onDocumentChange: (model) => console.log(model),
	onDocumentError: (error) => console.error(error),
});
await editor.load(bytes);
const saved = await editor.save();
editor.update({ readOnly: true });
editor.destroy();
```

Framework entry points `/react`, `/vue` and `/solid` export `WordEditor`; `/angular` exports `WordEditorComponent`; `/svelte` exports a default Svelte component. See [bindings](docs/bindings.md), [architecture](docs/architecture.md), [reuse audit](docs/reuse-audit.md) and [parity milestones](docs/parity-roadmap.md).
