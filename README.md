<div align="center">

# docx-viewer

**A browser Word document editor with one document model, one web component and thin adapters for your framework.**
An early implementation: not Microsoft Word parity, and not lossless export.

[![docs](https://img.shields.io/badge/docs-christophervr.github.io-6366f1.svg)](https://christophervr.github.io/docx-viewer/)
[![license](https://img.shields.io/badge/license-Apache--2.0-blue.svg)](LICENSE)
[![CI](https://github.com/ChristopherVR/docx-viewer/actions/workflows/ci.yml/badge.svg)](https://github.com/ChristopherVR/docx-viewer/actions/workflows/ci.yml)

[**Live demo**](https://christophervr.github.io/docx-viewer/demo/) &nbsp;&middot;&nbsp;
[**Documentation**](https://christophervr.github.io/docx-viewer/) &nbsp;&middot;&nbsp;
[**Getting started**](#getting-started) &nbsp;&middot;&nbsp;
[**Packages**](#packages)

![The docx-viewer editor showing a sample document with the Word-style ribbon](https://raw.githubusercontent.com/ChristopherVR/docx-viewer/main/.github/assets/editor.png)

</div>

> **Not published to npm yet.** The Word packages are deliberately unpublished and untagged. Build from this repository to try them; the import paths below show the intended API.

## Why docx-viewer?

- **One editor, every framework.** A ProseMirror-backed `<docx-editor>` web component owns rendering, selection, commands, history and styling. React, Vue, Angular, Svelte, Solid and vanilla adapters only handle lifecycle and events.
- **One document model.** A framework-neutral DOCX model, parser and serializer sit under the editor.
- **Careful preservation.** A no-op save returns the original bytes, supported edits keep package parts outside the model, and unsafe edits are rejected instead of silently dropped.
- **Honest limits.** Unsupported features are reported, not hidden. See the [support roadmap](docs/parity-roadmap.md) and [outstanding work](docs/outstanding-work.md).

## Feature tour

|                        |                                                                                                                                                                         |
| ---------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Editor**             | Word-style ribbon, title bar, File backstage, status bar, undo/redo, find and replace, zoom and read-only mode in one web component.                                    |
| **Print Layout**       | Paginated view with line breaking, widow/orphan and keep rules, columns, table row splitting and printing. An approximation of Word, not Word's pagination.             |
| **Document structure** | Styles, lists and numbering, tables, sections and columns, headers and footers, footnotes and fields such as TOC. Inline text boxes, header watermark and page borders. |
| **Review**             | Tracked changes (accept, reject, navigate) and a comments pane with replies and resolve.                                                                                |
| **File formats**       | DOCX, and legacy Word 97-2003 `.doc` via the shared `ole2` codecs (main-body text and constrained paragraph edits). Password-protected files are unsupported.           |
| **Collaboration**      | A transport-neutral protocol over validated ProseMirror steps with transient presence. Your application owns networking, identity, permissions and storage.             |
| **Localization**       | Editor interface in English, French, German, Spanish and Simplified Chinese through a `locale` option. Document content is never translated.                            |

## Getting started

### 1. Build from source

```bash
git clone https://github.com/ChristopherVR/docx-viewer.git
cd docx-viewer
bun install
bun run demo   # vanilla demo; add ?framework=react|vue|angular|svelte|solid
```

Development resolves the private `@christophervr/ooxml-core` from a sibling checkout; see [AGENTS.md](AGENTS.md) and the [ooxml-core plan](docs/ooxml-core-plan.md).

### 2. Mount the editor (intended API)

```tsx
import { createDocument } from '@christophervr/docx-viewer/core';
import { WordEditor } from '@christophervr/docx-viewer/react';

export function Editor() {
	const [model, setModel] = useState(() => createDocument());
	return <WordEditor documentModel={model} onDocumentChange={setModel} />;
}
```

<details>
<summary><strong>Vue 3</strong></summary>

```vue
<script setup lang="ts">
import { WordEditor } from '@christophervr/docx-viewer/vue';
</script>
<template>
	<WordEditor :document-model="model" @document-change="model = $event" />
</template>
```

</details>

<details>
<summary><strong>Angular</strong></summary>

```ts
import { WordEditorComponent } from '@christophervr/docx-viewer/angular';

@Component({
	standalone: true,
	imports: [WordEditorComponent],
	template: '<word-editor [documentModel]="model" (documentChange)="model = $event" />',
})
export class EditorComponent {
	model = initialDocument;
}
```

</details>

<details>
<summary><strong>Svelte 5</strong></summary>

```svelte
<script>
  import WordEditor from '@christophervr/docx-viewer/svelte';
  let model = $state(initialDocument);
</script>
<WordEditor documentModel={model} ondocumentchange={next => model = next} />
```

</details>

<details>
<summary><strong>SolidJS</strong></summary>

```tsx
import { createSignal } from 'solid-js';
import { WordEditor } from '@christophervr/docx-viewer/solid';

const [model, setModel] = createSignal(initialDocument);
// <WordEditor documentModel={model()} onDocumentChange={setModel} />
```

</details>

<details>
<summary><strong>Vanilla JavaScript</strong></summary>

```ts
import { createDocument, mountEditor } from '@christophervr/docx-viewer';

const editor = mountEditor(container, {
	documentModel: createDocument(),
	onDocumentChange: (model) => console.log(model),
	onDocumentError: (error) => console.error(error),
});
await editor.load(bytes);
const saved = await editor.save();
editor.destroy();
```

</details>

See the [bindings guide](docs/bindings.md) for props, events, saving and file commands.

## Packages

| Package                             | What it is                                                                                                                        |
| ----------------------------------- | --------------------------------------------------------------------------------------------------------------------------------- |
| `@christophervr/docx-viewer`        | Umbrella entry point (`packages/viewer`): editor API plus `/react`, `/vue`, `/angular`, `/svelte`, `/solid`, `/vanilla` subpaths. |
| `@christophervr/docx-core`          | DOCX model, parser and preserving serializer (`packages/core`).                                                                   |
| `@christophervr/docx-document`      | Format detection and loading (`packages/document`).                                                                               |
| `@christophervr/docx-legacy`        | Legacy `.doc` adapter built on `@christophervr/ole2` (`packages/legacy`).                                                         |
| `@christophervr/docx-layout`        | Pagination engine behind Print Layout (`packages/layout`).                                                                        |
| `@christophervr/docx-web-component` | The shared `<docx-editor>` (`packages/web-component`).                                                                            |
| `@christophervr/docx-bindings`      | Framework lifecycle and event adapters (`packages/bindings`).                                                                     |

None of these are published yet. Roadmap: shared OOXML logic is moving into the private `@christophervr/ooxml-core`, leaving this repository with only the UI. That migration is in progress, not a shipped feature.

## Development

```bash
bun run typecheck       # tsc and svelte-check
bun run test            # Vitest
bun run check:shared    # shared-code boundary check
bun run build:packages
bun run pack:smoke
bunx playwright install chromium
bun run test:browser    # Playwright contract tests
bun run fmt:check       # oxfmt
bun install --cwd docs && bun run --cwd docs docs:build   # site and all demos
```

`PLAYWRIGHT_CHROMIUM_EXECUTABLE` selects an explicit Chromium and `PLAYWRIGHT_PORT` (default 4180) the preview port. Publishing is intentionally out of scope for now; see the [release policy](docs/releasing.md).

## Documentation

[Architecture](docs/architecture.md) &middot; [Framework bindings](docs/bindings.md) &middot; [Editing text](docs/editing.md) &middot; [Collaboration](docs/collaboration.md) &middot; [Support roadmap](docs/parity-roadmap.md) &middot; [Outstanding work](docs/outstanding-work.md) &middot; [ooxml-core plan](docs/ooxml-core-plan.md) &middot; [Reuse audit](docs/reuse-audit.md)

## License

[Apache License 2.0](LICENSE). See [`NOTICE`](NOTICE) for attributions.
