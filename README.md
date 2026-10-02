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

> **Not published to npm yet.** Nothing has been released or tagged. Build from this repository to try the editor; the install commands and import paths below are the intended API of the packages the release workflow will publish.

## Why docx-viewer?

- **One editor, every framework.** A ProseMirror-backed `<docx-editor>` web component owns rendering, selection, commands, history and styling. React, Vue, Angular, Svelte, Solid and vanilla adapters only handle lifecycle and events.
- **One document model.** A framework-neutral DOCX model, parser and serializer sit under the editor.
- **Careful preservation.** A no-op save returns the original bytes, supported edits keep package parts outside the model, and unsafe edits are rejected instead of silently dropped.
- **Honest limits.** Unsupported features are reported, not hidden. See the [support roadmap](docs/parity-roadmap.md) and [outstanding work](docs/outstanding-work.md).

## Feature tour

|                        |                                                                                                                                                                                                                                  |
| ---------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Editor**             | Word-style ribbon, title bar, File backstage, status bar, undo/redo, find and replace, zoom and read-only mode in one web component.                                                                                             |
| **Print Layout**       | Paginated view with line breaking, widow/orphan and keep rules, columns, table row splitting and printing. An approximation of Word, not Word's pagination.                                                                      |
| **Document structure** | Styles, lists and numbering, tables, sections and columns, headers and footers, footnotes and fields such as TOC. Inline text boxes, header watermark and page borders. SmartArt shows its saved drawing, read-only (see below). |
| **Review**             | Tracked changes (accept, reject, navigate) and a comments pane with replies and resolve.                                                                                                                                         |
| **File formats**       | DOCX, and legacy Word 97-2003 `.doc` via the shared `ole2` codecs inside `@christophervr/ooxml-core` (main-body text and constrained paragraph edits). Password-protected files are unsupported.                                 |
| **Collaboration**      | A transport-neutral protocol over validated ProseMirror steps with transient presence. Your application owns networking, identity, permissions and storage.                                                                      |
| **Localization**       | Editor interface in English, French, German, Spanish and Simplified Chinese through a `locale` option. Document content is never translated.                                                                                     |

## Getting started

### 1. Build from source

```bash
git clone https://github.com/ChristopherVR/docx-viewer.git
cd docx-viewer
bun install
bun run demo   # vanilla demo; add ?framework=react|vue|angular|svelte|solid
```

The DOCX logic comes from the published [`@christophervr/ooxml-core`](https://github.com/ChristopherVR/ooxml-core) package; see [AGENTS.md](AGENTS.md) and the [ooxml-core plan](docs/ooxml-core-plan.md).

### 2. Mount the editor (intended API)

Install one self-contained editor package for your framework, for example `npm install @christophervr/docx-react-viewer react`. The editor packages are `@christophervr/docx-react-viewer`, `-vue-viewer`, `-angular-viewer`, `-svelte-viewer`, `-solid-viewer` and `-vanilla-viewer`. Each one bundles the editor, layout engine and legacy `.doc` reader, brings the document model (`@christophervr/docx-core`) with it, and re-exports it, so a single install and a single import path are all an application needs. Install `@christophervr/docx-core` on its own only for headless use (parsing and serializing without an editor).

```tsx
import { createDocument, WordEditor } from '@christophervr/docx-react-viewer';

export function Editor() {
	const [model, setModel] = useState(() => createDocument());
	return <WordEditor documentModel={model} onDocumentChange={setModel} />;
}
```

<details>
<summary><strong>Vue 3</strong></summary>

```vue
<script setup lang="ts">
import { ref } from 'vue';
import { createDocument, WordEditor } from '@christophervr/docx-vue-viewer';

const model = ref(createDocument());
</script>
<template>
	<WordEditor :document-model="model" @document-change="model = $event" />
</template>
```

</details>

<details>
<summary><strong>Angular</strong></summary>

```ts
import { createDocument, WordEditorComponent } from '@christophervr/docx-angular-viewer';

@Component({
	standalone: true,
	imports: [WordEditorComponent],
	template: '<word-editor [documentModel]="model" (documentChange)="model = $event" />',
})
export class EditorComponent {
	model = createDocument();
}
```

</details>

<details>
<summary><strong>Svelte 5</strong></summary>

```svelte
<script>
  import WordEditor from '@christophervr/docx-svelte-viewer';
  import { createDocument } from '@christophervr/docx-svelte-viewer/runtime';
  let model = $state(createDocument());
</script>
<WordEditor documentModel={model} ondocumentchange={next => model = next} />
```

</details>

<details>
<summary><strong>SolidJS</strong></summary>

```tsx
import { createSignal } from 'solid-js';
import { createDocument, WordEditor } from '@christophervr/docx-solid-viewer';

export function Editor() {
	const [model, setModel] = createSignal(createDocument());
	return <WordEditor documentModel={model()} onDocumentChange={setModel} />;
}
```

</details>

<details>
<summary><strong>Vanilla JavaScript</strong></summary>

```ts
import { createDocument, mountEditor } from '@christophervr/docx-vanilla-viewer';

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

Seven packages are published (nothing has been released yet), each versioned independently:

| Package                              | What it is                                                                                                                           |
| ------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------ |
| `@christophervr/docx-core`           | DOCX model, parser and preserving serializer (`packages/core`, a thin entry point over `@christophervr/ooxml-core/docx`). DOCX only. |
| `@christophervr/docx-react-viewer`   | React component (`packages/react`).                                                                                                  |
| `@christophervr/docx-vue-viewer`     | Vue 3 component (`packages/vue`).                                                                                                    |
| `@christophervr/docx-angular-viewer` | Angular standalone component (`packages/angular`).                                                                                   |
| `@christophervr/docx-svelte-viewer`  | Svelte 5 component (`packages/svelte`).                                                                                              |
| `@christophervr/docx-solid-viewer`   | Solid component (`packages/solid`).                                                                                                  |
| `@christophervr/docx-vanilla-viewer` | `mountEditor` and the plain `<docx-editor>` web component, no framework (`packages/vanilla`).                                        |

Each `*-viewer` package is self-contained: the editor, layout engine, document loading and the legacy `.doc` reader come from `@christophervr/ooxml-core` (`/docx/layout`, `/docx/load`, which inlines the shared ole2 codecs). A package depends on `@christophervr/docx-core`, `@christophervr/ooxml-core`, `@christophervr/office-ui` (the shared web components that draw SmartArt; installed for you, never imported by your code), the ProseMirror libraries and its framework peer, so `@christophervr/ole2` does not need to be published or installed. DOCX and legacy `.doc` files are both opened by every editor package.

The workspace also holds **private** packages that are never published and are inlined into each editor package at build time: `web-component` (the shared `<docx-editor>`, including the browser text measurer behind Print Layout) and `bindings` (framework lifecycle and event adapters). The pagination engine, document loading and collaboration helpers are in `@christophervr/ooxml-core`. Shared OOXML logic is moving into the public `@christophervr/ooxml-core`, leaving this repository with only the UI; that migration is in progress, not a shipped feature.

## SmartArt

A SmartArt graphic (inline or floating) is displayed from the drawing the producing application saved in the file (`dsp:drawing`), drawn as SVG by `<office-ui-smartart>`. It is **read-only**: the diagram is not editable, its layout is **not recomputed**, and it is not claimed to look like Word. Preset outlines the renderer cannot draw appear as rectangles, gradient and pattern fills are approximated, transparency and 3D are dropped, and text is one line per shape without autofit; each case is listed in the diagram's info popover (the small `i` button) together with the file's own notices. A diagram with no saved drawing shows a labelled box listing its node text. Its parts and markup are preserved byte for byte on save, including when other text is edited. The accessible name comes from the diagram's alt text, title or object name.

## Development

```bash
bun run typecheck       # tsc and svelte-check
bun run test            # Vitest
bun run check:shared    # shared-code boundary check
bun run build:packages
bun run check:published   # no tarball may import an internal package or ole2
bun run pack:smoke
bunx playwright install chromium
bun run test:browser    # Playwright contract tests
bun run fmt:check       # oxfmt
bun install --cwd docs && bun run --cwd docs docs:build   # site and all demos
```

`PLAYWRIGHT_CHROMIUM_EXECUTABLE` selects an explicit Chromium and `PLAYWRIGHT_PORT` (default 4180) the preview port. Publishing is not enabled yet; see the [release policy](docs/releasing.md).

## Documentation

[Architecture](docs/architecture.md) &middot; [Framework bindings](docs/bindings.md) &middot; [Editing text](docs/editing.md) &middot; [Collaboration](docs/collaboration.md) &middot; [Support roadmap](docs/parity-roadmap.md) &middot; [Outstanding work](docs/outstanding-work.md) &middot; [ooxml-core plan](docs/ooxml-core-plan.md) &middot; [Reuse audit](docs/reuse-audit.md)

## License

[Apache License 2.0](LICENSE). See [`NOTICE`](NOTICE) for attributions.
