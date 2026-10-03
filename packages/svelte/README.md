# docx-svelte-viewer

[![npm version](https://img.shields.io/npm/v/docx-svelte-viewer.svg)](https://www.npmjs.com/package/docx-svelte-viewer)
[![license](https://img.shields.io/npm/l/docx-svelte-viewer.svg)](https://github.com/ChristopherVR/docx-viewer/blob/main/LICENSE)
[![types](https://img.shields.io/npm/types/docx-svelte-viewer.svg)](https://www.npmjs.com/package/docx-svelte-viewer)

> A browser Word document editor for Svelte 5, using one shared editor and the canonical OOXML document engine.

[Live demo](https://christophervr.github.io/docx-viewer/demo/) | [npm](https://www.npmjs.com/package/docx-svelte-viewer) | [Full docs](https://christophervr.github.io/docx-viewer/) | [Source](https://github.com/ChristopherVR/docx-viewer)

## Install

```bash
npm install docx-svelte-viewer
```

Use the framework peers declared by this package: `svelte` (`^5.0.0`). Existing framework apps normally already provide them.

The package bundles its editor UI and adapter. The core model and shared controls
are installed as regular registry dependencies and the model API is re-exported.
One viewer package is enough; no separate core install is required.

## Quick start

```svelte
<script>
  import { createDocument } from 'docx-svelte-viewer/runtime';
  import WordEditor from 'docx-svelte-viewer';
  let model = $state(createDocument());
</script>
<WordEditor documentModel={model} ondocumentchange={next => model = next} />
```

Mount in the browser and give the host a height. Use one editor package per
application: each registers the same editor custom element.

## Features

| Feature    | Description                                                                     |
| ---------- | ------------------------------------------------------------------------------- |
| Editing    | Shared ProseMirror editor, ribbon, dialogs and document change events.          |
| Files      | Modern DOCX and limited legacy Word 97-2003 DOC loading through core.           |
| Model      | `createDocument`, `loadDocx`, `saveDocx` and core types from the package entry. |
| Frameworks | The same `<docx-editor>` UI behind six thin adapters.                           |

## API

The shared props, event handlers, `loadDocument` and `detectDocumentFormat` helpers
are available from the package. For an imperative browser integration,
`mountEditor(container, options)` returns a binding with `load`, `save` and
`destroy`. ProseMirror, `docx-core`, `ooxml-core` and `ooxml-ui` are regular
dependencies; document and layout algorithms live in OOXML core.

The root exports the Svelte component; JavaScript helpers and model exports are
under `docx-svelte-viewer/runtime`.

## Limitations

This is an early editor. It does not establish Microsoft Word layout parity or
lossless saving for unsupported features. Legacy DOC support is limited to
main-body text with constrained paragraph edits. SmartArt is display-only and
charts render as placeholders.

## Documentation

[Framework guide](https://christophervr.github.io/docx-viewer/frameworks/svelte) |
[Bindings](https://christophervr.github.io/docx-viewer/bindings)

## License

Apache-2.0.
