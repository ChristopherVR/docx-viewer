# docx-vanilla-viewer

[![npm version](https://img.shields.io/npm/v/docx-vanilla-viewer.svg)](https://www.npmjs.com/package/docx-vanilla-viewer)
[![license](https://img.shields.io/npm/l/docx-vanilla-viewer.svg)](https://github.com/ChristopherVR/docx-viewer/blob/main/LICENSE)
[![types](https://img.shields.io/npm/types/docx-vanilla-viewer.svg)](https://www.npmjs.com/package/docx-vanilla-viewer)

> A browser Word document editor for plain JavaScript, using one shared editor and the canonical OOXML document engine.

[Live demo](https://christophervr.github.io/docx-viewer/demo/) | [npm](https://www.npmjs.com/package/docx-vanilla-viewer) | [Full docs](https://christophervr.github.io/docx-viewer/) | [Source](https://github.com/ChristopherVR/docx-viewer)

## Install

```bash
npm install docx-vanilla-viewer
```

The package bundles its editor UI and adapter. The core model and shared controls
are installed as regular registry dependencies and the model API is re-exported.
One viewer package is enough; no separate core install is required.

## Quick start

```ts
import { createDocument, mountEditor } from 'docx-vanilla-viewer';

const container = document.createElement('div');
container.style.height = '600px';
document.body.append(container);
const editor = mountEditor(container, {
	documentModel: createDocument(),
	onDocumentChange: (model) => console.log(model),
	onDocumentError: (error) => console.error(error),
});
const bytes = new Uint8Array(await (await fetch('/example.docx')).arrayBuffer());
await editor.load(bytes); // DOCX, or a legacy Word 97-2003 .doc
const saved = await editor.save();
editor.destroy();
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

## Limitations

This is an early editor. It does not establish Microsoft Word layout parity or
lossless saving for unsupported features. Legacy DOC support is limited to
main-body text with constrained paragraph edits. SmartArt is display-only and
charts render as placeholders.

## Documentation

[Framework guide](https://christophervr.github.io/docx-viewer/frameworks/vanilla) |
[Bindings](https://christophervr.github.io/docx-viewer/bindings)

## License

Apache-2.0.
