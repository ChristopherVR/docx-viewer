# docx-vanilla-viewer

The DOCX editor for plain JavaScript. One self-contained package: the editor web component, the layout engine, the legacy Word 97-2003 `.doc` reader and the plain JavaScript adapter are bundled in, so you install nothing else from this project except `docx-core` (the document model, `createDocument`, `loadDocx` and `saveDocx`).

```sh
npm install docx-vanilla-viewer docx-core
```

```ts
import { createDocument } from 'docx-core';
import { mountEditor } from 'docx-vanilla-viewer';

const editor = mountEditor(container, {
	documentModel: createDocument(),
	onDocumentChange: (model) => console.log(model),
	onDocumentError: (error) => console.error(error),
});
await editor.load(bytes); // DOCX, or a legacy Word 97-2003 .doc
const saved = await editor.save();
editor.destroy();
```

This package also owns the plain `<docx-editor>` custom element: `registerDocxEditor()` defines it, and `mountEditor()` registers and mounts it for you. The element, the shared option types and the `loadDocument` / `detectDocumentFormat` helpers are exported from the package root.

Notes:

- DOCX and legacy `.doc` files are both opened; the format is sniffed from the bytes, never from the file name. Legacy `.doc` support is limited to main-body text with constrained paragraph edits.
- The ProseMirror libraries are regular dependencies so your package manager can dedupe them with other ProseMirror users.
- Use one editor package per application: each bundles its own copy of the editor and registers the `<docx-editor>` element.
- An early editor: not Microsoft Word layout parity, and saving is not lossless for unsupported features. See the [support roadmap](https://github.com/ChristopherVR/docx-viewer/blob/main/docs/parity-roadmap.md).

Documentation: [bindings guide](https://github.com/ChristopherVR/docx-viewer/blob/main/docs/bindings.md), [plain JavaScript guide](https://github.com/ChristopherVR/docx-viewer/blob/main/docs/frameworks/vanilla.md). Licensed under Apache-2.0.
