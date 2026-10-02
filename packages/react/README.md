# docx-react-viewer

The DOCX editor for React. One self-contained package: the editor web component, the layout engine, the legacy Word 97-2003 `.doc` reader and the React adapter are bundled in, so you install nothing else from this project except `docx-core` (the document model, `createDocument`, `loadDocx` and `saveDocx`).

```sh
npm install docx-react-viewer docx-core react
```

```tsx
import { useState } from 'react';
import { createDocument } from 'docx-core';
import { WordEditor } from 'docx-react-viewer';

export function Editor() {
	const [model, setModel] = useState(() => createDocument());
	return <WordEditor documentModel={model} onDocumentChange={setModel} />;
}
```

The shared option types and the `loadDocument` / `detectDocumentFormat` helpers are exported from the package root.

Notes:

- DOCX and legacy `.doc` files are both opened; the format is sniffed from the bytes, never from the file name. Legacy `.doc` support is limited to main-body text with constrained paragraph edits.
- The ProseMirror libraries are regular dependencies so your package manager can dedupe them with other ProseMirror users. react is a peer dependency.
- Use one editor package per application: each bundles its own copy of the editor and registers the `<docx-editor>` element.
- An early editor: not Microsoft Word layout parity, and saving is not lossless for unsupported features. See the [support roadmap](https://github.com/ChristopherVR/docx-viewer/blob/main/docs/parity-roadmap.md).

Documentation: [bindings guide](https://github.com/ChristopherVR/docx-viewer/blob/main/docs/bindings.md), [React guide](https://github.com/ChristopherVR/docx-viewer/blob/main/docs/frameworks/react.md). Licensed under Apache-2.0.
