# docx-angular-viewer

The DOCX editor for Angular. One self-contained package: the editor web component, the layout engine, the legacy Word 97-2003 `.doc` reader and the Angular adapter are bundled in, so you install nothing else from this project except `docx-core` (the document model, `createDocument`, `loadDocx` and `saveDocx`).

```sh
npm install docx-angular-viewer docx-core @angular/core
```

```ts
import { Component } from '@angular/core';
import { createDocument } from 'docx-core';
import { WordEditorComponent } from 'docx-angular-viewer';

@Component({
	standalone: true,
	imports: [WordEditorComponent],
	template: '<word-editor [documentModel]="model" (documentChange)="model = $event" />',
})
export class EditorComponent {
	model = createDocument();
}
```

The shared option types and the `loadDocument` / `detectDocumentFormat` helpers are exported from the package root.

Notes:

- DOCX and legacy `.doc` files are both opened; the format is sniffed from the bytes, never from the file name. Legacy `.doc` support is limited to main-body text with constrained paragraph edits.
- The ProseMirror libraries are regular dependencies so your package manager can dedupe them with other ProseMirror users. @angular/core is a peer dependency.
- Use one editor package per application: each bundles its own copy of the editor and registers the `<docx-editor>` element.
- An early editor: not Microsoft Word layout parity, and saving is not lossless for unsupported features. See the [support roadmap](https://github.com/ChristopherVR/docx-viewer/blob/main/docs/parity-roadmap.md).

Documentation: [bindings guide](https://github.com/ChristopherVR/docx-viewer/blob/main/docs/bindings.md), [Angular guide](https://github.com/ChristopherVR/docx-viewer/blob/main/docs/frameworks/angular.md). Licensed under Apache-2.0.
