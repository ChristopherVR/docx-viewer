# @christophervr/docx-svelte-viewer

The DOCX editor for Svelte 5. One self-contained package: the editor web component, the layout engine, the legacy Word 97-2003 `.doc` reader and the Svelte 5 adapter are bundled in, so you install nothing else from this project except `@christophervr/docx-core` (the document model, `createDocument`, `loadDocx` and `saveDocx`).

```sh
npm install @christophervr/docx-svelte-viewer @christophervr/docx-core svelte
```

```svelte
<script>
  import { createDocument } from '@christophervr/docx-core';
  import WordEditor from '@christophervr/docx-svelte-viewer';
  let model = $state(createDocument());
</script>
<WordEditor documentModel={model} ondocumentchange={next => model = next} />
```

The package root is the component (a default export). The plain-JavaScript helpers (`mountEditor`, `loadDocument`, shared types) are under `@christophervr/docx-svelte-viewer/runtime`.

Notes:

- DOCX and legacy `.doc` files are both opened; the format is sniffed from the bytes, never from the file name. Legacy `.doc` support is limited to main-body text with constrained paragraph edits.
- The ProseMirror libraries are regular dependencies so your package manager can dedupe them with other ProseMirror users. svelte is a peer dependency.
- Use one editor package per application: each bundles its own copy of the editor and registers the `<docx-editor>` element.
- An early editor: not Microsoft Word layout parity, and saving is not lossless for unsupported features. See the [support roadmap](https://github.com/ChristopherVR/docx-viewer/blob/main/docs/parity-roadmap.md).

Documentation: [bindings guide](https://github.com/ChristopherVR/docx-viewer/blob/main/docs/bindings.md), [Svelte 5 guide](https://github.com/ChristopherVR/docx-viewer/blob/main/docs/frameworks/svelte.md). Licensed under Apache-2.0.
