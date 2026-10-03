# docx-vue-viewer

[![npm version](https://img.shields.io/npm/v/docx-vue-viewer.svg)](https://www.npmjs.com/package/docx-vue-viewer)
[![license](https://img.shields.io/npm/l/docx-vue-viewer.svg)](https://github.com/ChristopherVR/docx-viewer/blob/main/LICENSE)
[![types](https://img.shields.io/npm/types/docx-vue-viewer.svg)](https://www.npmjs.com/package/docx-vue-viewer)

> A browser Word document editor for Vue 3, using one shared editor and the canonical OOXML document engine.

[Live demo](https://christophervr.github.io/docx-viewer/demo/) | [npm](https://www.npmjs.com/package/docx-vue-viewer) | [Full docs](https://christophervr.github.io/docx-viewer/) | [Source](https://github.com/ChristopherVR/docx-viewer)

## Install

```bash
npm install docx-vue-viewer
```

Use the framework peers declared by this package: `vue` (`^3.5.0`). Existing framework apps normally already provide them.

The package bundles its editor UI and adapter. The core model and shared controls
are installed as regular registry dependencies and the model API is re-exported.
One viewer package is enough; no separate core install is required.

## Quick start

```vue
<script setup lang="ts">
import { shallowRef } from 'vue';
import { createDocument, WordEditor } from 'docx-vue-viewer';

const model = shallowRef(createDocument());
</script>
<template>
	<WordEditor :document-model="model" @document-change="model = $event" />
</template>
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
are available from the package. Framework handles expose `load`, `save` and
`download`. The vanilla package supplies `mountEditor(container, options)` for
framework-free lifecycle management. ProseMirror, `docx-core`, `ooxml-core` and `ooxml-ui` are regular
dependencies; document and layout algorithms live in OOXML core.

## Limitations

This is an early editor. It does not establish Microsoft Word layout parity or
lossless saving for unsupported features. Legacy DOC support is limited to
main-body text with constrained paragraph edits. SmartArt is display-only and
charts render as placeholders.

## Documentation

[Framework guide](https://christophervr.github.io/docx-viewer/frameworks/vue) |
[Bindings](https://christophervr.github.io/docx-viewer/bindings)

## License

Apache-2.0.
