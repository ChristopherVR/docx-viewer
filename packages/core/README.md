# docx-core

[![npm version](https://img.shields.io/npm/v/docx-core.svg)](https://www.npmjs.com/package/docx-core)
[![license](https://img.shields.io/npm/l/docx-core.svg)](https://github.com/ChristopherVR/docx-viewer/blob/main/LICENSE)
[![types](https://img.shields.io/npm/types/docx-core.svg)](https://www.npmjs.com/package/docx-core)

> The DOM-free document model and file API, re-exported from the canonical OOXML core.

[Live demo](https://christophervr.github.io/docx-viewer/demo/) | [npm](https://www.npmjs.com/package/docx-core) | [Full docs](https://christophervr.github.io/docx-viewer/) | [Source](https://github.com/ChristopherVR/docx-viewer)

The framework-neutral DOCX document model, parser and preserving serializer. A thin entry point: `docx-core` re-exports `ooxml-core/docx` and `docx-core/embedded` re-exports `ooxml-core/docx/embedded`; the logic lives in [ooxml-core](https://github.com/ChristopherVR/ooxml).

## Install

```bash
npm install docx-core
```

## Quick start

```ts
import { createDocument, loadDocx, saveDocx } from 'docx-core';

const model = createDocument();
const bytes = await saveDocx(model);
const loaded = await loadDocx(bytes);
```

## API and limitations

It handles modern `.docx` only. To open legacy `.doc` files use one of the editor packages (`docx-react-viewer`, `-vue-viewer`, `-angular-viewer`, `-svelte-viewer`, `-solid-viewer`, `-vanilla-viewer`), whose `loadDocument` detects both formats. Saving is not lossless for unsupported features.

## Documentation

[Full docs](https://christophervr.github.io/docx-viewer/) | [Core source](https://github.com/ChristopherVR/ooxml)

## License

Apache-2.0.
