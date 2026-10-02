# docx-core

The framework-neutral DOCX document model, parser and preserving serializer. A thin entry point: `docx-core` re-exports `ooxml-core/docx` and `docx-core/embedded` re-exports `ooxml-core/docx/embedded`; the logic lives in [ooxml-core](https://github.com/ChristopherVR/ooxml-core).

```sh
npm install docx-core
```

```ts
import { createDocument, loadDocx, saveDocx } from 'docx-core';

const model = createDocument();
const bytes = await saveDocx(model);
const loaded = await loadDocx(bytes);
```

It handles modern `.docx` only. To open legacy `.doc` files use one of the editor packages (`docx-react-viewer`, `-vue-viewer`, `-angular-viewer`, `-svelte-viewer`, `-solid-viewer`, `-vanilla-viewer`), whose `loadDocument` detects both formats. Saving is not lossless for unsupported features.
