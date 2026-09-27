# @christophervr/docx-viewer

Install the Word viewer family with one package. Framework entry points stay isolated: import only the subpath you use so React, Vue, Angular, Svelte or Solid remains an application-level peer dependency.

```sh
npm install @christophervr/docx-viewer
```

```ts
import { createDocument, mountEditor } from '@christophervr/docx-viewer';
const editor = mountEditor(document.querySelector('#editor')!, {
	documentModel: createDocument(),
});
```

Subpaths are `./core`, `./document`, `./legacy`, `./web-component`, `./vanilla`, `./react`, `./vue`, `./angular`, `./svelte` and `./solid`. Framework-specific peers must also be installed when using those entries. The package root and core/document entries can be imported in SSR environments; call `mountEditor` or `registerDocxEditor` only in a browser.

The framework bindings share one editor and document model. See the [React](https://github.com/ChristopherVR/docx-viewer/blob/main/docs/frameworks/react.md), [Vue](https://github.com/ChristopherVR/docx-viewer/blob/main/docs/frameworks/vue.md), [Angular](https://github.com/ChristopherVR/docx-viewer/blob/main/docs/frameworks/angular.md), [Svelte](https://github.com/ChristopherVR/docx-viewer/blob/main/docs/frameworks/svelte.md), [Solid](https://github.com/ChristopherVR/docx-viewer/blob/main/docs/frameworks/solid.md), and [vanilla JavaScript](https://github.com/ChristopherVR/docx-viewer/blob/main/docs/frameworks/vanilla.md) guides.
