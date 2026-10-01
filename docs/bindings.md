# Framework bindings

::: warning Not published to npm yet
Nothing has been released yet, so there is nothing to `npm install` today. Build from the repository (`bun install`, `bun run demo`). Each framework has its own self-contained package: `@christophervr/docx-react-viewer`, `-vue-viewer`, `-angular-viewer`, `-svelte-viewer`, `-solid-viewer` and `-vanilla-viewer`. Each bundles the editor, layout engine and legacy `.doc` reader, and needs only `@christophervr/docx-core` plus its framework as a peer. Use one editor package per application.
:::

All bindings mount `<docx-editor>` through the same `mountEditor` function. A property assignment is an external document replacement, not an edit event. `document-change` carries a `DocumentModel` directly; `document-error` carries an `Error`. Both events bubble across the shadow boundary.

React:

```tsx
import { WordEditor } from '@christophervr/docx-react-viewer';
<WordEditor documentModel={model} readOnly={false} onDocumentChange={setModel} />;
```

`ref` exposes `element`, `load(bytes)` and `save()`.

Vue:

```vue
<script setup>
import { WordEditor } from '@christophervr/docx-vue-viewer';
</script>
<template>
	<WordEditor :document-model="model" @document-change="model = $event" />
</template>
```

Angular:

```ts
import { WordEditorComponent } from '@christophervr/docx-angular-viewer';
// Add WordEditorComponent to your standalone component imports.
```

```html
<word-editor [documentModel]="model" [readOnly]="false" (documentChange)="model = $event" />
```

Svelte 5:

```svelte
<script>
  import WordEditor from '@christophervr/docx-svelte-viewer';
  let model = $state(initialDocument);
</script>
<WordEditor documentModel={model} ondocumentchange={next => model = next} />
```

SolidJS:

```tsx
import { createSignal } from 'solid-js';
import { WordEditor } from '@christophervr/docx-solid-viewer';

function DocumentEditor() {
	const [model, setModel] = createSignal(initialDocument);
	return <WordEditor documentModel={model()} onDocumentChange={setModel} />;
}
```

Solid uses the same `documentModel`, `readOnly`, `locale`, `onDocumentChange`, and
`onDocumentError` options as the shared adapter. Use `editorRef` to obtain the
mounted editor's `element`, `load`, and `save` methods. The Solid owner cleans up
the editor and its listeners when the component unmounts.

Custom element:

```ts
import { registerDocxEditor } from '@christophervr/docx-vanilla-viewer';
registerDocxEditor();
const element = document.createElement('docx-editor');
element.documentModel = model;
document.body.append(element);
```

The vanilla mount adapter is exported from `@christophervr/docx-vanilla-viewer` and returns update, destroy, load, save, and element operations. It avoids feedback resets when a parent returns the emitted model. Supply a new object for external model changes; mutating a model in place is not a supported reactivity mechanism.

Framework packages use their own native lifecycle and ref APIs. No framework template contains a toolbar, page renderer or document command implementation. SSR imports do not register elements; mounting is a client operation. The current browser contract validates the installed versions, not every historical peer version.

Set the optional `locale` prop on any framework editor to localize the ribbon and search controls (`en`, `fr`, `de`, `es` or `zh-CN`; for example, `locale="de"`). This affects interface text only; it does not translate document content.

## Window chrome and file commands

The editor draws Word-style chrome inside its shadow root: a title bar (quick-access Save,
Undo and Redo, the file name and save state, a "Tell me what you want to do" command search,
comments and an Editing/Viewing mode switch), a **File** tab with a backstage (Info with
compatibility notes, New, Open, Save, Save a copy as DOCX, Print), and a status bar (page,
words, compatibility notes, Web/Print Layout and zoom).

Set `element.fileName` to name the open document. File commands are announced first as a
cancelable `file-command` event whose `detail.command` is `new`, `open`, `save`, `export` or
`print`. Call `preventDefault()` to handle it in your application (for example to save to your
own storage); otherwise the editor uses its browser-only default: a file picker, a download, or
the print dialog. Choosing **Viewing** raises `readonly-change`.

```ts
editor.fileName = 'Quarterly report.docx';
editor.addEventListener('file-command', (event) => {
	const { command } = (event as CustomEvent<{ command: string }>).detail;
	if (command === 'save') {
		event.preventDefault();
		void editor.save().then(uploadToMyStorage);
	}
});
```

## Saving, dirty state and UI options

`save()` resolves to a `Blob` of the serialized document (`saveBytes()` returns the raw
`Uint8Array`), `download(fileName?)` saves and starts a browser download, and `dirty` /
`dirty-change` report unsaved edits. `dirty` becomes true on `document-change` and false after
File > Save, `download()`, loading a document or `markClean()`; `save()` alone does not clear it
because only the host knows whether the Blob was persisted. Every adapter exposes
`save`, `download`, `markClean` and `dirty` on its handle and `onDirtyChange` / `onPageChange`
(or the framework's event equivalent).

| Property (attribute)                 | Default | Effect                                                                                                                                                                                                                                                                                                |
| ------------------------------------ | ------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `showThumbnails` (`show-thumbnails`) | `false` | Left rail of page thumbnails. Thumbnails need Print Layout; in Draft view the rail says so instead of listing pages.                                                                                                                                                                                  |
| `showToolbar` (`show-toolbar`)       | `true`  | `show-toolbar="false"` hides the ribbon; the title bar and status bar remain.                                                                                                                                                                                                                         |
| `hiddenActions`                      | `[]`    | Ribbon controls to hide by stable kebab-case id (`RIBBON_ACTION_IDS`, e.g. `'bold'`, `'insert-table'`, `'track-changes'`, `'paste'`, `'word-count'`); empty groups and tabs hide too. A colour button and its dropdown caret hide together. Unknown ids are ignored. Ids do not change with `locale`. |

**Deprecated:** before ids existed `hiddenActions` took the English control label (`'Bold'`,
`'Insert table'`). Those `LegacyRibbonLabel` strings are still accepted for one more release: they
are mapped to ids and the element dispatches one `document-warning` per assignment. Reading
`hiddenActions` always returns ids. `ribbon-action` events never carried labels (their detail is
keyed by command type), so they are unchanged.

`page-change` (`{ page, pageCount }`) fires from Print Layout only. Pagination comes from this
editor's own layout engine and is an approximation, not Word's pagination; the status bar and
rail tooltips say so.
