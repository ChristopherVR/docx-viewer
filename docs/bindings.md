# Framework bindings

Install the entry-point package once, then add the matching framework if your application does not already include it:

```sh
npm install @christophervr/docx-viewer
```

All bindings mount `<docx-editor>` through the same `mountEditor` function. A property assignment is an external document replacement, not an edit event. `document-change` carries a `DocumentModel` directly; `document-error` carries an `Error`. Both events bubble across the shadow boundary.

React:

```tsx
import { WordEditor } from '@christophervr/docx-viewer/react';
<WordEditor documentModel={model} readOnly={false} onDocumentChange={setModel} />;
```

`ref` exposes `element`, `load(bytes)` and `save()`.

Vue:

```vue
<script setup>
import { WordEditor } from '@christophervr/docx-viewer/vue';
</script>
<template>
	<WordEditor :document-model="model" @document-change="model = $event" />
</template>
```

Angular:

```ts
import { WordEditorComponent } from '@christophervr/docx-viewer/angular';
// Add WordEditorComponent to your standalone component imports.
```

```html
<word-editor [documentModel]="model" [readOnly]="false" (documentChange)="model = $event" />
```

Svelte 5:

```svelte
<script>
  import WordEditor from '@christophervr/docx-viewer/svelte';
  let model = $state(initialDocument);
</script>
<WordEditor documentModel={model} ondocumentchange={next => model = next} />
```

SolidJS:

```tsx
import { createSignal } from 'solid-js';
import { WordEditor } from '@christophervr/docx-viewer/solid';

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
import { registerDocxEditor } from '@christophervr/docx-viewer/web-component';
registerDocxEditor();
const element = document.createElement('docx-editor');
element.documentModel = model;
document.body.append(element);
```

The vanilla mount adapter is exported from `@christophervr/docx-viewer/vanilla` and returns update, destroy, load, save, and element operations. It avoids feedback resets when a parent returns the emitted model. Supply a new object for external model changes; mutating a model in place is not a supported reactivity mechanism.

Framework packages use their own native lifecycle and ref APIs. No framework template contains a toolbar, page renderer or document command implementation. SSR imports do not register elements; mounting is a client operation. The current browser contract validates the installed versions, not every historical peer version.

Set the optional `locale` prop on any framework editor to localize the ribbon and search controls (for example, `locale="fr"`). This affects interface text only; it does not translate document content.
