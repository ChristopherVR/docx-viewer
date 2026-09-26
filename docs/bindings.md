# Framework bindings

All bindings mount `<docx-editor>` through the same `mountEditor` function. A property assignment is an external document replacement, not an edit event. `document-change` carries a `DocumentModel` directly; `document-error` carries an `Error`. Both events bubble across the shadow boundary.

React:

```tsx
import { WordEditor } from '@christophervr/docx-bindings/react';
<WordEditor documentModel={model} readOnly={false} onDocumentChange={setModel} />;
```

`ref` exposes `element`, `load(bytes)` and `save()`.

Vue:

```vue
<script setup>
import { WordEditor } from '@christophervr/docx-bindings/vue';
</script>
<template>
	<WordEditor :document-model="model" @document-change="model = $event" />
</template>
```

Angular:

```ts
import { WordEditorComponent } from '@christophervr/docx-bindings/angular';
// Add WordEditorComponent to your standalone component imports.
```

```html
<word-editor [documentModel]="model" [readOnly]="false" (documentChange)="model = $event" />
```

Svelte 5:

```svelte
<script>
  import WordEditor from '@christophervr/docx-bindings/svelte';
  let model = $state(initialDocument);
</script>
<WordEditor documentModel={model} ondocumentchange={next => model = next} />
```

Custom element:

```ts
import { registerDocxEditor } from '@christophervr/docx-web-component';
registerDocxEditor();
const element = document.createElement('docx-editor');
element.documentModel = model;
document.body.append(element);
```

The vanilla mount adapter is exported from `@christophervr/docx-bindings/vanilla` and returns update, destroy, load, save, and element operations. It avoids feedback resets when a parent returns the emitted model. Supply a new object for external model changes; mutating a model in place is not a supported reactivity mechanism.

Framework packages use their own native lifecycle and ref APIs. No framework template contains a toolbar, page renderer or document command implementation. SSR imports do not register elements; mounting is a client operation. The current browser contract validates the installed versions, not every historical peer version.
