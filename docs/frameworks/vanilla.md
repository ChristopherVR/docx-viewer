# Vanilla JavaScript

Install the entry-point package once, then add the matching framework if your application does not already include it:

```sh
npm install @christophervr/docx-viewer
```

Import the vanilla mount adapter and provide a host element. The returned handle includes load, save, update, destroy, and the custom element.

```ts
import { mountEditor } from '@christophervr/docx-viewer/vanilla';

const editor = await mountEditor(document.querySelector('#editor'), {
	documentModel: model,
	onDocumentChange(next) {
		model = next;
	},
});
```

The custom element can also be registered directly from `@christophervr/docx-viewer/web-component`. See the [complete binding contract](/bindings) or [try the vanilla demo](/demo-vanilla/).
