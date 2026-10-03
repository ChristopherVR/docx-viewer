# Vanilla JavaScript

Import the vanilla mount adapter and provide a host element. The returned handle includes load, save, update, destroy, and the custom element.

```ts
import { mountEditor } from 'docx-vanilla-viewer';

const editor = await mountEditor(document.querySelector('#editor'), {
	documentModel: model,
	onDocumentChange(next) {
		model = next;
	},
});
```

The custom element can also be registered directly from `docx-viewer/web-component`. See the [complete binding contract](/bindings) or [try the vanilla demo](/demo-vanilla/).
