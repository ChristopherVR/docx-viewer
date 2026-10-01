# Vanilla JavaScript

::: warning Not published to npm yet
The Word packages are unpublished, so there is nothing to `npm install` today. Build from the repository (`bun install`, `bun run demo`). The imports below show the intended API of the `@christophervr/docx-viewer` package, with the matching framework as an application-level peer.
:::

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
