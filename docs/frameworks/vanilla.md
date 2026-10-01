# Vanilla JavaScript

::: warning Not published to npm yet
Nothing has been released yet, so there is nothing to `npm install` today. Build from the repository (`bun install`, `bun run demo`). The imports below are the intended API of the self-contained `@christophervr/docx-vanilla-viewer` package; it needs only `@christophervr/docx-core` next to it.
:::

Import the vanilla mount adapter and provide a host element. The returned handle includes load, save, update, destroy, and the custom element.

```ts
import { mountEditor } from '@christophervr/docx-vanilla-viewer';

const editor = await mountEditor(document.querySelector('#editor'), {
	documentModel: model,
	onDocumentChange(next) {
		model = next;
	},
});
```

The custom element can also be registered directly from `@christophervr/docx-viewer/web-component`. See the [complete binding contract](/bindings) or [try the vanilla demo](/demo-vanilla/).
