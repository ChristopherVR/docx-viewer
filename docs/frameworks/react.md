# React

::: warning Not published to npm yet
The Word packages are unpublished, so there is nothing to `npm install` today. Build from the repository (`bun install`, `bun run demo`). The imports below show the intended API of the `@christophervr/docx-viewer` package, with the matching framework as an application-level peer.
:::

The React adapter mounts the shared `<docx-editor>` element and forwards model updates through `onDocumentChange`.

```tsx
import { WordEditor } from '@christophervr/docx-viewer/react';

export function Editor({ model, setModel }) {
	return <WordEditor documentModel={model} onDocumentChange={setModel} />;
}
```

Use the component ref for `load(bytes)` and `save()`. See the [complete binding contract](/bindings) or [try the React demo](/demo/).
