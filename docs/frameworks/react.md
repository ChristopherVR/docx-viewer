# React

::: warning Not published to npm yet
Nothing has been released yet, so there is nothing to `npm install` today. Build from the repository (`bun install`, `bun run demo`). The imports below are the intended API of the self-contained `docx-react-viewer` package; it needs only `docx-core` and react next to it.
:::

The React adapter mounts the shared `<docx-editor>` element and forwards model updates through `onDocumentChange`.

```tsx
import { WordEditor } from 'docx-react-viewer';

export function Editor({ model, setModel }) {
	return <WordEditor documentModel={model} onDocumentChange={setModel} />;
}
```

Use the component ref for `load(bytes)` and `save()`. See the [complete binding contract](/bindings) or [try the React demo](/demo/).
