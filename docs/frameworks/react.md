# React

The React adapter mounts the shared `<docx-editor>` element and forwards model updates through `onDocumentChange`.

```tsx
import { WordEditor } from 'docx-react-viewer';

export function Editor({ model, setModel }) {
	return <WordEditor documentModel={model} onDocumentChange={setModel} />;
}
```

Use the component ref for `load(bytes)` and `save()`. See the [complete binding contract](/bindings) or [try the React demo](/demo/).
