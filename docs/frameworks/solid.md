# SolidJS

The SolidJS adapter mounts the shared `<docx-editor>` element and forwards model changes through `onDocumentChange`.

```tsx
import { WordEditor } from 'docx-solid-viewer';

function Editor(props) {
	return (
		<WordEditor
			documentModel={props.model}
			onDocumentChange={props.setModel}
			editorRef={(handle) => props.setHandle(handle)}
		/>
	);
}
```

The adapter calls `editorRef` with a handle exposing `load(bytes)` and `save()`. See the [complete binding contract](/bindings) or [try the Solid demo](/demo-solid/).
