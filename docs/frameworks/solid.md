# SolidJS

Install the entry-point package once, then add `solid-js` if your application does not already include it:

```sh
npm install @christophervr/docx-viewer
```

The SolidJS adapter mounts the shared `<docx-editor>` element and forwards model changes through `onDocumentChange`.

```tsx
import { WordEditor } from '@christophervr/docx-viewer/solid';

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
