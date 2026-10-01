# SolidJS

::: warning Not published to npm yet
The Word packages are unpublished, so there is nothing to `npm install` today. Build from the repository (`bun install`, `bun run demo`). The imports below show the intended API of the `@christophervr/docx-viewer` package, with the matching framework as an application-level peer.
:::

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
