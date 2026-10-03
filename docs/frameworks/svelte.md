# Svelte

Use the Svelte adapter as a component and handle model changes with the event callback.

```svelte
<script>
	import WordEditor from 'docx-svelte-viewer';
	let model = $state(initialDocument);
</script>

<WordEditor documentModel={model} ondocumentchange={(next) => (model = next)} />
```

The package root is the component (default export). The plain-JavaScript helpers, `mountEditor`, `loadDocument` and the shared types, are under `docx-svelte-viewer/runtime`.

See the [complete binding contract](/bindings) or [try the Svelte demo](/demo-svelte/).
