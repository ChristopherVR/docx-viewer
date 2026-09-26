# Svelte

Use the Svelte adapter as a component and handle model changes with the event callback.

```svelte
<script>
	import WordEditor from '@christophervr/docx-bindings/svelte';
	let model = $state(initialDocument);
</script>

<WordEditor documentModel={model} ondocumentchange={(next) => (model = next)} />
```

See the [complete binding contract](/bindings) or [try the Svelte demo](/demo-svelte/).
