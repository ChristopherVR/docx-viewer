# Svelte

::: warning Not published to npm yet
The Word packages are unpublished, so there is nothing to `npm install` today. Build from the repository (`bun install`, `bun run demo`). The imports below show the intended API of the `@christophervr/docx-viewer` package, with the matching framework as an application-level peer.
:::

Use the Svelte adapter as a component and handle model changes with the event callback.

```svelte
<script>
	import WordEditor from '@christophervr/docx-viewer/svelte';
	let model = $state(initialDocument);
</script>

<WordEditor documentModel={model} ondocumentchange={(next) => (model = next)} />
```

See the [complete binding contract](/bindings) or [try the Svelte demo](/demo-svelte/).
