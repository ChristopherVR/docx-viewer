# Svelte

::: warning Not published to npm yet
Nothing has been released yet, so there is nothing to `npm install` today. Build from the repository (`bun install`, `bun run demo`). The imports below are the intended API of the self-contained `@christophervr/docx-svelte-viewer` package; it needs only `@christophervr/docx-core` and svelte next to it.
:::

Use the Svelte adapter as a component and handle model changes with the event callback.

```svelte
<script>
	import WordEditor from '@christophervr/docx-svelte-viewer';
	let model = $state(initialDocument);
</script>

<WordEditor documentModel={model} ondocumentchange={(next) => (model = next)} />
```

The package root is the component (default export). The plain-JavaScript helpers, `mountEditor`, `loadDocument` and the shared types, are under `@christophervr/docx-svelte-viewer/runtime`.

See the [complete binding contract](/bindings) or [try the Svelte demo](/demo-svelte/).
