# Vue

::: warning Not published to npm yet
The Word packages are unpublished, so there is nothing to `npm install` today. Build from the repository (`bun install`, `bun run demo`). The imports below show the intended API of the `@christophervr/docx-viewer` package, with the matching framework as an application-level peer.
:::

The Vue adapter exposes the editor as a component with model props and Vue events.

```vue
<script setup>
import { WordEditor } from '@christophervr/docx-viewer/vue';
const model = defineModel();
</script>

<template>
	<WordEditor :document-model="model" @document-change="model = $event" />
</template>
```

See the [complete binding contract](/bindings) or [try the Vue demo](/demo-vue/).
