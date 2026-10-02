# Vue

::: warning Not published to npm yet
Nothing has been released yet, so there is nothing to `npm install` today. Build from the repository (`bun install`, `bun run demo`). The imports below are the intended API of the self-contained `docx-vue-viewer` package; it needs only `docx-core` and vue next to it.
:::

The Vue adapter exposes the editor as a component with model props and Vue events.

```vue
<script setup>
import { WordEditor } from 'docx-vue-viewer';
const model = defineModel();
</script>

<template>
	<WordEditor :document-model="model" @document-change="model = $event" />
</template>
```

See the [complete binding contract](/bindings) or [try the Vue demo](/demo-vue/).
