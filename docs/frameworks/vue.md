# Vue

Install the entry-point package once, then add the matching framework if your application does not already include it:

```sh
npm install @christophervr/docx-viewer
```

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
