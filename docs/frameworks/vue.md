# Vue

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
