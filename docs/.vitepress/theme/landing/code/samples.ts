/**
 * Framework metadata + code samples for the landing page. None of this is
 * localized: package names, entry points and code are identical in every
 * locale. Samples mirror docs/bindings.md and the per-framework guides; keep
 * them in sync with those pages. The Word packages are not published to npm
 * yet, so `entry` is the intended import path, not an install command.
 */

export interface FrameworkSample {
	id: string;
	label: string;
	/** Intended import path of the adapter (API preview, not yet on npm). */
	entry: string;
	file: string;
	docsHref: string;
	code: string;
}

export const FRAMEWORKS: FrameworkSample[] = [
	{
		id: 'react',
		label: 'React',
		entry: 'docx-react-viewer',
		file: 'Editor.tsx',
		docsHref: '/frameworks/react',
		code: `import { useState } from 'react';
import { createDocument } from 'docx-core';
import { WordEditor } from 'docx-react-viewer';

export function Editor() {
  const [model, setModel] = useState(() => createDocument());
  return (
    <WordEditor
      documentModel={model}
      readOnly={false}
      onDocumentChange={setModel}
    />
  );
}`,
	},
	{
		id: 'vue',
		label: 'Vue 3',
		entry: 'docx-vue-viewer',
		file: 'Editor.vue',
		docsHref: '/frameworks/vue',
		code: `<script setup lang="ts">
import { ref } from 'vue';
import { createDocument } from 'docx-core';
import { WordEditor } from 'docx-vue-viewer';

const model = ref(createDocument());
</script>

<template>
  <WordEditor :document-model="model" @document-change="model = $event" />
</template>`,
	},
	{
		id: 'angular',
		label: 'Angular',
		entry: 'docx-angular-viewer',
		file: 'editor.component.ts',
		docsHref: '/frameworks/angular',
		code: `import { Component } from '@angular/core';
import { createDocument } from 'docx-core';
import { WordEditorComponent } from 'docx-angular-viewer';

@Component({
  selector: 'app-editor',
  standalone: true,
  imports: [WordEditorComponent],
  template: '<word-editor [documentModel]="model" (documentChange)="model = $event" />',
})
export class EditorComponent {
  model = createDocument();
}`,
	},
	{
		id: 'svelte',
		label: 'Svelte 5',
		entry: 'docx-svelte-viewer',
		file: 'Editor.svelte',
		docsHref: '/frameworks/svelte',
		code: `<script lang="ts">
  import { createDocument } from 'docx-core';
  import WordEditor from 'docx-svelte-viewer';

  let model = $state(createDocument());
</script>

<WordEditor documentModel={model} ondocumentchange={(next) => (model = next)} />`,
	},
	{
		id: 'solid',
		label: 'SolidJS',
		entry: 'docx-solid-viewer',
		file: 'Editor.tsx',
		docsHref: '/frameworks/solid',
		code: `import { createSignal } from 'solid-js';
import { createDocument } from 'docx-core';
import { WordEditor } from 'docx-solid-viewer';

export function Editor() {
  const [model, setModel] = createSignal(createDocument());
  return <WordEditor documentModel={model()} onDocumentChange={setModel} />;
}`,
	},
	{
		id: 'vanilla',
		label: 'Vanilla JS',
		entry: 'docx-vanilla-viewer',
		file: 'main.ts',
		docsHref: '/frameworks/vanilla',
		code: `import { createDocument } from 'docx-core';
import { mountEditor } from 'docx-vanilla-viewer';

const editor = mountEditor(document.querySelector('#editor')!, {
  documentModel: createDocument(),
  onDocumentChange: (model) => console.log(model),
  onDocumentError: (error) => console.error(error),
});

await editor.load(bytes); // DOCX or legacy DOC bytes
const saved = await editor.save(); // original format, edits applied
editor.destroy();`,
	},
];

/** Load and save outside any UI: the framework-neutral document API. */
export const CORE_SAMPLE = `import { loadDocument } from 'docx-vanilla-viewer';

// Any editor package bundles this loader; it detects DOCX or legacy DOC from the bytes.
const loaded = await loadDocument(bytes);
console.log(loaded.model.blocks.length);

// Returns bytes in the original format. An untouched
// document comes back byte-for-byte; unsafe edits reject.
const output = await loaded.save(loaded.model);`;
