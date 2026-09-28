import { defineComponent, h, onBeforeUnmount, onMounted, ref, watch, type PropType } from 'vue';
import type { DocumentModel } from '@christophervr/docx-core';
import type { EditorThemeMode } from '@christophervr/docx-web-component';
import {
	EDITOR_EVENT_NAMES,
	EDITOR_PROP_KEYS,
	eventOptions,
	mountEditor,
	pickEditorProps,
	type EditorBinding,
	type EditorPropKey,
} from './index';
export const WordEditor = defineComponent({
	name: 'WordEditor',
	props: {
		documentModel: Object as PropType<DocumentModel>,
		readOnly: Boolean,
		locale: String,
		theme: String as PropType<EditorThemeMode>,
	} satisfies Record<EditorPropKey, unknown>,
	emits: [...EDITOR_EVENT_NAMES],
	setup(props, { emit, expose }) {
		const host = ref<HTMLElement>();
		let binding: EditorBinding | undefined;
		const options = () => ({
			...pickEditorProps(props),
			...eventOptions({
				'document-change': (model) => emit('document-change', model),
				'document-error': (error) => emit('document-error', error),
			}),
		});
		onMounted(() => {
			binding = mountEditor(host.value!, options());
		});
		watch(
			() => EDITOR_PROP_KEYS.map((key) => props[key]),
			() => binding?.update(options()),
		);
		onBeforeUnmount(() => binding?.destroy());
		expose({
			get element() {
				return binding?.element;
			},
			async load(input: Uint8Array | ArrayBuffer) {
				if (!binding) throw new Error('Editor is not mounted');
				await binding.load(input);
			},
			async save() {
				if (!binding) throw new Error('Editor is not mounted');
				return binding.save();
			},
		});
		return () => h('div', { ref: host });
	},
});
