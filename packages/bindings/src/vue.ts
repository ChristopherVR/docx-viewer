import { defineComponent, h, onBeforeUnmount, onMounted, ref, watch, type PropType } from 'vue';
import type { DocumentModel } from '@christophervr/docx-core';
import { mountEditor, type EditorBinding } from './index';
export const WordEditor = defineComponent({
	name: 'WordEditor',
	props: {
		documentModel: Object as PropType<DocumentModel>,
		readOnly: Boolean,
		locale: String,
		theme: String as PropType<'light' | 'dark' | 'auto'>,
	},
	emits: {
		'document-change': (_model: DocumentModel) => true,
		'document-error': (_error: Error) => true,
	},
	setup(props, { emit, expose }) {
		const host = ref<HTMLElement>();
		let binding: EditorBinding | undefined;
		const options = () => ({
			documentModel: props.documentModel,
			readOnly: props.readOnly,
			locale: props.locale,
			theme: props.theme,
			onDocumentChange: (model: DocumentModel) => emit('document-change', model),
			onDocumentError: (error: Error) => emit('document-error', error),
		});
		onMounted(() => {
			binding = mountEditor(host.value!, options());
		});
		watch(
			() => [props.documentModel, props.readOnly, props.locale, props.theme],
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
