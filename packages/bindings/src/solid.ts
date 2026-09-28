import { createEffect, createSignal, onCleanup, onMount } from 'solid-js';
import type { EditorBinding, EditorHandle, EditorOptions } from './index';
import { eventOptions, mountEditor, pickEditorProps } from './index';

export interface WordEditorProps extends EditorOptions {
	class?: string;
	editorRef?: (handle: EditorHandle) => void;
}

/** Solid lifecycle adapter over the shared editor binding. */
export function WordEditor(props: WordEditorProps) {
	const host = document.createElement('div');
	const [binding, setBinding] = createSignal<EditorBinding>();
	const options = (): EditorOptions => ({
		...pickEditorProps(props),
		...eventOptions({
			'document-change': props.onDocumentChange,
			'document-error': props.onDocumentError,
			'page-change': props.onPageChange,
			'dirty-change': props.onDirtyChange,
		}),
	});
	onMount(() => {
		const mounted = mountEditor(host, options());
		setBinding(mounted);
		props.editorRef?.({
			get element() {
				return mounted.element;
			},
			load: (input) => mounted.load(input),
			save: () => mounted.save(),
			download: (fileName) => mounted.download(fileName),
			markClean: () => mounted.markClean(),
			get dirty() {
				return mounted.dirty;
			},
		});
	});
	createEffect(() => {
		const current = binding();
		if (!current) return;
		host.className = props.class ?? '';
		current.update(options());
	});
	onCleanup(() => binding()?.destroy());
	return host;
}
