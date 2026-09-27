// @vitest-environment jsdom
import { createRoot, createSignal } from 'solid-js';
import { render } from 'solid-js/web';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createDocument } from '@christophervr/docx-core';
import { WordEditor } from './solid';

afterEach(() => document.body.replaceChildren());

describe('Solid editor adapter', () => {
	it('mounts once, reacts to option changes, forwards events, and tears down', () => {
		let dispose!: () => void;
		const [readOnly, setReadOnly] = createSignal(false);
		const [className, setClassName] = createSignal('initial');
		const onDocumentChange = vi.fn();
		let element: HTMLElement | undefined;
		const host = document.createElement('div');
		document.body.append(host);
		const props = {
			get readOnly() {
				return readOnly();
			},
			get class() {
				return className();
			},
			locale: 'fr-FR',
			onDocumentChange,
			editorRef: (handle: { element: HTMLElement }) => (element = handle.element),
		};
		createRoot((stop) => {
			dispose = stop;
			render(() => WordEditor(props), host);
		});
		const editor = host.querySelector('docx-editor') as HTMLElement & {
			readOnly: boolean;
			locale: string;
		};
		expect(editor).toBeDefined();
		expect(editor.readOnly).toBe(false);
		expect(editor.locale).toBe('fr');
		expect(element).toBe(editor);
		setReadOnly(true);
		setClassName('updated');
		expect(editor.readOnly).toBe(true);
		expect(host.firstElementChild?.className).toBe('updated');
		const model = createDocument();
		editor.dispatchEvent(new CustomEvent('document-change', { detail: model }));
		expect(onDocumentChange).toHaveBeenCalledWith(model);
		dispose();
		expect(host.querySelector('docx-editor')).toBeNull();
	});
});
