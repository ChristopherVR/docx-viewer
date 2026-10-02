// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { createDocument, saveDocx } from 'docx-core';
import type { EditorView } from 'prosemirror-view';
import { undo } from 'prosemirror-history';
import { CollaborationAuthority, DocxEditorElement } from './index';
import { modelToDoc } from './model-adapter';

afterEach(() => document.body.replaceChildren());
function mount() {
	const element = document.createElement('docx-editor') as DocxEditorElement;
	document.body.append(element);
	return element;
}
const viewOf = (element: DocxEditorElement) => (element as unknown as { view: EditorView }).view;

describe('collaboration component lifecycle', () => {
	it('undoes a local edit without undoing another author’s text', () => {
		const a = mount();
		const b = mount();
		const authority = new CollaborationAuthority({
			sessionId: 'undo',
			doc: modelToDoc(a.documentModel!),
		});
		a.startCollaboration({ sessionId: 'undo', clientId: 'a' });
		b.startCollaboration({ sessionId: 'undo', clientId: 'b' });
		const relay = (sender: DocxEditorElement) => {
			const result = authority.submit(sender.getPendingCollaboration());
			if (result.status === 'rejected') throw new Error(result.reason);
			expect(a.receiveCollaboration(result.batch)).toBe('applied');
			expect(b.receiveCollaboration(result.batch)).toBe('applied');
		};
		viewOf(a).dispatch(viewOf(a).state.tr.insertText('A', 1));
		relay(a);
		viewOf(b).dispatch(viewOf(b).state.tr.insertText('B', 2));
		relay(b);
		expect(undo(viewOf(a).state, viewOf(a).dispatch)).toBe(true);
		relay(a);
		expect(viewOf(a).state.doc.textContent).toBe('B');
		expect(viewOf(b).state.doc.textContent).toBe('B');
	});
	it('retains pending state across detach and rejects document replacement until the session is stopped', () => {
		const a = mount();
		const authority = new CollaborationAuthority({
			sessionId: 'life',
			doc: modelToDoc(a.documentModel!),
		});
		a.startCollaboration({ sessionId: 'life', clientId: 'alice' });
		viewOf(a).dispatch(viewOf(a).state.tr.insertText('pending', 1));
		const pending = a.getPendingCollaboration()!;
		a.remove();
		document.body.append(a);
		expect(a.getPendingCollaboration()).toEqual(pending);
		expect(viewOf(a).state.doc.textContent).toBe('pending');
		expect(() => {
			a.documentModel = createDocument();
		}).toThrow(/Stop collaboration/);
		expect(() => a.stopCollaboration()).toThrow(/pending/);
		const accepted = authority.submit(pending);
		if (accepted.status === 'rejected') throw new Error(accepted.reason);
		expect(a.receiveCollaboration(accepted.batch)).toBe('applied');
		expect(a.receiveCollaboration(accepted.batch)).toBe('duplicate');
		expect(a.getPendingCollaboration()).toBeNull();
		a.stopCollaboration();
		a.documentModel = createDocument();
		expect(viewOf(a).state.doc.textContent).toBe('');
	});

	it('applies remote edits in read-only mode while preserving the loaded DOCX package context', async () => {
		const zip = await JSZip.loadAsync(await saveDocx(createDocument()));
		zip.file('custom/preserved.txt', 'untouched package payload');
		const source = await zip.generateAsync({ type: 'uint8array' });
		const a = mount();
		const b = mount();
		await Promise.all([a.load(source), b.load(source)]);
		const authority = new CollaborationAuthority({
			sessionId: 'package',
			doc: modelToDoc(a.documentModel!),
		});
		a.startCollaboration({ sessionId: 'package', clientId: 'a' });
		b.startCollaboration({ sessionId: 'package', clientId: 'b' });
		b.readOnly = true;
		viewOf(a).dispatch(viewOf(a).state.tr.insertText('مرحبا coauthors', 1));
		const accepted = authority.submit(a.getPendingCollaboration());
		if (accepted.status === 'rejected') throw new Error(accepted.reason);
		expect(a.receiveCollaboration(accepted.batch)).toBe('applied');
		expect(b.receiveCollaboration(accepted.batch)).toBe('applied');
		expect(viewOf(b).state.doc.textContent).toBe('مرحبا coauthors');
		expect(b.readOnly).toBe(true);
		const saved = await JSZip.loadAsync(await b.saveBytes());
		expect(await saved.file('custom/preserved.txt')!.async('string')).toBe(
			'untouched package payload',
		);
		expect(await saved.file('word/document.xml')!.async('string')).toContain('مرحبا coauthors');
	});
});
