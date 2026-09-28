// @vitest-environment jsdom
import { afterEach, describe, expect, expectTypeOf, it } from 'vitest';
import { createDocument, type DocumentModel } from '@christophervr/docx-core';
import { DocxEditorElement, registerDocxEditor } from './index';
import { DOCX_EDITOR_EVENTS, emit, on, type DocxEditorEventMap } from './events';
import type { FileCommandDetail } from './file-commands';
import type { RibbonAction } from './ribbon';
import { at } from './test-support';

registerDocxEditor();
afterEach(() => document.body.replaceChildren());

describe('typed event contract', () => {
	it('emit dispatches a bubbling, composed CustomEvent carrying the detail', () => {
		const target = document.createElement('div');
		document.body.append(target);
		const seen: CustomEvent[] = [];
		document.addEventListener('document-warning', (event) => seen.push(event as CustomEvent));
		expect(emit(target, 'document-warning', 'careful')).toBe(true);
		expect(seen).toHaveLength(1);
		const event = at(seen, 0);
		expect(event).toBeInstanceOf(CustomEvent);
		expect(event.detail).toBe('careful');
		expect(event.bubbles && event.composed && !event.cancelable).toBe(true);
	});

	it('emit reports cancellation only for cancelable events', () => {
		const target = document.createElement('div');
		on(target, 'file-command', (event) => event.preventDefault());
		expect(emit(target, 'file-command', { command: 'open' }, { cancelable: true })).toBe(false);
		expect(emit(target, 'file-command', { command: 'open' })).toBe(true);
	});

	it('delivers typed details through the element and does not fire change on model assignment', () => {
		const editor = document.createElement('docx-editor');
		document.body.append(editor);
		const models: DocumentModel[] = [];
		const commands: FileCommandDetail[] = [];
		editor.addEventListener('document-change', (event) => models.push(event.detail));
		editor.addEventListener('file-command', (event) => commands.push(event.detail));
		editor.documentModel = createDocument();
		expect(models).toHaveLength(0);
		emit(editor, 'file-command', { command: 'save' }, { cancelable: true });
		expect(commands).toEqual([{ command: 'save' }]);
	});

	it('lists every mapped event exactly once', () => {
		expect(new Set(DOCX_EDITOR_EVENTS).size).toBe(DOCX_EDITOR_EVENTS.length);
		const complete: Record<keyof DocxEditorEventMap, true> = {
			'document-change': true,
			'document-error': true,
			'document-warning': true,
			'readonly-change': true,
			'ribbon-action': true,
			'file-command': true,
			'page-change': true,
			'dirty-change': true,
			'presence-send': true,
			'collaboration-send': true,
		};
		expect([...DOCX_EDITOR_EVENTS].sort()).toEqual(Object.keys(complete).sort());
	});

	it('types listener details from the event name', () => {
		const editor = document.createElement('docx-editor');
		expectTypeOf(editor).toEqualTypeOf<DocxEditorElement>();
		editor.addEventListener('document-change', (event) => {
			expectTypeOf(event.detail).toEqualTypeOf<DocumentModel>();
		});
		editor.addEventListener('document-error', (event) => {
			expectTypeOf(event.detail).toEqualTypeOf<Error>();
		});
		editor.addEventListener('readonly-change', (event) => {
			expectTypeOf(event.detail).toEqualTypeOf<boolean>();
		});
		editor.addEventListener('ribbon-action', (event) => {
			expectTypeOf(event.detail).toEqualTypeOf<RibbonAction>();
		});
		editor.addEventListener('file-command', (event) => {
			expectTypeOf(event.detail).toEqualTypeOf<FileCommandDetail>();
		});
		// Untyped names still fall back to the DOM signature.
		editor.addEventListener('click', (event) => expectTypeOf(event.clientX).toBeNumber());
		editor.addEventListener('custom-thing', (event) => expectTypeOf(event).toEqualTypeOf<Event>());
		expectTypeOf<DocxEditorEventMap['document-change']['detail']>().toEqualTypeOf<DocumentModel>();
	});
});
