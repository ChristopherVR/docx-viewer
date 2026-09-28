// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createDocument } from '@christophervr/docx-core';
import {
	EDITOR_EVENT_NAMES,
	EDITOR_PROP_KEYS,
	eventOptions,
	mountEditor,
	pickEditorProps,
} from './index';
import { DOCX_EDITOR_EVENTS } from '@christophervr/docx-web-component';
afterEach(() => document.body.replaceChildren());
describe('shared binding lifecycle', () => {
	it('forwards theme to the element and defaults back to auto', () => {
		const host = document.createElement('div');
		document.body.append(host);
		const binding = mountEditor(host, { theme: 'dark' });
		expect(binding.element.theme).toBe('dark');
		expect(binding.element.getAttribute('theme')).toBe('dark');
		binding.update({});
		expect(binding.element.theme).toBe('auto');
		binding.destroy();
	});
	it('does not overwrite edits on unrelated parent updates or model feedback', () => {
		const host = document.createElement('div');
		document.body.append(host);
		const source = createDocument();
		const change = vi.fn();
		const binding = mountEditor(host, { documentModel: source, onDocumentChange: change });
		const edited = createDocument();
		edited.blocks = [{ type: 'paragraph', id: 'p1', runs: [{ text: 'In progress' }] }];
		binding.element.documentModel = edited;
		binding.element.dispatchEvent(new CustomEvent('document-change', { detail: edited }));
		expect(change).toHaveBeenCalledWith(edited);
		binding.update({ documentModel: source, readOnly: true });
		expect(binding.element.documentModel).toBe(edited);
		expect(binding.element.readOnly).toBe(true);
		binding.update({ documentModel: source, locale: 'fr-FR' });
		expect(binding.element.locale).toBe('fr');
		binding.update({ documentModel: edited });
		expect(binding.element.documentModel).toBe(edited);
		const replacement = createDocument();
		binding.update({ documentModel: replacement });
		expect(binding.element.documentModel).toBe(replacement);
		binding.destroy();
	});
	it('forwards latest callbacks and removes listeners and owned element on destroy', () => {
		const host = document.createElement('div');
		document.body.append(host);
		const sentinel = document.createElement('span');
		host.append(sentinel);
		const first = vi.fn();
		const second = vi.fn();
		const error = vi.fn();
		const binding = mountEditor(host, { onDocumentChange: first });
		binding.update({ onDocumentChange: second, onDocumentError: error });
		const model = createDocument();
		const failure = new Error('Expected test failure');
		binding.element.dispatchEvent(new CustomEvent('document-change', { detail: model }));
		binding.element.dispatchEvent(new CustomEvent('document-error', { detail: failure }));
		expect(first).not.toHaveBeenCalled();
		expect(second).toHaveBeenCalledWith(model);
		expect(error).toHaveBeenCalledWith(failure);
		binding.destroy();
		binding.destroy();
		binding.element.dispatchEvent(new CustomEvent('document-change', { detail: model }));
		expect(second).toHaveBeenCalledTimes(1);
		expect(host.children).toHaveLength(1);
		expect(host.firstChild).toBe(sentinel);
	});
});

describe('shared option keys', () => {
	it('only lists real element events and props', () => {
		for (const name of EDITOR_EVENT_NAMES) expect(DOCX_EDITOR_EVENTS).toContain(name);
		const element = document.createElement('docx-editor');
		for (const key of EDITOR_PROP_KEYS) expect(key in element).toBe(true);
	});
	it('picks exactly the shared props and maps event handlers', () => {
		const model = createDocument();
		const extra = { documentModel: model, readOnly: true, locale: 'fr', other: 1 };
		expect(pickEditorProps(extra)).toEqual({ documentModel: model, readOnly: true, locale: 'fr' });
		const change = vi.fn();
		const options = eventOptions({ 'document-change': change, 'document-error': undefined });
		options.onDocumentChange?.(model);
		expect(change).toHaveBeenCalledWith(model);
		expect(options.onDocumentError).toBeUndefined();
	});
});
