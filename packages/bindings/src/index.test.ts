// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createDocument } from 'docx-core';
import {
	EDITOR_EVENT_NAMES,
	EDITOR_PROP_KEYS,
	eventOptions,
	mountEditor,
	pickEditorProps,
} from './index';
import { DOCX_EDITOR_EVENTS } from 'docx-web-component';
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
		const options = eventOptions({
			'document-change': change,
			'document-error': undefined,
			'page-change': undefined,
			'dirty-change': undefined,
		});
		options.onDocumentChange?.(model);
		expect(change).toHaveBeenCalledWith(model);
		expect(options.onDocumentError).toBeUndefined();
	});
});

describe('customisation, page and save bindings', () => {
	it('forwards showThumbnails, showToolbar and hiddenActions with defaults on reset', () => {
		const host = document.createElement('div');
		document.body.append(host);
		const binding = mountEditor(host, {
			showThumbnails: true,
			showToolbar: false,
			hiddenActions: ['bold'],
		});
		expect(binding.element.showThumbnails).toBe(true);
		expect(binding.element.showToolbar).toBe(false);
		expect(binding.element.hiddenActions).toEqual(['bold']);
		expect(binding.element.hasAttribute('show-thumbnails')).toBe(true);
		binding.update({});
		expect(binding.element.showThumbnails).toBe(false);
		expect(binding.element.showToolbar).toBe(true);
		expect(binding.element.hiddenActions).toEqual([]);
		binding.destroy();
	});
	it('forwards de, es and zh-CN locales and their region variants', () => {
		const host = document.createElement('div');
		document.body.append(host);
		const binding = mountEditor(host, { locale: 'de-DE' });
		expect(binding.element.locale).toBe('de');
		for (const [input, expected] of [
			['es-MX', 'es'],
			['zh-Hans', 'zh-CN'],
			['zh-TW', 'en'],
			['fr', 'fr'],
		] as const) {
			binding.update({ locale: input });
			expect(binding.element.locale).toBe(expected);
		}
		binding.destroy();
	});
	it('maps deprecated English labels to ids and warns once, not on every update', () => {
		const host = document.createElement('div');
		document.body.append(host);
		const binding = mountEditor(host, { hiddenActions: ['Bold', 'print'] });
		const warnings: string[] = [];
		binding.element.addEventListener('document-warning', (event) =>
			warnings.push((event as CustomEvent<string>).detail),
		);
		binding.update({ hiddenActions: ['Bold', 'print'] });
		expect(binding.element.hiddenActions).toEqual(['bold', 'print']);
		expect(warnings).toEqual([]);
		binding.update({ hiddenActions: ['Italic'] });
		expect(binding.element.hiddenActions).toEqual(['italic']);
		expect(warnings).toHaveLength(1);
		expect(warnings[0]).toContain("'Italic'");
		binding.destroy();
	});
	it('surfaces page-change and dirty-change and exposes save/download/markClean', async () => {
		const host = document.createElement('div');
		document.body.append(host);
		const onPageChange = vi.fn();
		const onDirtyChange = vi.fn();
		const binding = mountEditor(host, { onPageChange, onDirtyChange });
		binding.element.dispatchEvent(
			new CustomEvent('page-change', { detail: { page: 2, pageCount: 5 } }),
		);
		expect(onPageChange).toHaveBeenCalledWith({ page: 2, pageCount: 5 });
		expect(binding.dirty).toBe(false);
		binding.element.dispatchEvent(new CustomEvent('dirty-change', { detail: true }));
		expect(onDirtyChange).toHaveBeenCalledWith(true);
		expect(await binding.save()).toBeInstanceOf(Blob);
		binding.markClean();
		expect(binding.dirty).toBe(false);
		expect(typeof binding.download).toBe('function');
		binding.destroy();
		binding.element.dispatchEvent(new CustomEvent('dirty-change', { detail: false }));
		expect(onDirtyChange).toHaveBeenCalledTimes(1);
	});
	it('lists the new keys and events once', () => {
		expect(EDITOR_PROP_KEYS).toEqual(
			expect.arrayContaining(['showThumbnails', 'showToolbar', 'hiddenActions']),
		);
		expect(EDITOR_EVENT_NAMES).toEqual(expect.arrayContaining(['page-change', 'dirty-change']));
	});
});
