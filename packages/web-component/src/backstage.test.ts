// @vitest-environment jsdom
import { createDocument } from '@christophervr/docx-core';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createBackstage, type BackstageHandlers } from './backstage';
import { documentStats } from './document-stats';
import { localizeElement } from './localization';

afterEach(() => (document.body.innerHTML = ''));

function setup(overrides: Partial<BackstageHandlers> = {}) {
	const model = createDocument();
	model.blocks = [{ type: 'paragraph', id: 'p', runs: [{ text: 'Hello there' }] }];
	const calls: Array<[string, string | undefined]> = [];
	const options: Array<[string, string]> = [];
	const handlers: BackstageHandlers = {
		fileCommand: (command, name) => calls.push([command, name]),
		close: vi.fn(),
		summary: () => ({
			fileName: 'Report.docx',
			model,
			words: 2,
			stats: documentStats(model),
			saveState: 'dirty',
		}),
		options: () => ({ locale: 'en', theme: 'auto', author: 'Ann' }),
		setOption: (key, value) => options.push([key, value]),
		...overrides,
	};
	const backstage = createBackstage(handlers);
	document.body.append(backstage.element);
	return { backstage, handlers, calls, options };
}
const nav = (root: HTMLElement) =>
	[...root.querySelectorAll('.dve-backstage-nav > button')].map((b) => b.textContent);
const click = (root: HTMLElement, label: string) =>
	[...root.querySelectorAll<HTMLButtonElement>('.dve-backstage-content button')]
		.find((b) => b.textContent === label)!
		.click();

describe('File backstage', () => {
	it("lists Word's file actions in order", () => {
		const { backstage } = setup();
		expect(nav(backstage.element)).toEqual([
			'Back to document',
			'Info',
			'New',
			'Open',
			'Save',
			'Save As',
			'Print',
			'Export',
			'Options',
		]);
		expect(backstage.element.querySelectorAll('[role="separator"]')).toHaveLength(2);
	});

	it('Info shows the file, its state and the properties panel', () => {
		const { backstage } = setup();
		backstage.open('info');
		const text = backstage.element.textContent!;
		expect(text).toContain('Report.docx');
		expect(text).toContain('Unsaved changes');
		expect(text).toContain('Characters (with spaces)');
		expect(backstage.element.querySelector('.dve-backstage-properties')!.textContent).toContain(
			'11',
		);
		expect(text).toContain('Letter'.toUpperCase());
	});

	it('Save runs the save command and closes', () => {
		const { backstage, calls, handlers } = setup();
		backstage.open();
		[...backstage.element.querySelectorAll<HTMLButtonElement>('.dve-backstage-nav button')]
			.find((b) => b.textContent === 'Save')!
			.click();
		expect(calls).toEqual([['save', undefined]]);
		expect(handlers.close).toHaveBeenCalled();
	});

	it('Save As sends the chosen file name, and ignores an empty one', () => {
		const { backstage, calls } = setup();
		backstage.open('saveAs');
		const name = backstage.element.querySelector<HTMLInputElement>(
			'input[aria-label="File name"]',
		)!;
		expect(name.value).toBe('Report.docx');
		name.value = 'Final draft';
		click(backstage.element, 'Save');
		expect(calls).toEqual([['saveAs', 'Final draft']]);
		backstage.open('saveAs');
		const again = backstage.element.querySelector<HTMLInputElement>(
			'input[aria-label="File name"]',
		)!;
		again.value = '   ';
		click(backstage.element, 'Save');
		expect(calls).toHaveLength(1);
	});

	it('Export offers PDF (via print), DOCX and plain text', () => {
		const { backstage, calls } = setup();
		backstage.open('export');
		click(backstage.element, 'Create PDF');
		backstage.open('export');
		click(backstage.element, 'Save a copy as DOCX');
		backstage.open('export');
		click(backstage.element, 'Export as plain text');
		expect(calls.map(([command]) => command)).toEqual(['print', 'export', 'exportText']);
	});

	it('Print shows paper facts and prints', () => {
		const { backstage, calls } = setup();
		backstage.open('print');
		expect(backstage.element.textContent).toContain('Paper size');
		click(backstage.element, 'Print');
		expect(calls).toEqual([['print', undefined]]);
	});

	it('Options reads the current values and reports changes', async () => {
		const { backstage, options } = setup();
		backstage.open('options');
		const language = backstage.element.querySelector<HTMLSelectElement>(
			'select[aria-label="Display language"]',
		)!;
		expect(language.value).toBe('en');
		language.value = 'de';
		language.dispatchEvent(new Event('change'));
		const author = backstage.element.querySelector<HTMLInputElement>(
			'input[aria-label="Author name"]',
		)!;
		expect(author.value).toBe('Ann');
		author.value = 'Bea';
		author.dispatchEvent(new Event('change'));
		expect(options).toEqual([
			['locale', 'de'],
			['author', 'Bea'],
		]);
	});

	it('follows the display language', () => {
		const { backstage } = setup();
		backstage.element.dataset.editorLocale = 'fr';
		backstage.open('saveAs');
		expect(backstage.element.querySelector('h2')!.textContent).toBe('Enregistrer sous');
		localizeElement(backstage.element, 'fr');
		expect(backstage.element.querySelector('.dve-backstage-nav')!.textContent).toContain(
			'Exporter',
		);
	});
});
