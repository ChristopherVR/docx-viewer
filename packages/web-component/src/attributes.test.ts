// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { DocxEditorElement, registerDocxEditor } from './index';

registerDocxEditor();
afterEach(() => document.body.replaceChildren());
const make = () => document.createElement('docx-editor') as DocxEditorElement;

describe('attribute reflection', () => {
	it('observes exactly the reflected attributes', () => {
		expect(DocxEditorElement.observedAttributes).toEqual([
			'locale',
			'read-only',
			'file-name',
			'review-author',
			'theme',
		]);
	});

	it('attributes drive properties', () => {
		const editor = make();
		document.body.append(editor);
		editor.setAttribute('locale', 'fr-FR');
		editor.setAttribute('read-only', '');
		editor.setAttribute('file-name', 'Plan.docx');
		editor.setAttribute('review-author', 'Ada');
		expect(editor.locale).toBe('fr');
		expect(editor.readOnly).toBe(true);
		expect(editor.fileName).toBe('Plan.docx');
		expect(editor.reviewAuthor).toBe('Ada');
		editor.removeAttribute('read-only');
		editor.removeAttribute('file-name');
		editor.removeAttribute('review-author');
		expect(editor.readOnly).toBe(false);
		expect(editor.fileName).toBe('Document1.docx');
		expect(editor.reviewAuthor).toBe('Author');
	});

	it('properties reflect to attributes, including before connection', () => {
		const editor = make();
		editor.locale = 'fr';
		editor.readOnly = true;
		editor.fileName = 'A.docx';
		editor.reviewAuthor = 'Grace';
		expect(editor.getAttribute('locale')).toBe('fr');
		expect(editor.hasAttribute('read-only')).toBe(true);
		expect(editor.getAttribute('file-name')).toBe('A.docx');
		expect(editor.getAttribute('review-author')).toBe('Grace');
		editor.readOnly = false;
		expect(editor.hasAttribute('read-only')).toBe(false);
	});

	it('does not loop and keeps the author markup when it already means the same value', () => {
		const editor = make();
		document.body.append(editor);
		const setter = vi.spyOn(editor, 'setAttribute');
		editor.setAttribute('locale', 'FR-fr');
		expect(editor.locale).toBe('fr');
		expect(editor.getAttribute('locale')).toBe('FR-fr');
		editor.locale = 'fr'; // same locale: no attribute write beyond the author's own
		expect(setter).toHaveBeenCalledTimes(1);
		editor.readOnly = true;
		editor.readOnly = true;
		expect(editor.hasAttribute('read-only')).toBe(true);
	});

	it('theme attribute and property stay in sync', () => {
		const editor = make();
		editor.setAttribute('theme', 'dark');
		expect(editor.theme).toBe('dark');
		editor.theme = 'light';
		expect(editor.getAttribute('theme')).toBe('light');
		editor.setAttribute('theme', 'bogus');
		expect(editor.theme).toBe('auto');
	});

	it('an unsupported locale attribute falls back through the property normalizer', () => {
		const editor = make();
		editor.setAttribute('locale', 'xx-YY');
		expect(editor.locale).toBe('en');
	});

	it('a chrome-driven read-only change updates the attribute', () => {
		const editor = make();
		document.body.append(editor);
		editor.readOnly = true;
		expect(editor.hasAttribute('read-only')).toBe(true);
	});

	it('upgrades attributes present in markup', () => {
		document.body.innerHTML =
			'<docx-editor read-only locale="fr" file-name="M.docx"></docx-editor>';
		const editor = document.body.firstElementChild as DocxEditorElement;
		expect(editor.readOnly).toBe(true);
		expect(editor.locale).toBe('fr');
		expect(editor.fileName).toBe('M.docx');
	});
});
