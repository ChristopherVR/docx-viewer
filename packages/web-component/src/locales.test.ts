// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { createDocument } from '@christophervr/docx-core';
import { DocxEditorElement, registerDocxEditor } from './index';
import {
	EDITOR_LOCALES,
	localizeElement,
	normalizeEditorLocale,
	translate,
	translateTemplate,
	type EditorLocale,
	type LocalizationKey,
} from './localization';
import { strings } from './localization-strings';
import { RIBBON_ACTION_LABELS } from './ribbon-action-ids';
import { createRibbon, setRibbonLocale } from './ribbon';

registerDocxEditor();

const englishKeys = Object.keys(strings.en) as LocalizationKey[];
const placeholders = (text: string) => [...text.matchAll(/\{\w+\}/g)].map((m) => m[0]).sort();
/** Terms that are legitimately identical to English in some locale (loanwords, symbols, units). */
const MAY_MATCH_ENGLISH =
	/^(OK|Zoom|Normal|Layout|Magenta|Orange|Gold|Original|Info|Link|Format|Cyan|Portrait|Table|Editor)/i;

describe('locale catalogue', () => {
	it('lists every locale and one strings module for each', () => {
		expect([...EDITOR_LOCALES]).toEqual(['en', 'fr', 'de', 'es', 'zh-CN']);
		expect(Object.keys(strings).sort()).toEqual([...EDITOR_LOCALES].sort());
	});

	for (const locale of EDITOR_LOCALES) {
		describe(locale, () => {
			it('defines every key with no extras and no empty strings', () => {
				expect(Object.keys(strings[locale]).sort()).toEqual([...englishKeys].sort());
				for (const key of englishKeys) {
					const value = strings[locale][key];
					expect(typeof value, key).toBe('string');
					expect(value.trim(), key).not.toBe('');
				}
			});

			it('keeps every {placeholder} of the English template', () => {
				for (const key of englishKeys)
					expect(placeholders(strings[locale][key]), `${locale}:${key}`).toEqual(
						placeholders(strings.en[key]),
					);
			});

			it('translates the display text of every ribbon action label', () => {
				for (const label of Object.values(RIBBON_ACTION_LABELS))
					expect(label in strings.en, label).toBe(true);
			});

			if (locale !== 'en')
				it('does not leave ordinary UI text in English', () => {
					const same = englishKeys.filter(
						(key) =>
							strings[locale][key] === strings.en[key] &&
							!key.includes('.') &&
							!MAY_MATCH_ENGLISH.test(strings.en[key]) &&
							!/^[^A-Za-z]*$/.test(strings.en[key]),
					);
					// Whole-word loanwords that Word itself keeps (e.g. German "Layout", "Original").
					expect(same.filter((key) => key.split(' ').length > 2)).toEqual([]);
				});
		});
	}
});

describe('normalizeEditorLocale', () => {
	it.each([
		['en', 'en'],
		['en-GB', 'en'],
		['fr', 'fr'],
		['fr-CA', 'fr'],
		['FR_be', 'fr'],
		['de', 'de'],
		['de-DE', 'de'],
		['de-AT', 'de'],
		['es', 'es'],
		['es-MX', 'es'],
		['es_ES', 'es'],
		['zh', 'zh-CN'],
		['zh-CN', 'zh-CN'],
		['zh_cn', 'zh-CN'],
		['zh-Hans', 'zh-CN'],
		['zh-Hans-CN', 'zh-CN'],
		['zh-SG', 'zh-CN'],
		[' de-DE ', 'de'],
		['zh-TW', 'en'],
		['zh-Hant', 'en'],
		['zh-HK', 'en'],
		['ja', 'en'],
		['', 'en'],
		[undefined, 'en'],
		[null, 'en'],
	] as const)('%s -> %s', (input, expected) => {
		expect(normalizeEditorLocale(input)).toBe(expected);
	});
});

describe('localized rendering', () => {
	const sample: Array<[EditorLocale, string, string]> = [
		['de', 'Fett', 'Dokumenteditor'],
		['es', 'Negrita', 'Editor de documentos'],
		['zh-CN', '加粗', '文档编辑器'],
	];
	it.each(sample)('%s ribbon and shell labels follow the display locale', (locale, bold, host) => {
		const editor = document.createElement('docx-editor') as DocxEditorElement;
		editor.documentModel = createDocument();
		document.body.append(editor);
		editor.locale = locale === 'de' ? 'de-DE' : locale;
		expect(editor.locale).toBe(locale);
		expect(editor.getAttribute('locale')).toBe(locale);
		const root = editor.shadowRoot!;
		const control = root.querySelector<HTMLElement>('[data-localearialabel="Bold"]')!;
		expect(control.getAttribute('aria-label')).toBe(bold);
		expect(editor.getAttribute('aria-label')).toBe(host);
		expect(root.querySelector('.dve-paper')!.getAttribute('aria-label')).toBe(
			translate(locale, 'Document page'),
		);
		expect(root.querySelector('.dve-file-input')!.getAttribute('aria-label')).toBe(
			translate(locale, 'Open'),
		);
		editor.locale = 'en';
		expect(editor.getAttribute('aria-label')).toBe('Document editor');
		expect(control.getAttribute('aria-label')).toBe('Bold');
		editor.remove();
	});

	it('localizes the File tab and file inputs when the locale is set before connecting', () => {
		const editor = document.createElement('docx-editor') as DocxEditorElement;
		editor.locale = 'es-MX';
		document.body.append(editor);
		const root = editor.shadowRoot!;
		expect(root.querySelector('.dve-file-tab')!.textContent).toBe('Archivo');
		expect(root.querySelector('.dve-picture-input')!.getAttribute('aria-label')).toBe(
			'Insertar imagen',
		);
		editor.remove();
	});

	it('localizes group captions, table glyph buttons, spacing options and language names', () => {
		const ribbon = createRibbon('de');
		const group = ribbon.querySelector<HTMLElement>('[data-label="Tables"]')!;
		expect(group.getAttribute('aria-label')).toBe('Steuerelemente: Tabellen');
		expect(ribbon.querySelector('[data-localearialabel="Insert table"]')!.textContent).toBe(
			'▦ Tabelle',
		);
		const after = ribbon.querySelector<HTMLSelectElement>(
			'[data-localearialabel="Spacing after"]',
		)!;
		expect([...after.options].map((option) => option.textContent)).toContain('Nach: 12 Pt.');
		const language = ribbon.querySelector<HTMLSelectElement>(
			'[data-localearialabel="Text language"]',
		)!;
		const german = language.querySelector('option[value="de-DE"]')!;
		expect(german.textContent).toBe('Deutsch (Deutschland)');
		expect(ribbon.querySelector('[data-localearialabel="Increase indent"]')).not.toBeNull();
		const zh = createRibbon('zh-CN');
		expect(
			zh.querySelector('[data-localearialabel="Text language"] option[value="de-DE"]')!.textContent,
		).toBe('德语（德国）');
	});

	it('names ribbon tabs with tab-specific keys, distinct from dialog actions', () => {
		const expected = {
			en: ['Home', 'Insert', 'Layout', 'References', 'Review', 'View', 'Table'],
			fr: ['Accueil', 'Insertion', 'Disposition', 'Références', 'Révision', 'Affichage', 'Tableau'],
			de: ['Start', 'Einfügen', 'Layout', 'Verweise', 'Überprüfen', 'Ansicht', 'Tabelle'],
			es: ['Inicio', 'Insertar', 'Diseño', 'Referencias', 'Revisar', 'Vista', 'Tabla'],
			'zh-CN': ['开始', '插入', '布局', '引用', '审阅', '视图', '表格'],
		} as const;
		for (const locale of EDITOR_LOCALES) {
			const ribbon = createRibbon(locale);
			const names = [...ribbon.querySelectorAll('[role=tab]')].map((tab) => tab.textContent);
			expect(names, locale).toEqual([...expected[locale]]);
		}
		expect(strings.fr['tab.insert']).toBe('Insertion');
		expect(strings.fr.Insert).toBe('Insérer');
		const french = createRibbon('en');
		setRibbonLocale(french, 'fr');
		expect(french.querySelector('#dve-tab-insert')?.textContent).toBe('Insertion');
		setRibbonLocale(french, 'en');
		expect(french.querySelector('#dve-tab-insert')?.textContent).toBe('Insert');
	});

	it('translates dynamic line-spacing labels and templates in every locale', () => {
		expect(translateTemplate('de', 'status.page', { current: 2, total: 5 })).toBe('Seite 2 von 5');
		expect(translateTemplate('es', 'presence.cursor', { name: 'Ana' })).toBe('Cursor de Ana');
		const host = document.createElement('div');
		const texts = ['1.0 lines', '1.15 lines', 'Automatic 1.1 lines (264/240)', 'At least 12 pt'];
		for (const text of texts)
			host.append(Object.assign(document.createElement('p'), { textContent: text }));
		localizeElement(host, 'de');
		expect([...host.children].map((child) => child.textContent)).toEqual([
			'1.0 Zeile',
			'1.15 Zeilen',
			'Automatisch 1.1 Zeilen (264/240)',
			'Mindestens 12 Pt.',
		]);
		localizeElement(host, 'en');
		expect([...host.children].map((child) => child.textContent)).toEqual(texts);
	});
});
