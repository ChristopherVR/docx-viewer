// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { EditorState } from 'prosemirror-state';
import { schema } from './schema';
import { createRibbon, setRibbonLocale } from './ribbon';
import { createSearchPanel } from './search-panel';
import { normalizeEditorLocale, translate } from './localization';
import { at } from './test-support';

describe('shared editor localization', () => {
	it('normalizes display locale with English fallback and dictionary coverage', () => {
		expect(normalizeEditorLocale('fr-CA')).toBe('fr');
		expect(normalizeEditorLocale('de-DE')).toBe('de');
		expect(normalizeEditorLocale('pt-BR')).toBe('en');
		expect(translate('fr', 'status.words')).toBe('{count} mots');
	});

	it('localizes ribbon labels while keeping action payloads and control state intact', () => {
		const ribbon = createRibbon();
		const format = ribbon.querySelector<HTMLButtonElement>('[aria-label="Bold"]')!;
		const zoom = ribbon.querySelector<HTMLSelectElement>('[aria-label="Zoom"]')!;
		zoom.value = '125';
		setRibbonLocale(ribbon, 'fr');
		expect(ribbon.getAttribute('aria-label')).toBe('Mise en forme du document');
		expect(ribbon.querySelector('[role="tab"]')?.textContent).toBe('Accueil');
		expect(format.getAttribute('aria-label')).toBe('Gras');
		expect(ribbon.querySelector<HTMLButtonElement>('[aria-label="Liste à puces"]')).not.toBeNull();
		expect(ribbon.querySelector<HTMLElement>('[data-label="Font"]')?.dataset.caption).toBe(
			'Police',
		);
		expect(JSON.parse(format.dataset.action!).key).toBe('bold');
		expect(zoom.value).toBe('125');
		expect(ribbon.querySelector<HTMLElement>('[data-label="Clipboard"]')?.dataset.caption).toBe(
			'Presse-papiers',
		);
		expect(ribbon.querySelector('[data-label="Clipboard"] .ribbon-large')?.textContent).toBe(
			'Coller',
		);
		expect(ribbon.querySelector('.ribbon-split [data-split-caret]')?.getAttribute('title')).toBe(
			'Options de surlignage',
		);
		setRibbonLocale(ribbon, 'en');
		expect(format.getAttribute('aria-label')).toBe('Bold');
		expect(ribbon.querySelector<HTMLElement>('[data-label="Font"]')?.dataset.caption).toBe('Font');
		expect(zoom.value).toBe('125');
	});

	it('localizes find status and labels without changing document-language content', () => {
		const doc = schema.node('doc', null, [
			schema.nodes.paragraph.create({ id: 'p' }, schema.text('bonjour')),
		]);
		const panel = createSearchPanel({ getView: () => undefined }, 'fr');
		expect(panel.element.getAttribute('aria-label')).toBe('Rechercher et remplacer');
		expect(panel.element.querySelector('[role="status"]')?.textContent).toBe(
			'Saisissez le texte à rechercher',
		);
		panel.setLocale('en');
		expect(panel.element.getAttribute('aria-label')).toBe('Find and replace');
		expect(EditorState.create({ doc }).doc.firstChild?.attrs.id).toBe('p');
	});
});
