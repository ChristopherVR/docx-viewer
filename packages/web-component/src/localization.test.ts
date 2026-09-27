// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { EditorState } from 'prosemirror-state';
import { schema } from './schema';
import { createRibbon, setRibbonLocale } from './ribbon';
import { createSearchPanel } from './search-panel';
import { normalizeEditorLocale, translate } from './localization';

describe('shared editor localization', () => {
	it('normalizes display locale with English fallback and dictionary coverage', () => {
		expect(normalizeEditorLocale('fr-CA')).toBe('fr');
		expect(normalizeEditorLocale('de-DE')).toBe('en');
		expect(translate('fr', 'status.words')).toBe('{count} mots');
	});

	it('localizes ribbon labels while keeping action payloads and control state intact', () => {
		const ribbon = createRibbon();
		const format = ribbon.querySelector<HTMLButtonElement>('[aria-label="Bold"]')!;
		const highlight = ribbon.querySelector<HTMLSelectElement>('[aria-label="Text highlight"]')!;
		highlight.value = 'cyan';
		setRibbonLocale(ribbon, 'fr');
		expect(ribbon.getAttribute('aria-label')).toBe('Mise en forme du document');
		expect(ribbon.querySelector('[role="tab"]')?.textContent).toBe('Accueil');
		expect(format.getAttribute('aria-label')).toBe('Gras');
		expect(JSON.parse(format.dataset.action!).key).toBe('bold');
		expect(highlight.value).toBe('cyan');
		expect(highlight.options[0].textContent).toBe('Sans surlignage');
		setRibbonLocale(ribbon, 'en');
		expect(format.getAttribute('aria-label')).toBe('Bold');
		expect(highlight.value).toBe('cyan');
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
