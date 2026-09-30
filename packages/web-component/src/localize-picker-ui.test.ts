// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { openColorGridPopover } from './ribbon-color-grid';
import { galleryLook, menuGallery } from './ribbon-gallery-menu';
import { createZoomDialog } from './zoom-dialog';
import { translate } from './localization';

afterEach(() => (document.body.innerHTML = ''));

function frenchHost() {
	const host = document.createElement('div');
	host.dataset.editorLocale = 'fr';
	document.body.append(host);
	const anchor = document.createElement('button');
	host.append(anchor);
	return { host, anchor };
}

describe('pickers follow the display language', () => {
	it('names the colour picker sections and swatches in French', () => {
		const { anchor } = frenchHost();
		openColorGridPopover(anchor, () => {}, { noneLabel: 'No Color' });
		const pop = document.querySelector('.color-grid')!;
		const text = pop.textContent!;
		expect(text).toContain('Couleurs du thème');
		expect(text).toContain('Couleurs standard');
		expect(text).toContain('Autres couleurs…');
		expect(text).toContain('Aucune couleur');
		expect(pop.querySelector('[aria-label="Rouge foncé"]')).not.toBeNull();
		expect(pop.querySelector('[aria-label="Bleu, Accentuation 1"]')).not.toBeNull();
	});

	it('translates gallery details and the trailing command', () => {
		const { host } = frenchHost();
		const say = (text: string) => translate('fr', text as never);
		expect(galleryLook('margins', 'narrow', say).detail).toBe(
			'Haut: 0.5"  Bas: 0.5"  Gauche: 0.5"  Droite: 0.5"',
		);
		const menu = menuGallery(
			'Line spacing',
			'lineSpacing',
			'lineSpacing',
			[['auto:240', '1.0']],
			(value) => ({ type: 'paragraph', key: 'lineSpacing', value }),
			{
				compact: true,
				commands: [
					{ label: 'Line Spacing Options…', action: { type: 'formatDialog', kind: 'paragraph' } },
				],
			},
		);
		host.append(menu);
		menu.click();
		expect(document.querySelector('.gallery-command')!.textContent).toBe('Options d’interligne…');
	});

	it('shows the Zoom dialog in French', () => {
		const { host } = frenchHost();
		const dialog = createZoomDialog({
			percent: () => 100,
			setPercent: () => {},
			fit: () => {},
			restoreFocus: () => {},
		});
		dialog.setLocale('fr');
		host.append(dialog.element);
		dialog.open();
		const text = dialog.element.textContent!;
		expect(text).toContain('Zoom sur');
		expect(text).toContain('Page entière');
		expect(text).toContain('Deux pages');
		expect(text).toContain('Aperçu');
		expect(dialog.element.getAttribute('aria-label')).toBe('Paramètres de zoom');
	});
});
