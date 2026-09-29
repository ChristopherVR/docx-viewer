import { resolveRunFormatting, type DocumentModel, type TextRun } from '@christophervr/docx-core';

/** One tile: a style id (`''` is the paragraph's own default) and the name shown under its sample. */
export interface GalleryStyle {
	id: string;
	name: string;
}

const SAMPLE = 'AaBbCc';
const HEX = /^#[0-9a-f]{6}$/i;

/** Applies the style's resolved run formatting to `element` (a scaled preview, not a layout). */
function previewStyle(element: HTMLElement, id: string, model: DocumentModel): void {
	const run = resolveRunFormatting({ text: SAMPLE } as TextRun, {
		paragraphStyleId: id || undefined,
		paragraphCatalog: model.paragraphStyles,
		runCatalog: model.characterStyles,
	});
	if (run.fontFamily) element.style.fontFamily = run.fontFamily;
	if (run.fontSize) element.style.fontSize = `${Math.min(20, Math.max(11, run.fontSize * 1.2))}px`;
	if (run.bold) element.style.fontWeight = '700';
	if (run.italic) element.style.fontStyle = 'italic';
	if (run.color && HEX.test(run.color)) element.style.color = run.color;
}

/**
 * Word's Styles gallery: a scrolling strip of tiles, each previewing its style. Tiles are rebuilt
 * only when the list changes; `selected` just moves the pressed state. `choose` receives the id.
 */
export function syncStyleGallery(
	gallery: HTMLElement,
	styles: readonly GalleryStyle[],
	selected: string,
	model: DocumentModel,
	choose: (id: string) => void,
	disabled: boolean,
): void {
	const key = JSON.stringify(styles);
	if (gallery.dataset.key !== key) {
		gallery.dataset.key = key;
		gallery.replaceChildren(
			...styles.map((style) => {
				const tile = document.createElement('button');
				tile.type = 'button';
				tile.className = 'style-tile';
				tile.dataset.styleId = style.id;
				tile.setAttribute('aria-label', style.name);
				tile.title = style.name;
				tile.addEventListener('mousedown', (event) => event.preventDefault());
				tile.addEventListener('click', () => choose(style.id));
				const sample = document.createElement('span');
				sample.className = 'style-sample';
				sample.textContent = SAMPLE;
				previewStyle(sample, style.id, model);
				const name = document.createElement('span');
				name.className = 'style-name';
				name.textContent = style.name;
				tile.append(sample, name);
				return tile;
			}),
		);
	}
	for (const tile of gallery.querySelectorAll<HTMLButtonElement>('.style-tile')) {
		tile.setAttribute('aria-pressed', String(tile.dataset.styleId === selected));
		tile.disabled = disabled;
	}
	gallery
		.querySelector('[aria-pressed="true"]')
		?.scrollIntoView?.({ block: 'nearest', inline: 'nearest' });
}
