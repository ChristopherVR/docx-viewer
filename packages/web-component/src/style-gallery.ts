import { resolveRunFormatting, type DocumentModel, type TextRun } from '@christophervr/docx-core';

/** One tile: a style id (`''` is the paragraph's own default) and the name shown under its sample. */
export interface GalleryStyle {
	id: string;
	name: string;
}

export type StyleKind = 'paragraph' | 'character';

const SAMPLE = 'AaBbCc';
const HEX = /^#[0-9a-f]{6}$/i;

/** Styles Word keeps out of the quick gallery (they are reachable in the Styles pane). */
const HIDDEN_FROM_GALLERY =
	/^(toc\s*\d|table of|footnote|endnote|header|footer|hyperlink|followed|index|caption|annotation|balloon|line number|page number|plain table|table grid)/i;

/** The paragraph styles Word shows as gallery tiles: the default first, technical styles left out. */
export function recommendedStyles(all: readonly GalleryStyle[]): GalleryStyle[] {
	return all.filter((style, index) => index === 0 || !HIDDEN_FROM_GALLERY.test(style.name));
}

const HIDDEN_CHARACTER =
	/(char$|reference$|default paragraph font|hyperlink|page number|line number|placeholder)/i;

/** The character styles Word offers in the gallery (Emphasis, Strong, ...), not linked or technical ones. */
export function recommendedCharacterStyles(all: readonly GalleryStyle[]): GalleryStyle[] {
	return all.filter((style) => !HIDDEN_CHARACTER.test(style.name));
}

/** Applies the style's resolved run formatting to `element` (a scaled preview, not a layout). */
function previewStyle(
	element: HTMLElement,
	id: string,
	model: DocumentModel,
	kind: StyleKind,
): void {
	const run = resolveRunFormatting(
		{ text: SAMPLE, ...(kind === 'character' && id ? { style: id } : {}) } as TextRun,
		{
			...(kind === 'paragraph' && id ? { paragraphStyleId: id } : {}),
			paragraphCatalog: model.paragraphStyles,
			runCatalog: model.characterStyles,
		},
	);
	if (run.fontFamily) element.style.fontFamily = run.fontFamily;
	if (run.fontSize) element.style.fontSize = `${Math.min(20, Math.max(11, run.fontSize * 1.2))}px`;
	if (run.bold) element.style.fontWeight = '700';
	if (run.italic) element.style.fontStyle = 'italic';
	if (run.underline) element.style.textDecoration = 'underline';
	if (run.color && HEX.test(run.color)) element.style.color = run.color;
}

/** A tile previewing `style` on paper; `choose` receives the style id when it is clicked. */
export function createStyleTile(
	style: GalleryStyle,
	model: DocumentModel,
	kind: StyleKind,
	choose: (id: string) => void,
): HTMLButtonElement {
	const tile = document.createElement('button');
	tile.type = 'button';
	tile.className = 'style-tile';
	tile.dataset.styleId = style.id;
	tile.dataset.styleKind = kind;
	tile.setAttribute('aria-label', style.name);
	tile.title = style.name;
	tile.addEventListener('mousedown', (event) => event.preventDefault());
	tile.addEventListener('click', () => choose(style.id));
	const sample = document.createElement('span');
	sample.className = 'style-sample';
	sample.textContent = SAMPLE;
	previewStyle(sample, style.id, model, kind);
	const name = document.createElement('span');
	name.className = 'style-name';
	name.textContent = style.name;
	tile.append(sample, name);
	return tile;
}

/** Marks the tile for `selected` pressed and (un)disables all of them. */
export function markSelectedTiles(root: ParentNode, selected: string, disabled: boolean): void {
	for (const tile of root.querySelectorAll<HTMLButtonElement>('.style-tile')) {
		tile.setAttribute('aria-pressed', String(tile.dataset.styleId === selected));
		tile.disabled = disabled;
	}
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
			...styles.map((style) => createStyleTile(style, model, 'paragraph', choose)),
		);
	}
	markSelectedTiles(gallery, selected, disabled);
	gallery
		.querySelector('[aria-pressed="true"]')
		?.scrollIntoView?.({ block: 'nearest', inline: 'nearest' });
}
