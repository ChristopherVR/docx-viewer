import type { DocumentModel } from '@christophervr/docx-core';
import { localeOf, translate } from './localization';
import { mountPopover } from './ribbon-popover';
import { ribbonIcon } from './ribbon-icons';
import { createStyleTile, markSelectedTiles, type GalleryStyle } from './style-gallery';

/** Everything the expanded gallery and the Styles pane show and can do. */
export interface StylesData {
	/** Recommended styles, for the gallery and its expanded view. */
	paragraph: GalleryStyle[];
	character: GalleryStyle[];
	/** Every style, for the Styles pane. */
	allParagraph: GalleryStyle[];
	allCharacter: GalleryStyle[];
	selectedParagraph: string;
	selectedCharacter: string;
	model: DocumentModel;
	disabled: boolean;
	chooseParagraph(id: string): void;
	chooseCharacter(id: string): void;
	clearFormatting(): void;
	openPane(): void;
}

function section(title: string, tiles: HTMLElement[]): HTMLElement {
	const wrap = document.createElement('div');
	wrap.className = 'styles-section';
	const heading = document.createElement('div');
	heading.className = 'styles-section-title';
	heading.textContent = title;
	const grid = document.createElement('div');
	grid.className = 'styles-grid';
	grid.append(...tiles);
	wrap.append(heading, grid);
	return wrap;
}

function commandButton(label: string, icon: Parameters<typeof ribbonIcon>[0], run: () => void) {
	const button = document.createElement('button');
	button.type = 'button';
	button.className = 'styles-command';
	button.append(ribbonIcon(icon, 16));
	const text = document.createElement('span');
	text.textContent = label;
	button.append(text);
	button.addEventListener('mousedown', (event) => event.preventDefault());
	button.addEventListener('click', run);
	return button;
}

/**
 * Word's expanded Styles gallery: every paragraph style and every character style as preview
 * tiles, then Clear Formatting and the Styles pane. Choosing a tile applies it and closes.
 */
export function openStylesGallery(
	anchor: HTMLElement,
	data: StylesData,
	onClose: () => void,
): void {
	const locale = localeOf(anchor);
	const pop = document.createElement('div');
	pop.className = 'ribbon-popover styles-popover';
	pop.setAttribute('role', 'dialog');
	pop.setAttribute('aria-label', translate(locale, 'Styles'));
	const finish = (run: () => void) => () => {
		run();
		onClose();
	};
	pop.append(
		section(
			translate(locale, 'Paragraph styles'),
			data.paragraph.map((style) =>
				createStyleTile(
					style,
					data.model,
					'paragraph',
					finish(() => data.chooseParagraph(style.id)),
				),
			),
		),
	);
	if (data.character.length)
		pop.append(
			section(
				translate(locale, 'Character styles'),
				[{ id: '', name: translate(locale, 'No character style') }, ...data.character].map(
					(style) =>
						createStyleTile(
							style,
							data.model,
							'character',
							finish(() => data.chooseCharacter(style.id)),
						),
				),
			),
		);
	const commands = document.createElement('div');
	commands.className = 'styles-commands';
	commands.append(
		commandButton(
			translate(locale, 'Clear formatting'),
			'clear',
			finish(() => data.clearFormatting()),
		),
		commandButton(
			translate(locale, 'Styles pane'),
			'stylesPane',
			finish(() => data.openPane()),
		),
	);
	pop.append(commands);
	for (const grid of pop.querySelectorAll<HTMLElement>('.styles-grid')) {
		const kind = grid.querySelector<HTMLElement>('.style-tile')?.dataset.styleKind;
		markSelectedTiles(
			grid,
			kind === 'character' ? data.selectedCharacter : data.selectedParagraph,
			data.disabled,
		);
	}
	for (const button of commands.querySelectorAll('button'))
		(button as HTMLButtonElement).disabled = data.disabled && button.textContent !== 'Styles pane';
	mountPopover(anchor, pop, anchor, 'end');
	pop.querySelector<HTMLElement>('[aria-pressed="true"]')?.focus({ preventScroll: true });
}
