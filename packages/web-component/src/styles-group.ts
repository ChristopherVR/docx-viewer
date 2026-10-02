import type { DocumentModel } from 'docx-core';
import type { EditorView } from 'prosemirror-view';
import { characterStyleAtSelection } from './character-style-picker';
import { emit } from './events';
import { translate, translateTemplate, type EditorLocale } from './localization';
import { ribbonIcon, type RibbonIcon } from './ribbon-icons';
import { launcher, row } from './ribbon-parts';
import { closeRibbonPopover } from './ribbon-popover';
import {
	recommendedCharacterStyles,
	recommendedStyles,
	syncStyleGallery,
	type GalleryStyle,
} from './style-gallery';
import { createStylesPane, type StylesPane } from './styles-pane';
import { openStylesGallery, type StylesData } from './styles-popover';

interface Latest {
	view: EditorView;
	model: DocumentModel;
	locale: EditorLocale;
}

/** Scrolls the gallery by roughly one tile and a half per press. */
const TILE_STEP = 170;

function stripButton(label: string, icon: RibbonIcon): HTMLButtonElement {
	const button = document.createElement('button');
	button.type = 'button';
	button.className = 'style-strip-button';
	button.setAttribute('aria-label', label);
	button.title = label;
	button.append(ribbonIcon(icon, 12));
	button.addEventListener('mousedown', (event) => event.preventDefault());
	return button;
}

function build(toolbar: HTMLElement, latest: Latest): HTMLElement {
	const { locale } = latest;
	const group = document.createElement('div');
	group.className = 'ribbon-group';
	group.dataset.label = 'Styles';
	group.dataset.caption = translate(locale, 'Styles');
	group.dataset.stylesGroup = '';
	group.setAttribute('role', 'group');
	group.setAttribute(
		'aria-label',
		translateTemplate(locale, 'ribbon.groupControls', { group: group.dataset.caption }),
	);
	const gallery = document.createElement('div');
	gallery.className = 'ribbon-gallery';
	gallery.setAttribute('role', 'group');
	gallery.dataset.styleGallery = '';
	gallery.addEventListener(
		'wheel',
		(event) => {
			if (Math.abs(event.deltaY) <= Math.abs(event.deltaX)) return;
			gallery.scrollLeft += event.deltaY;
			event.preventDefault();
		},
		{ passive: false },
	);
	const previous = stripButton(translate(locale, 'Previous styles'), 'chevronLeft');
	const next = stripButton(translate(locale, 'Next styles'), 'chevronRight');
	const more = stripButton(translate(locale, 'More styles'), 'stylesMoreArrow');
	previous.addEventListener('click', () =>
		gallery.scrollBy({ left: -TILE_STEP, behavior: 'smooth' }),
	);
	next.addEventListener('click', () => gallery.scrollBy({ left: TILE_STEP, behavior: 'smooth' }));
	const strip = document.createElement('div');
	strip.className = 'style-strip';
	strip.append(previous, next, more);
	// The select is the styles' state holder (value, disabled, options for every style); the
	// gallery, the expanded gallery and the pane are how a person changes it.
	const state = document.createElement('div');
	state.hidden = true;
	state.dataset.styleState = '';
	const select = document.createElement('select');
	select.dataset.paragraphStyles = '';
	state.append(select);
	const pane = createStylesPane(() => (paneLauncher.ariaPressed = 'false'));
	const paneLauncher = launcher(translate(locale, 'Styles pane'), { type: 'search' });
	paneLauncher.removeAttribute('data-action');
	paneLauncher.setAttribute('aria-pressed', 'false');
	paneLauncher.addEventListener('click', () => {
		if (pane.isOpen) pane.close();
		else {
			pane.open();
			paneLauncher.setAttribute('aria-pressed', 'true');
			render();
		}
	});
	toolbar.parentElement?.querySelector('.dve-body')?.append(pane.element);
	group.append(row(gallery, strip), state, paneLauncher);

	const apply = (value: string) => {
		const { view, model } = latest;
		if (!view.editable) return;
		if (value && !model.paragraphStyles?.styles[value]) return;
		let tr = view.state.tr;
		view.state.doc.nodesBetween(view.state.selection.from, view.state.selection.to, (node, pos) => {
			if (node.type.name === 'paragraph') tr = tr.setNodeAttribute(pos, 'style', value);
		});
		if (tr.docChanged) view.dispatch(tr);
		view.focus();
	};
	select.addEventListener('change', () => apply(select.value));
	gallery.addEventListener('dve-choose-style', (event) =>
		apply((event as CustomEvent<string>).detail),
	);
	const data = (): StylesData => currentData(group, latest, apply);
	function render() {
		pane.render(data(), latest.locale);
	}
	Object.assign(group, { renderPane: render, stylesPane: pane as StylesPane });
	more.addEventListener('click', () => {
		if (more.getAttribute('aria-expanded') === 'true') return closeRibbonPopover();
		openStylesGallery(more, data(), closeRibbonPopover);
	});
	return group;
}

/** The state the gallery, popover and pane draw from, for the current view and model. */
function currentData(
	group: HTMLElement,
	latest: Latest,
	applyParagraph: (id: string) => void,
): StylesData {
	const { view, model, locale } = latest;
	const definitions = Object.values(model.paragraphStyles?.styles ?? {});
	const normal = definitions.find((style) => style.isDefault);
	const all: GalleryStyle[] = [
		{ id: '', name: normal?.name || translate(locale, 'styles.inherit') },
		...definitions
			.filter((style) => !style.isDefault)
			.map((style) => ({ id: style.id, name: style.name || style.id })),
	];
	const select = group.querySelector<HTMLSelectElement>('[data-paragraph-styles]')!;
	const characters = Object.values(model.characterStyles?.styles ?? {})
		.filter((style) => style.type === 'character')
		.map((style) => ({ id: style.id, name: style.name || style.id }));
	return {
		paragraph: recommendedStyles(all),
		character: recommendedCharacterStyles(characters),
		allParagraph: all,
		allCharacter: characters,
		selectedParagraph: String(view.state.selection.$from.parent.attrs.style ?? ''),
		selectedCharacter: characterStyleAtSelection(view),
		model,
		disabled: select.disabled,
		chooseParagraph: applyParagraph,
		chooseCharacter: (value) => emit(group, 'ribbon-action', { type: 'characterStyle', value }),
		clearFormatting: () => emit(group, 'ribbon-action', { type: 'clear' }),
		openPane: () => {
			(group as HTMLElement & { stylesPane: StylesPane }).stylesPane.open();
			group.querySelector('.ribbon-launcher')?.setAttribute('aria-pressed', 'true');
			(group as HTMLElement & { renderPane(): void }).renderPane();
		},
	};
}

/** Builds the Styles group on first use and keeps its gallery, tiles and state select current. */
export function syncStylePicker(
	toolbar: HTMLElement,
	view: EditorView,
	model: DocumentModel,
	locale: EditorLocale,
): void {
	const latest: Latest = { view, model, locale };
	let group = toolbar.querySelector<HTMLElement & { latest?: Latest }>('[data-styles-group]');
	if (!group) {
		group = build(toolbar, latest) as HTMLElement & { latest?: Latest };
		group.latest = latest;
		const home = toolbar.querySelector('#dve-panel-home');
		home?.insertBefore(group, home.querySelector('.ribbon-group[data-label="Editing"]'));
	}
	Object.assign(group.latest ?? {}, latest);
	const select = group.querySelector<HTMLSelectElement>('[data-paragraph-styles]')!;
	select.setAttribute('aria-label', translate(locale, 'styles.label'));
	const definitions = Object.values(model.paragraphStyles?.styles ?? {});
	const selected = String(view.state.selection.$from.parent.attrs.style ?? '');
	const options = [{ id: '', name: translate(locale, 'styles.inherit') }, ...definitions];
	if (selected && !definitions.some((style) => style.id === selected))
		options.push({ id: selected, name: selected });
	select.replaceChildren(...options.map((style) => new Option(style.name || style.id, style.id)));
	select.value = selected;
	select.disabled = !view.editable || !definitions.length;
	const gallery = group.querySelector<HTMLElement>('[data-style-gallery]')!;
	const normal = definitions.find((style) => style.isDefault);
	const tiles = recommendedStyles([
		{ id: '', name: normal?.name || translate(locale, 'styles.inherit') },
		...definitions
			.filter((style) => !style.isDefault)
			.map((style) => ({ id: style.id, name: style.name || style.id })),
	]);
	syncStyleGallery(
		gallery,
		tiles,
		selected,
		model,
		(id) => {
			select.value = id;
			gallery.dispatchEvent(new CustomEvent('dve-choose-style', { detail: id }));
		},
		select.disabled,
	);
	(group as HTMLElement & { renderPane?(): void }).renderPane?.();
}

/** Removes the Styles group and its pane, so a new document builds them afresh. */
export function resetStylePicker(toolbar?: HTMLElement): void {
	const group = toolbar?.querySelector<HTMLElement & { stylesPane?: StylesPane }>(
		'[data-styles-group]',
	);
	group?.stylesPane?.element.remove();
	group?.remove();
}
