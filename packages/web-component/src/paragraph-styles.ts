import { Plugin } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { Decoration, DecorationSet } from 'prosemirror-view';
import {
	resolveParagraphFormatting,
	type DocumentModel,
	type Paragraph,
} from '@christophervr/docx-core';
import { paragraphStyle } from './schema';
import { translate, type EditorLocale } from './localization';

/** Derived display formatting stays out of document attributes and collaboration steps. */
export function paragraphStylesPlugin(getModel: () => DocumentModel) {
	return new Plugin({
		props: {
			decorations(state) {
				const catalog = getModel().paragraphStyles;
				if (!catalog) return null;
				const decorations: Decoration[] = [];
				state.doc.descendants((node, pos) => {
					if (node.type.name !== 'paragraph') return;
					const direct = Object.fromEntries(
						Object.entries(node.attrs).filter(
							([key, value]) => value != null && (key !== 'style' || value !== ''),
						),
					);
					const effective = resolveParagraphFormatting(
						{ ...direct, id: String(node.attrs.id), type: 'paragraph', runs: [] } as Paragraph,
						catalog,
					);
					decorations.push(
						Decoration.node(pos, pos + node.nodeSize, {
							style: paragraphStyle(effective),
							...(effective.direction ? { dir: effective.direction } : {}),
						}),
					);
				});
				return DecorationSet.create(state.doc, decorations);
			},
		},
	});
}

export function syncStylePicker(
	toolbar: HTMLElement,
	view: EditorView,
	model: DocumentModel,
	locale: EditorLocale,
) {
	let select = toolbar.querySelector<HTMLSelectElement>('[data-paragraph-styles]');
	if (!select) {
		select = document.createElement('select');
		select.dataset.paragraphStyles = '';
		const group = document.createElement('div');
		group.className = 'ribbon-group';
		group.dataset.label = 'Styles';
		group.dataset.caption = translate(locale, 'Styles');
		group.setAttribute('role', 'group');
		group.setAttribute(
			'aria-label',
			locale === 'fr'
				? `Commandes : ${group.dataset.caption}`
				: `${group.dataset.caption} controls`,
		);
		group.append(select);
		toolbar.querySelector('#dve-panel-home')?.prepend(group);
		select.addEventListener('change', () => {
			if (!view.editable) return;
			const value = select!.value;
			if (value && !model.paragraphStyles?.styles[value]) return;
			let tr = view.state.tr;
			view.state.doc.nodesBetween(
				view.state.selection.from,
				view.state.selection.to,
				(node, pos) => {
					if (node.type.name === 'paragraph') tr = tr.setNodeAttribute(pos, 'style', value);
				},
			);
			if (tr.docChanged) view.dispatch(tr);
			view.focus();
		});
	}
	// Replace the listener closure on each new view through a view-scoped picker.
	select.setAttribute('aria-label', translate(locale, 'styles.label'));
	const definitions = Object.values(model.paragraphStyles?.styles ?? {});
	const selected = view.state.selection.$from.parent.attrs.style ?? '';
	const options = [{ id: '', name: translate(locale, 'styles.inherit') }, ...definitions];
	if (selected && !definitions.some((style) => style.id === selected))
		options.push({ id: selected, name: selected });
	select.replaceChildren(...options.map((style) => new Option(style.name || style.id, style.id)));
	select.value = selected;
	select.disabled = !view.editable || !definitions.length;
}

export function resetStylePicker(toolbar?: HTMLElement) {
	toolbar?.querySelector('[data-paragraph-styles]')?.parentElement?.remove();
}
