import { Plugin } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { Decoration, DecorationSet } from 'prosemirror-view';
import {
	resolveParagraphFormatting,
	type DocumentModel,
	type Paragraph,
} from '@christophervr/docx-core';
import { paragraphStyle } from './schema';
import { translate, translateTemplate, type EditorLocale } from './localization';
import { menuAround, row, stack } from './ribbon-parts';
import { syncStyleGallery } from './style-gallery';

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

export { resetStylePicker, syncStylePicker } from './styles-group';
