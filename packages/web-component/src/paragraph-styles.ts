import { Plugin } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { Decoration, DecorationSet } from 'prosemirror-view';
import {
	resolveParagraphFormatting,
	computeListLabels,
	displayListLabel,
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
				const model = getModel();
				const catalog = model.paragraphStyles;
				const paragraphs: Paragraph[] = [];
				state.doc.descendants((node) => {
					if (node.type.name === 'paragraph')
						paragraphs.push({
							type: 'paragraph',
							id: String(node.attrs.id),
							runs: [],
							style: node.attrs.style,
							...(node.attrs.numId != null
								? {
										numbering: {
											numId: Number(node.attrs.numId),
											level: Number(node.attrs.ilvl ?? 0),
										},
									}
								: {}),
						});
				});
				const labels = computeListLabels({ ...model, blocks: paragraphs });
				const decorations: Decoration[] = [];
				state.doc.descendants((node, pos) => {
					if (node.type.name !== 'paragraph') return;
					const direct = Object.fromEntries(
						Object.entries(node.attrs).filter(
							([key, value]) => value != null && (key !== 'style' || value !== ''),
						),
					);
					const paragraph = {
						...direct,
						id: String(node.attrs.id),
						type: 'paragraph',
						runs: [],
					} as Paragraph;
					const effective = catalog ? resolveParagraphFormatting(paragraph, catalog) : paragraph;
					const label = labels.get(String(node.attrs.id));
					const list = label
						? {
								listIndentLeftTwips: label.indentLeftTwips ?? null,
								listHangingTwips: label.hangingTwips ?? null,
								listFirstLineTwips: label.firstLineTwips ?? null,
							}
						: { listIndentLeftTwips: null, listHangingTwips: null, listFirstLineTwips: null };
					decorations.push(
						Decoration.node(pos, pos + node.nodeSize, {
							style: paragraphStyle({ ...effective, ...list }),
							...(label || node.attrs.listLabelText != null
								? {
										'data-list-label': label
											? `${displayListLabel(label.text)}${label.suffix === 'space' ? ' ' : label.suffix === 'none' ? '' : '\t'}`
											: '',
									}
								: {}),
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
