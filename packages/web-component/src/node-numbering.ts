import { resolveParagraphNumbering, type DocumentModel } from '@christophervr/docx-core';
import type { Node } from 'prosemirror-model';

/** Resolve a live paragraph's direct or inherited list without adding model properties. */
export function nodeNumbering(node: Node, model: DocumentModel) {
	return resolveParagraphNumbering(
		{
			type: 'paragraph',
			id: node.attrs.id,
			runs: [],
			style: node.attrs.style,
			...(node.attrs.numId != null
				? { numbering: { numId: Number(node.attrs.numId), level: Number(node.attrs.ilvl ?? 0) } }
				: {}),
		},
		model.paragraphStyles,
		model.numberingCatalog,
	);
}
