import type { NodeSpec } from 'prosemirror-model';

/** A visible, editable page or column break marker (`w:br` type page/column). */
export const pageBreakNodeSpec: NodeSpec = {
	group: 'inline',
	inline: true,
	atom: true,
	selectable: true,
	attrs: { kind: { default: 'page' } },
	leafText: () => '',
	parseDOM: [
		{
			tag: 'span[data-docx-break]',
			getAttrs: (el) => ({
				kind: (el as HTMLElement).dataset.docxBreak === 'column' ? 'column' : 'page',
			}),
		},
	],
	toDOM: (node) => [
		'span',
		{ class: 'dve-break-marker', 'data-docx-break': node.attrs.kind, contenteditable: 'false' },
		node.attrs.kind === 'column' ? 'Column Break' : 'Page Break',
	],
};

/** A read-only footnote/endnote reference mark; its number is derived from document order. */
export const noteReferenceNodeSpec: NodeSpec = {
	group: 'inline',
	inline: true,
	atom: true,
	selectable: false,
	attrs: {
		kind: { default: 'footnote' },
		id: { default: '' },
		number: { default: 1 },
		/** The number as shown, in the document's footnote/endnote number format. */
		label: { default: null },
		/** The reference run's own formatting (e.g. superscript, FootnoteReference style) as JSON.
		 *  Kept as an attribute, not marks, so text typed beside a reference doesn't inherit it. */
		format: { default: null },
	},
	leafText: (node) => String(node.attrs.label ?? node.attrs.number),
	parseDOM: [
		{
			tag: 'sup[data-docx-note-kind]',
			getAttrs: (el) => ({
				kind: (el as HTMLElement).dataset.docxNoteKind === 'endnote' ? 'endnote' : 'footnote',
				id: (el as HTMLElement).dataset.docxNoteId || '',
				number: Number((el as HTMLElement).textContent) || 1,
			}),
		},
	],
	toDOM: (node) => [
		'sup',
		{
			class: 'dve-note-reference',
			'data-docx-note-kind': node.attrs.kind,
			'data-docx-note-id': node.attrs.id,
			contenteditable: 'false',
		},
		String(node.attrs.label ?? node.attrs.number),
	],
};
