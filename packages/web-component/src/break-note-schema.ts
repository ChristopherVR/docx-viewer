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

/**
 * A complex field's structural run: a `w:fldChar` marker (begin/separate/end) or its instruction
 * text (`code`). Invisible while editing; its run formatting rides in `format` like note references.
 */
export const fieldMarkerNodeSpec: NodeSpec = {
	group: 'inline',
	inline: true,
	atom: true,
	selectable: false,
	attrs: { kind: { default: 'begin' }, code: { default: null }, format: { default: null } },
	leafText: () => '',
	parseDOM: [
		{
			tag: 'span[data-field-marker]',
			getAttrs: (el) => ({
				kind: (el as HTMLElement).dataset.fieldMarker || 'begin',
				code: (el as HTMLElement).dataset.fieldCode ?? null,
			}),
		},
	],
	toDOM: (node) => [
		'span',
		{
			class: 'dve-field-marker',
			'data-field-marker': node.attrs.kind,
			...(node.attrs.code != null ? { 'data-field-code': node.attrs.code } : {}),
			contenteditable: 'false',
		},
	],
};
