import { closeHistory } from 'prosemirror-history';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { refreshFieldResults } from './field-update';
import { schema } from './schema';

/** Word's built-in caption labels; the SEQ identifier is the English label. */
export const CAPTION_LABELS = ['Figure', 'Table', 'Equation'] as const;
export type CaptionLabel = (typeof CAPTION_LABELS)[number];

export const seqInstruction = (label: string) => ` SEQ ${label} \\* ARABIC `;
const SEQ = /^\s*SEQ\s+(\S+)/i;

/** The label a text node's field numbers (`Figure` for `SEQ Figure ...`), or undefined. */
export function seqLabelOf(node: ProseMirrorNode): string | undefined {
	if (!node.isText) return undefined;
	const mark = node.marks.find((item) => item.type === schema.marks.field);
	return mark ? SEQ.exec(String(mark.attrs.instr ?? ''))?.[1] : undefined;
}

/**
 * Renumbers every `SEQ label` field in document order from 1, as Word's Update Field does, so a
 * caption inserted in the middle shifts the ones after it. Returns the transaction for chaining.
 */
export function renumberCaptions(tr: Transaction, label: string): Transaction {
	const fields: Array<{ from: number; to: number; node: ProseMirrorNode }> = [];
	tr.doc.descendants((node, pos) => {
		if (seqLabelOf(node)?.toLowerCase() === label.toLowerCase())
			fields.push({ from: pos, to: pos + node.nodeSize, node });
	});
	// Later fields first, so replacing one never shifts the positions still to be processed.
	for (let index = fields.length - 1; index >= 0; index--) {
		const { from, to, node } = fields[index]!;
		if (node.text !== String(index + 1))
			tr.replaceWith(from, to, schema.text(String(index + 1), node.marks));
	}
	return tr;
}

export interface CaptionOptions {
	/** Label identifier, e.g. `Figure`. */
	label: string;
	/** The label as shown in the document language, e.g. `Abbildung`. */
	labelText: string;
	/** Text after the number (`Figure 1: text`). */
	text: string;
	position: 'above' | 'below';
	/** The document's Caption style id, if it has one. */
	style?: string;
}

/**
 * Insert > Caption: a paragraph "Label N" (N is a `SEQ` field) with optional text, placed above
 * or below the block holding the selection, then every caption of that label is renumbered.
 * Without a Caption style the text is italic. Returns false for a read-only view.
 */
export function insertCaption(view: EditorView, options: CaptionOptions): boolean {
	if (!view.editable) return false;
	const { $from } = view.state.selection;
	const index = $from.index(0);
	let at = 0;
	view.state.doc.forEach((node, offset, i) => {
		if (i === index) at = options.position === 'above' ? offset : offset + node.nodeSize;
	});
	const italic = options.style ? [] : [schema.marks.italic!.create()];
	const number = schema.text('1', [
		...italic,
		schema.marks.field!.create({ instr: seqInstruction(options.label), simple: true }),
	]);
	const content = [
		schema.text(`${options.labelText} `, italic),
		number,
		...(options.text.trim() ? [schema.text(`: ${options.text.trim()}`, italic)] : []),
	];
	const paragraph = schema.nodes.paragraph!.create(
		options.style ? { style: options.style } : {},
		content,
	);
	const tr = view.state.tr.insert(at, paragraph);
	// Renumber every caption of the label, and bring REF results that quote them up to date.
	refreshFieldResults(tr, undefined, true);
	view.dispatch(closeHistory(tr).scrollIntoView());
	return true;
}
