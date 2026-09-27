import type { Node as ProseMirrorNode } from 'prosemirror-model';
import { Plugin, TextSelection, type EditorState, type Transaction } from 'prosemirror-state';

/** Field marker kinds in document order. */
function markerKinds(doc: ProseMirrorNode): string[] {
	const kinds: string[] = [];
	doc.descendants((node) => {
		if (node.type.name === 'fieldMarker') kinds.push(String(node.attrs.kind));
	});
	return kinds;
}

/**
 * Whether complex-field markers nest like Word expects: each `begin` is followed by its code, an
 * optional `separate` and a matching `end`. Fields may nest and may span paragraphs.
 */
export function fieldsBalanced(doc: ProseMirrorNode): boolean {
	const stack: ('code' | 'result')[] = [];
	for (const kind of markerKinds(doc)) {
		if (kind === 'begin') stack.push('code');
		else if (kind === 'code') {
			if (stack.at(-1) !== 'code') return false;
		} else if (kind === 'separate') {
			if (stack.at(-1) !== 'code') return false;
			stack[stack.length - 1] = 'result';
		} else if (kind === 'end') {
			if (!stack.length) return false;
			stack.pop();
		}
	}
	return stack.length === 0;
}

/**
 * Keeps complex fields structurally intact: an edit that would delete only part of a field (its
 * begin, separator or end marker) is rejected, while editing the result text or deleting a whole
 * field is allowed. Documents that arrive unbalanced are left alone.
 */
export function fieldGuardPlugin(): Plugin {
	return new Plugin({
		props: {
			handleTextInput: (view, from, to, text) => {
				const transaction = replaceAroundMarkers(view.state, from, to, text);
				if (transaction) view.dispatch(transaction);
				return Boolean(transaction);
			},
			handleKeyDown: (view, event) => {
				if (event.key !== 'Backspace' && event.key !== 'Delete') return false;
				const transaction = deleteAroundMarkers(view.state, event.key === 'Backspace');
				if (transaction) view.dispatch(transaction);
				return Boolean(transaction);
			},
		},
		appendTransaction: (transactions, _old, state) =>
			transactions.some((transaction) => transaction.docChanged)
				? markFieldResults(state)
				: transactions.some((transaction) => transaction.selectionSet)
					? trimSelection(state)
					: null,
		filterTransaction(transaction, state) {
			if (!transaction.docChanged) return true;
			let hadMarkers = false;
			state.doc.descendants((node) => {
				if (node.type.name === 'fieldMarker') hadMarkers = true;
				return !hadMarkers;
			});
			if (!hadMarkers || !fieldsBalanced(state.doc)) return true;
			return fieldsBalanced(transaction.doc);
		},
	});
}

/**
 * Tags text typed into a complex field's result (between its separator and end) with the field
 * mark, so replacing a whole result keeps it a field result as it does in Word.
 */
function markFieldResults(state: EditorState): Transaction | null {
	const fieldMark = state.schema.marks.field;
	const stack: { code: string; inResult: boolean }[] = [];
	const missing: { from: number; to: number; instr: string }[] = [];
	state.doc.descendants((node, pos) => {
		if (node.type.name === 'fieldMarker') {
			const kind = node.attrs.kind;
			if (kind === 'begin') stack.push({ code: '', inResult: false });
			else if (kind === 'code' && stack.length)
				stack[stack.length - 1].code += node.attrs.code ?? '';
			else if (kind === 'separate' && stack.length) stack[stack.length - 1].inResult = true;
			else if (kind === 'end') stack.pop();
			return false;
		}
		const top = stack.at(-1);
		if (node.isText && top?.inResult && !fieldMark.isInSet(node.marks))
			missing.push({ from: pos, to: pos + node.nodeSize, instr: top.code.trim() });
		return true;
	});
	if (!missing.length) return null;
	const transaction = state.tr;
	for (const { from, to, instr } of missing)
		transaction.addMark(from, to, fieldMark.create({ instr, simple: false }));
	return transaction;
}

const isMarker = (node: ProseMirrorNode | null | undefined) => node?.type.name === 'fieldMarker';

/** Drops hidden field markers from a text selection's edges, so a selected result is just its text. */
function trimSelection(state: EditorState): Transaction | null {
	const { selection, doc } = state;
	if (!(selection instanceof TextSelection) || selection.empty) return null;
	let { from, to } = selection;
	while (from < to && isMarker(doc.resolve(from).nodeAfter)) from++;
	while (to > from && isMarker(doc.resolve(to).nodeBefore)) to--;
	if (from === selection.from && to === selection.to) return null;
	const reversed = selection.anchor > selection.head;
	return state.tr.setSelection(
		TextSelection.create(doc, reversed ? to : from, reversed ? from : to),
	);
}

function markersIn(
	doc: ProseMirrorNode,
	from: number,
	to: number,
): { pos: number; size: number }[] {
	const markers: { pos: number; size: number }[] = [];
	doc.nodesBetween(from, to, (node, pos) => {
		if (isMarker(node) && pos >= from && pos + node.nodeSize <= to)
			markers.push({ pos, size: node.nodeSize });
	});
	return markers;
}

/**
 * Replaces `from..to` with `text` but keeps any field markers inside the range when removing them
 * would break a field, as Word does when a selection crosses a field boundary.
 */
export function replaceAroundMarkers(
	state: EditorState,
	from: number,
	to: number,
	text: string,
): Transaction | null {
	if (from === to) return null;
	const markers = markersIn(state.doc, from, to);
	if (!markers.length) return null;
	const whole = state.tr.delete(from, to);
	if (fieldsBalanced(whole.doc) || !fieldsBalanced(state.doc)) return null;
	const marks = state.doc.resolve(from).marksAcross(state.doc.resolve(to)) ?? [];
	const transaction = state.tr;
	let end = to;
	for (const marker of [...markers].reverse()) {
		const after = marker.pos + marker.size;
		if (after < end) transaction.delete(after, end);
		end = marker.pos;
	}
	if (from < end) transaction.delete(from, end);
	if (text) transaction.insert(from, state.schema.text(text, marks));
	return transaction.setSelection(TextSelection.create(transaction.doc, from + text.length));
}

/** Backspace/Delete that steps over hidden field markers instead of deleting them. */
export function deleteAroundMarkers(state: EditorState, backward: boolean): Transaction | null {
	const { selection, doc } = state;
	if (!(selection instanceof TextSelection)) return null;
	if (!selection.empty) return replaceAroundMarkers(state, selection.from, selection.to, '');
	let pos = selection.from;
	const step = () => (backward ? doc.resolve(pos).nodeBefore : doc.resolve(pos).nodeAfter);
	if (!isMarker(step())) return null;
	while (isMarker(step())) pos += backward ? -doc.resolve(pos).nodeBefore!.nodeSize : 1;
	const neighbor = step();
	if (!neighbor?.isText) return state.tr.setSelection(TextSelection.create(doc, pos));
	const transaction = backward ? state.tr.delete(pos - 1, pos) : state.tr.delete(pos, pos + 1);
	const caret = backward ? pos - 1 : pos;
	return transaction.setSelection(TextSelection.create(transaction.doc, caret));
}
