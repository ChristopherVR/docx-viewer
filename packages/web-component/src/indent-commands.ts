import type { EditorState } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { findLocalizedControl } from './localization';

export type IndentSide = 'left' | 'right';

const TWIPS_PER_INCH = 1440;
const MAX_INCHES = 22;

/** The attribute that holds a side's indent: the logical (start/end) one when the paragraph uses it. */
function indentAttr(attrs: Record<string, unknown>, side: IndentSide): string {
	if (side === 'left')
		return attrs.indentStartTwips != null ? 'indentStartTwips' : 'indentLeftTwips';
	return attrs.indentEndTwips != null ? 'indentEndTwips' : 'indentRightTwips';
}

/** Paragraphs the command applies to: every one touched by the selection, else the caret's. */
function paragraphPositions(state: EditorState): number[] {
	const positions: number[] = [];
	state.doc.nodesBetween(state.selection.from, state.selection.to, (node, pos) => {
		if (node.type.name === 'paragraph') positions.push(pos);
	});
	if (!positions.length && state.selection.$from.parent.type.name === 'paragraph')
		positions.push(state.selection.$from.before());
	return positions;
}

/** Sets the left or right indent of the selected paragraphs, in inches (Word's default unit). */
export function setIndent(view: EditorView, side: IndentSide, inches: number): boolean {
	if (!view.editable || !Number.isFinite(inches)) return false;
	const twips = Math.round(Math.min(MAX_INCHES, Math.max(0, inches)) * TWIPS_PER_INCH);
	let tr = view.state.tr;
	for (const pos of paragraphPositions(view.state)) {
		const node = tr.doc.nodeAt(pos);
		if (node)
			tr = tr.setNodeMarkup(pos, undefined, {
				...node.attrs,
				[indentAttr(node.attrs, side)]: twips,
			});
	}
	if (!tr.docChanged) return false;
	view.dispatch(tr);
	return true;
}

/** The common indent of the selected paragraphs in inches, or null when they differ. */
export function indentInches(state: EditorState, side: IndentSide): number | null {
	const values = new Set(
		paragraphPositions(state).map((pos) => {
			const attrs = state.doc.nodeAt(pos)?.attrs ?? {};
			return Number(attrs[indentAttr(attrs, side)] ?? 0) / TWIPS_PER_INCH;
		}),
	);
	return values.size === 1 ? ([...values][0] ?? 0) : null;
}

/** Shows the selection's indents in the Layout tab's two spinners. */
export function syncIndentInputs(toolbar: HTMLElement, state: EditorState): void {
	for (const [label, side] of [
		['Indent left', 'left'],
		['Indent right', 'right'],
	] as const) {
		const input = findLocalizedControl<HTMLInputElement>(toolbar, label);
		if (!input || document.activeElement === input) continue;
		const value = indentInches(state, side);
		input.value = value === null ? '' : String(Math.round(value * 100) / 100);
	}
}
