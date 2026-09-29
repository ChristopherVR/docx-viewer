import { closeHistory } from 'prosemirror-history';
import { TextSelection, type EditorState } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { schema } from './schema';

export type BorderPreset =
	| 'horizontal'
	| 'none'
	| 'bottom'
	| 'top'
	| 'left'
	| 'right'
	| 'all'
	| 'outside'
	| 'insideH';

type Side = 'top' | 'bottom' | 'left' | 'right' | 'between';
type BorderSide = { style: string; sizeEighthPoints: number; color?: string; spacePoints?: number };
type Borders = Partial<Record<Side, BorderSide>>;

/** Word's default border pen: a single 0.5 pt automatic-colour line, 1 pt from the text. */
const PEN: BorderSide = { style: 'single', sizeEighthPoints: 4, spacePoints: 1 };

const SIDES_OF: Record<Exclude<BorderPreset, 'none' | 'horizontal'>, Side[]> = {
	bottom: ['bottom'],
	top: ['top'],
	left: ['left'],
	right: ['right'],
	all: ['top', 'bottom', 'left', 'right'],
	outside: ['top', 'bottom', 'left', 'right'],
	insideH: ['between'],
};

function selectedParagraphs(state: EditorState): number[] {
	const positions: number[] = [];
	state.doc.nodesBetween(state.selection.from, state.selection.to, (node, pos) => {
		if (node.type.name === 'paragraph') positions.push(pos);
	});
	if (!positions.length && state.selection.$from.parent.type.name === 'paragraph')
		positions.push(state.selection.$from.before());
	return positions;
}

const bordersOf = (attrs: Record<string, unknown>): Borders =>
	typeof attrs.borders === 'object' && attrs.borders ? (attrs.borders as Borders) : {};

/** Sets the shading fill of the selected paragraphs to `#rrggbb`, or with `null` removes it. */
export function setShading(view: EditorView, value: string | null): boolean {
	if (!view.editable) return false;
	if (value !== null && !/^#[0-9a-f]{6}$/i.test(value)) return false;
	let tr = view.state.tr;
	for (const pos of selectedParagraphs(view.state)) {
		const node = tr.doc.nodeAt(pos);
		if (node)
			tr = tr.setNodeMarkup(pos, undefined, {
				...node.attrs,
				shadingFill: value ? value.toUpperCase() : null,
			});
	}
	if (!tr.docChanged) return false;
	view.dispatch(closeHistory(tr));
	return true;
}

/**
 * Borders > Horizontal Line: a new empty paragraph with a bottom border after the current one,
 * followed by an empty paragraph for the caret. It is saved as an ordinary paragraph border.
 */
export function insertHorizontalLine(view: EditorView): boolean {
	if (!view.editable) return false;
	const { $from } = view.state.selection;
	const at = $from.parent.type.name === 'paragraph' ? $from.after() : view.state.doc.content.size;
	const line = schema.nodes.paragraph!.create({
		borders: { bottom: { ...PEN, sizeEighthPoints: 6 } },
	});
	const after = schema.nodes.paragraph!.create();
	const tr = view.state.tr.insert(at, [line, after]);
	view.dispatch(
		closeHistory(
			tr.setSelection(TextSelection.create(tr.doc, at + line.nodeSize + 1)),
		).scrollIntoView(),
	);
	return true;
}

/**
 * Word's Borders menu. A side preset toggles: when every selected paragraph already has all of its
 * sides the borders are removed, otherwise they are added. `none` removes every border.
 */
export function setBorders(view: EditorView, preset: BorderPreset): boolean {
	if (!view.editable) return false;
	if (preset === 'horizontal') return insertHorizontalLine(view);
	const positions = selectedParagraphs(view.state);
	const sides = preset === 'none' ? [] : SIDES_OF[preset];
	const allHave =
		preset !== 'none' &&
		positions.every((pos) => {
			const borders = bordersOf(view.state.doc.nodeAt(pos)?.attrs ?? {});
			return sides.every((side) => borders[side] && borders[side]?.style !== 'none');
		});
	let tr = view.state.tr;
	for (const pos of positions) {
		const node = tr.doc.nodeAt(pos);
		if (!node) continue;
		let next: Borders | null;
		if (preset === 'none') next = null;
		else {
			const merged: Borders = { ...bordersOf(node.attrs) };
			for (const side of sides)
				if (allHave) delete merged[side];
				else merged[side] = { ...PEN };
			next = Object.keys(merged).length ? merged : null;
		}
		tr = tr.setNodeMarkup(pos, undefined, { ...node.attrs, borders: next });
	}
	if (!tr.docChanged) return false;
	view.dispatch(closeHistory(tr));
	return true;
}
