import type { Command } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { closeHistory } from 'prosemirror-history';
import type {
	DocumentModel,
	NumberingCatalog,
	NumberingLevelDefinition,
} from '@christophervr/docx-core';
import {
	createListDefinition,
	linkStylesToList,
	ensureListDefinition,
	resolveNumberingLevel,
} from '@christophervr/docx-core';
import type { Node as ProseMirrorNode } from 'prosemirror-model';

export type ListKind = 'bullet' | 'decimal';

function paragraphPositions(view: EditorView): number[] {
	const { state } = view;
	const positions: number[] = [];
	state.doc.nodesBetween(state.selection.from, state.selection.to, (node, pos) => {
		if (node.type.name === 'paragraph') positions.push(pos);
	});
	if (!positions.length && state.selection.$from.parent.type.name === 'paragraph')
		positions.push(state.selection.$from.before());
	return positions;
}

/** Classifies a paragraph's effective list kind from its numId/level, or `null` when not in a list. */
export function listKindOf(
	node: ProseMirrorNode,
	catalog: NumberingCatalog | undefined,
): ListKind | 'other' | null {
	if (node.attrs.numId == null) return null;
	const level = catalog
		? resolveNumberingLevel(catalog, String(node.attrs.numId), node.attrs.ilvl ?? 0)
		: undefined;
	if (!level) return 'other';
	return level.numFmt === 'bullet' ? 'bullet' : 'decimal';
}

/** Whether every paragraph touched by the selection is already the requested list kind. */
export function selectionIsListKind(
	view: EditorView,
	kind: ListKind,
	catalog: NumberingCatalog | undefined,
): boolean {
	const positions = paragraphPositions(view);
	if (!positions.length) return false;
	return positions.every((pos) => listKindOf(view.state.doc.nodeAt(pos)!, catalog) === kind);
}

/**
 * Toggles the requested list kind on every paragraph in the selection. The caller decides
 * `alreadyKind` (by checking `selectionIsListKind` against the live numbering catalog before
 * calling in): when true, numbering is removed from the selection; otherwise `ensureNumId` is
 * invoked to obtain a fresh, independent `numId` (the caller registers it in the numbering
 * catalog first) and it is applied to every selected paragraph, preserving any existing level.
 */
export function toggleList(
	view: EditorView,
	alreadyKind: boolean,
	ensureNumId: () => number,
): void {
	if (!view.editable) return;
	const positions = paragraphPositions(view);
	if (!positions.length) return;
	let transaction = closeHistory(view.state.tr);
	if (alreadyKind) {
		for (const pos of positions) {
			const node = transaction.doc.nodeAt(pos);
			if (node)
				transaction = transaction.setNodeMarkup(pos, undefined, {
					...node.attrs,
					numId: null,
					ilvl: null,
				});
		}
	} else {
		const numId = ensureNumId();
		for (const pos of positions) {
			const node = transaction.doc.nodeAt(pos);
			if (!node) continue;
			const level = node.attrs.numId != null ? (node.attrs.ilvl ?? 0) : 0;
			transaction = transaction.setNodeMarkup(pos, undefined, {
				...node.attrs,
				numId,
				ilvl: level,
			});
		}
	}
	view.dispatch(transaction.scrollIntoView());
}

/** Removes list numbering from every paragraph touched by the selection. */
export function removeList(view: EditorView): boolean {
	if (!view.editable) return false;
	const positions = paragraphPositions(view).filter(
		(pos) => view.state.doc.nodeAt(pos)?.attrs.numId != null,
	);
	if (!positions.length) return false;
	let transaction = closeHistory(view.state.tr);
	for (const pos of positions) {
		const node = transaction.doc.nodeAt(pos);
		if (node)
			transaction = transaction.setNodeMarkup(pos, undefined, {
				...node.attrs,
				numId: null,
				ilvl: null,
			});
	}
	view.dispatch(transaction.scrollIntoView());
	return true;
}

/** Changes the outline level (`ilvl`, clamped 0..8) of every list paragraph touched by the selection. */
export function changeListLevel(view: EditorView, delta: 1 | -1): boolean {
	if (!view.editable) return false;
	const positions = paragraphPositions(view).filter(
		(pos) => view.state.doc.nodeAt(pos)?.attrs.numId != null,
	);
	if (!positions.length) return false;
	let transaction = closeHistory(view.state.tr);
	for (const pos of positions) {
		const node = transaction.doc.nodeAt(pos);
		if (!node) continue;
		const level = Math.min(8, Math.max(0, Number(node.attrs.ilvl ?? 0) + delta));
		transaction = transaction.setNodeMarkup(pos, undefined, { ...node.attrs, ilvl: level });
	}
	view.dispatch(transaction.scrollIntoView());
	return true;
}

/** Word behavior: pressing Enter on an empty list item ends the list instead of adding another item. */
export const exitListOnEmptyEnter: Command = (state, dispatch, view) => {
	if (view && !view.editable) return false;
	const node = state.selection.$from.parent;
	if (node.type.name !== 'paragraph' || node.attrs.numId == null || node.content.size !== 0)
		return false;
	if (!dispatch) return true;
	const pos = state.selection.$from.before();
	dispatch(state.tr.setNodeMarkup(pos, undefined, { ...node.attrs, numId: null, ilvl: null }));
	return true;
};

function levelCommand(delta: 1 | -1): Command {
	return (state, dispatch, view) => {
		if (view && !view.editable) return false;
		const node = state.selection.$from.parent;
		if (node.type.name !== 'paragraph' || node.attrs.numId == null) return false;
		if (!dispatch) return true;
		const pos = state.selection.$from.before();
		const level = Math.min(8, Math.max(0, Number(node.attrs.ilvl ?? 0) + delta));
		dispatch(state.tr.setNodeMarkup(pos, undefined, { ...node.attrs, ilvl: level }));
		return true;
	};
}

/** Tab/Shift+Tab inside a list item change its outline level; otherwise the command is a no-op. */
export const indentListItem: Command = levelCommand(1);
export const outdentListItem: Command = levelCommand(-1);

export type ListAction =
	| 'bullet'
	| 'number'
	| 'multilevel'
	| 'outline'
	| 'increaseLevel'
	| 'decreaseLevel'
	| 'remove';

/** A definition and all selected list paragraphs are applied through one history transaction. */
export function applyCustomList(
	view: EditorView,
	model: DocumentModel,
	levels: readonly NumberingLevelDefinition[],
): void {
	toggleList(view, false, () => {
		const created = createListDefinition(model.numberingCatalog, levels);
		model.numberingCatalog = created.catalog;
		const linked = linkStylesToList(model.paragraphStyles, levels, created.numId);
		if (linked) model.paragraphStyles = linked;
		return created.numId;
	});
}

/**
 * Runs a Home > Paragraph list command. Applying a new list adds a fresh numbering definition to
 * `model.numberingCatalog` so unrelated lists never share counters.
 */
export function runListAction(view: EditorView, key: ListAction, model: DocumentModel): void {
	if (key === 'remove') removeList(view);
	else if (key === 'increaseLevel') changeListLevel(view, 1);
	else if (key === 'decreaseLevel') changeListLevel(view, -1);
	else if (key === 'multilevel' || key === 'outline') {
		// A gallery choice always starts a new list of that style, like Word.
		toggleList(view, false, () => {
			const created = ensureListDefinition(model.numberingCatalog, key);
			model.numberingCatalog = created.catalog;
			return created.numId;
		});
	} else {
		const kind = key === 'bullet' ? 'bullet' : 'decimal';
		const already = selectionIsListKind(view, kind, model.numberingCatalog);
		toggleList(view, already, () => {
			const created = ensureListDefinition(model.numberingCatalog, kind);
			model.numberingCatalog = created.catalog;
			return created.numId;
		});
	}
}
