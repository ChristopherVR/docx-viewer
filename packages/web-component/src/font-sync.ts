import { resolveRunFormatting } from 'docx-core';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { EditorState } from 'prosemirror-state';
import { findLocalizedControl } from './localization';
import { setComboValue, type ComboInput } from './ribbon-combo';
import { runOf, styleModelOf, themeFontOf } from './run-styles';
import { schema } from './schema';

const DEFAULT_FAMILY = 'Calibri';
const DEFAULT_SIZE = 11;

/** The family and size text shows in the editor: direct marks, then character/paragraph style and theme. */
export function effectiveFont(
	state: EditorState,
	text: ProseMirrorNode,
	paragraph: ProseMirrorNode,
): { family: string; size: number } {
	const model = styleModelOf(state);
	const run = runOf(text);
	if (!run || !model) {
		const font = text.marks.find((mark) => mark.type === schema.marks.font);
		return {
			family: String(font?.attrs.family ?? DEFAULT_FAMILY),
			size: Number(font?.attrs.size ?? DEFAULT_SIZE),
		};
	}
	const resolved = resolveRunFormatting(run, {
		runCatalog: model.characterStyles,
		paragraphCatalog: model.paragraphStyles,
		paragraphStyleId: paragraph.attrs.style || undefined,
	});
	return {
		family: resolved.fontFamily ?? themeFontOf(resolved, model.theme) ?? DEFAULT_FAMILY,
		size: resolved.fontSize ?? DEFAULT_SIZE,
	};
}

/** Shows the selection's font family and size in the two combo boxes (blank when they differ). */
export function syncFontControls(toolbar: HTMLElement, state: EditorState): void {
	const families = new Set<string>();
	const sizes = new Set<number>();
	const add = (text: ProseMirrorNode, paragraph: ProseMirrorNode) => {
		const { family, size } = effectiveFont(state, text, paragraph);
		families.add(family);
		sizes.add(size);
	};
	const { selection, doc } = state;
	if (selection.empty) {
		const marks = state.storedMarks ?? selection.$from.marks();
		add(schema.text('x', marks), selection.$from.parent);
	} else {
		let found = false;
		doc.nodesBetween(selection.from, selection.to, (node, _pos, parent) => {
			if (!node.isText || !parent) return;
			found = true;
			add(node, parent);
		});
		if (!found) add(schema.text('x', selection.$from.marks()), selection.$from.parent);
	}
	const family = findLocalizedControl<ComboInput>(toolbar, 'Font family');
	const size = findLocalizedControl<ComboInput>(toolbar, 'Font size');
	if (family?.comboItems)
		setComboValue(family, families.size === 1 ? ([...families][0] ?? null) : null);
	if (size?.comboItems) setComboValue(size, sizes.size === 1 ? String([...sizes][0]) : null);
}
