import { resolveRunFormatting } from 'docx-core';
import type { EditorState } from 'prosemirror-state';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import { runOf, styleModelOf } from './run-styles';
import { schema } from './schema';

export type Script = 'none' | 'superscript' | 'subscript';

/** Effective script formatting, including paragraph/character styles and typing marks. */
export function selectionScript(state: EditorState): Script | null {
	const model = styleModelOf(state);
	const values = new Set<Script>();
	const add = (text: ProseMirrorNode, paragraph: ProseMirrorNode) => {
		const run = runOf(text);
		const align =
			run && model
				? resolveRunFormatting(run, {
						runCatalog: model.characterStyles,
						paragraphCatalog: model.paragraphStyles,
						paragraphStyleId: paragraph.attrs.style || undefined,
					}).verticalAlign
				: run?.verticalAlign;
		values.add(align === 'superscript' || align === 'subscript' ? align : 'none');
	};
	if (!state.selection.empty)
		state.doc.nodesBetween(state.selection.from, state.selection.to, (node, _pos, parent) => {
			if (node.isText && parent) add(node, parent);
		});
	if (!values.size)
		add(
			schema.text('x', state.storedMarks ?? state.selection.$from.marks()),
			state.selection.$from.parent,
		);
	return values.size === 1 ? [...values][0]! : null;
}
