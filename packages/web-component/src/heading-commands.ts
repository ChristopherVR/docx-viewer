import type { DocumentModel } from 'docx-core';
import { closeHistory } from 'prosemirror-history';
import type { EditorView } from 'prosemirror-view';

export type HeadingLevel = 0 | 1 | 2 | 3;

/** The style id of Word's built-in "heading N" in this document, if it defines one. */
export function headingStyleId(model: DocumentModel, level: 1 | 2 | 3): string | undefined {
	const styles = Object.values(model.paragraphStyles?.styles ?? {});
	const wanted = new RegExp(`^heading\\s*${level}$`, 'i');
	return styles.find((style) => wanted.test(style.name ?? '') || wanted.test(style.id))?.id;
}

/**
 * References > Add Text: makes the selected paragraphs level 1-3 headings (so the table of
 * contents collects them), or with level 0 returns them to the default paragraph style. Returns
 * false when the document has no such heading style or the view is read-only.
 */
export function setHeadingLevel(
	view: EditorView,
	model: DocumentModel,
	level: HeadingLevel,
): boolean {
	if (!view.editable) return false;
	const style = level === 0 ? '' : headingStyleId(model, level);
	if (style === undefined) return false;
	let tr = view.state.tr;
	view.state.doc.nodesBetween(view.state.selection.from, view.state.selection.to, (node, pos) => {
		if (node.type.name === 'paragraph') tr = tr.setNodeAttribute(pos, 'style', style);
	});
	if (!tr.docChanged) return false;
	view.dispatch(closeHistory(tr));
	return true;
}
