import { closeHistory } from 'prosemirror-history';
import { NodeSelection, type EditorState } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { schema } from './schema';

export interface TextBoxSettings {
	lines: string[];
	widthPx: number;
	heightPx: number;
	border: boolean;
}

const lines = (value: unknown): string[] => {
	if (typeof value !== 'string') return [];
	try {
		const parsed: unknown = JSON.parse(value);
		return Array.isArray(parsed) ? parsed.map(String) : [];
	} catch {
		return [];
	}
};

/** The selected text box when it is the simple, editable form (inserted here or written the same way). */
export function selectedTextBox(
	state: EditorState,
): { pos: number; settings: TextBoxSettings } | null {
	const { selection } = state;
	if (!(selection instanceof NodeSelection) || selection.node.type !== schema.nodes.image)
		return null;
	const attrs = selection.node.attrs;
	if (!attrs.textBoxEditable) return null;
	return {
		pos: selection.from,
		settings: {
			lines: lines(attrs.textBoxText),
			widthPx: Number(attrs.widthPx),
			heightPx: Number(attrs.heightPx),
			border: attrs.textBoxBorder !== false,
		},
	};
}

const attrsFor = (settings: TextBoxSettings) => ({
	unsupported: 'Text box',
	textBoxText: JSON.stringify(settings.lines),
	textBoxEditable: true,
	textBoxBorder: settings.border,
	widthPx: settings.widthPx,
	heightPx: settings.heightPx,
});

/** Insert > Text Box: an inline text box at the selection, as one undo step. */
export function insertTextBox(view: EditorView, settings: TextBoxSettings): boolean {
	if (!view.editable) return false;
	const node = schema.nodes.image!.create(attrsFor(settings));
	view.dispatch(closeHistory(view.state.tr.replaceSelectionWith(node, false)).scrollIntoView());
	return true;
}

/** Applies new text, size and outline to the selected text box, as one undo step. */
export function updateTextBox(view: EditorView, pos: number, settings: TextBoxSettings): boolean {
	const node = view.state.doc.nodeAt(pos);
	if (!view.editable || !node || node.type !== schema.nodes.image || !node.attrs.textBoxEditable)
		return false;
	const tr = view.state.tr.setNodeMarkup(pos, undefined, { ...node.attrs, ...attrsFor(settings) });
	view.dispatch(closeHistory(tr.setSelection(NodeSelection.create(tr.doc, pos))));
	return true;
}
