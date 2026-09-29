import { closeHistory } from 'prosemirror-history';
import type { EditorView } from 'prosemirror-view';
import { schema } from './schema';

export interface CoverPageText {
	title: string;
	subtitle: string;
	author: string;
	date: string;
}

/**
 * Insert > Cover Page: a plain centred title page at the top of the document (title, subtitle,
 * author, date) that pushes the existing content onto the next page. It is one simple design,
 * not Word's gallery of building blocks. `titleStyle` is the document's Title style id, if it has
 * one. Returns false for a read-only view.
 */
export function insertCoverPage(
	view: EditorView,
	text: CoverPageText,
	titleStyle?: string,
): boolean {
	if (!view.editable) return false;
	const { doc } = view.state;
	const line = (value: string, attrs: Record<string, unknown> = {}) =>
		schema.nodes.paragraph!.create(
			{ align: 'center', spacingAfterTwips: 240, ...attrs },
			value ? schema.text(value) : undefined,
		);
	const cover = [
		line('', { spacingBeforeTwips: 2880 }),
		line(text.title, { ...(titleStyle ? { style: titleStyle } : {}), spacingAfterTwips: 480 }),
		line(text.subtitle),
		line(text.author, { spacingBeforeTwips: 1440 }),
		line(text.date),
	];
	let tr = view.state.tr.insert(0, cover);
	const first = doc.firstChild;
	if (first?.type.name === 'paragraph') {
		const pos = cover.reduce((size, node) => size + node.nodeSize, 0);
		tr = tr.setNodeAttribute(pos, 'pageBreakBefore', true);
	}
	view.dispatch(closeHistory(tr).scrollIntoView());
	return true;
}
