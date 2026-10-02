import { headingLevel, type DocumentModel, type Paragraph } from 'docx-core';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import { NodeSelection, TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';

export type GoToKind = 'heading' | 'bookmark' | 'table' | 'graphic';

export interface GoToTarget {
	label: string;
	/** Document position to select: a text position, or a node position for a picture. */
	pos: number;
	/** Set when the target is a picture (an inline atom that gets a node selection). */
	node?: boolean;
}

const clip = (text: string) => {
	const value = text.replace(/\s+/g, ' ').trim();
	return value.length > 60 ? `${value.slice(0, 57)}…` : value;
};

/**
 * Everything Word's Go To can jump to that this editor models: headings (by style), bookmarks,
 * tables and pictures, in document order. Bookmarks are listed once, at their paragraph.
 */
export function goToTargets(
	doc: ProseMirrorNode,
	model: DocumentModel,
	kind: GoToKind,
): GoToTarget[] {
	const targets: GoToTarget[] = [];
	let tables = 0;
	let pictures = 0;
	doc.descendants((node, pos) => {
		if (kind === 'heading' && node.type.name === 'paragraph') {
			const level = headingLevel(
				{ style: node.attrs.style || undefined } as Paragraph,
				model.paragraphStyles,
			);
			if (level && node.textContent.trim())
				targets.push({ label: `${'  '.repeat(level - 1)}${clip(node.textContent)}`, pos: pos + 1 });
		} else if (kind === 'bookmark' && node.type.name === 'paragraph') {
			for (const name of (node.attrs.bookmarks as string[] | undefined) ?? [])
				if (!name.startsWith('_')) targets.push({ label: name, pos: pos + 1 });
		} else if (kind === 'table' && node.type.name === 'table') {
			tables++;
			targets.push({ label: `${tables}: ${clip(node.textContent) || '…'}`, pos: pos + 3 });
			return false;
		} else if (kind === 'graphic' && node.type.name === 'image') {
			pictures++;
			targets.push({
				label: node.attrs.altText
					? `${pictures}: ${clip(String(node.attrs.altText))}`
					: String(pictures),
				pos,
				node: true,
			});
		}
		return true;
	});
	return targets;
}

/** Selects `target` and scrolls it into view. Returns false when the position no longer exists. */
export function goToTarget(view: EditorView, target: GoToTarget): boolean {
	const { doc } = view.state;
	if (target.pos < 0 || target.pos > doc.content.size) return false;
	const selection = target.node
		? doc.nodeAt(target.pos)?.type.name === 'image'
			? NodeSelection.create(doc, target.pos)
			: undefined
		: TextSelection.near(doc.resolve(target.pos));
	if (!selection) return false;
	view.dispatch(view.state.tr.setSelection(selection).scrollIntoView());
	view.focus();
	return true;
}

/**
 * Word's Select > Select Objects: selects the next picture after the selection, wrapping to the
 * first one. Returns false when the document has no pictures.
 */
export function selectNextObject(view: EditorView): boolean {
	const pictures: number[] = [];
	view.state.doc.descendants((node, pos) => {
		if (node.type.name === 'image') pictures.push(pos);
	});
	if (!pictures.length) return false;
	const after = view.state.selection.to;
	const next = pictures.find((pos) => pos >= after) ?? pictures[0]!;
	view.dispatch(
		view.state.tr.setSelection(NodeSelection.create(view.state.doc, next)).scrollIntoView(),
	);
	view.focus();
	return true;
}
