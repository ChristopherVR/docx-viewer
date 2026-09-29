import type { DocumentModel } from '@christophervr/docx-core';
import type { EditorView } from 'prosemirror-view';
import { resolvedAttrs } from './paragraph-format';
import { updateRuler, type RulerGeometry } from './ruler';

const px = (twips: unknown) =>
	typeof twips === 'number' && Number.isFinite(twips) ? twips / 15 : 0;

/** The ruler geometry for the paragraph holding the selection, from the document's page attributes. */
export function rulerGeometry(view: EditorView, model: DocumentModel): RulerGeometry {
	const { attrs } = view.state.doc;
	const paragraph = resolvedAttrs(view.state.selection.$from.parent.attrs, model);
	const hanging = px(paragraph.hangingTwips);
	return {
		pageWidth: Number(attrs.pageWidth) || 816,
		marginLeft: Number(attrs.marginLeft) || 0,
		marginRight: Number(attrs.marginRight) || 0,
		indentLeft: px(paragraph.indentStartTwips ?? paragraph.indentLeftTwips),
		indentRight: px(paragraph.indentEndTwips ?? paragraph.indentRightTwips),
		firstLine: hanging ? -hanging : px(paragraph.firstLineTwips),
	};
}

/** Redraws the ruler, if the View tab has it on, for the current selection and zoom. */
export function syncRuler(toolbar: HTMLElement, view: EditorView, model: DocumentModel): void {
	const frame = toolbar.parentElement;
	const ruler = frame?.querySelector<HTMLElement>('.dve-ruler');
	if (!ruler) return;
	const paper = frame?.querySelector<HTMLElement>('.dve-paper');
	const zoom = Number(paper?.style.getPropertyValue('--dve-zoom')) || 1;
	updateRuler(ruler, rulerGeometry(view, model), zoom);
}
