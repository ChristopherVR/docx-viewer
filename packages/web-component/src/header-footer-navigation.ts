import type { DocumentModel } from 'docx-core';
import type { EditorView } from 'prosemirror-view';
import { TextSelection } from 'prosemirror-state';
import { effectiveHeaderFooter } from './header-footer-link';
import { newHeaderFooterId } from './header-footer-commands';
import type { HeaderFooterContext } from './header-footer-ribbon';

/** An empty story can be opened without adding a part until the user edits it. */
export function storyPreview(
	model: DocumentModel,
	context: HeaderFooterContext | undefined,
): DocumentModel {
	if (
		!context ||
		!model.sections ||
		effectiveHeaderFooter(model, context.index, context.kind, context.slot)
	)
		return model;
	return {
		...model,
		sections: model.sections.map((section, i) =>
			i !== context.index
				? section
				: {
						...section,
						[context.kind]: {
							...section[context.kind],
							[context.slot]: {
								blocks: [{ type: 'paragraph', id: newHeaderFooterId(), runs: [] }],
							},
						},
					},
		),
	};
}
export function selectSectionStart(view: EditorView, model: DocumentModel, index: number): void {
	const end = model.sections?.[index - 1]?.endsAtBlockId;
	const block = end ? model.blocks.findIndex((block) => block.id === end) + 1 : 0;
	let position = 0;
	for (let i = 0; i < block && i < view.state.doc.childCount; i++)
		position += view.state.doc.child(i).nodeSize;
	view.dispatch(
		view.state.tr
			.setSelection(
				TextSelection.near(
					view.state.doc.resolve(Math.min(position + 1, view.state.doc.content.size)),
				),
			)
			.scrollIntoView(),
	);
}
