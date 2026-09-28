import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import type { EditorCore } from './editor-core';
import { modelToDoc } from './model-adapter';
import { bodyPlugins } from './editor-plugins';
import { resetStylePicker } from './paragraph-styles';
import { imageNodeView } from './image-media';

/** Rebuilds the ProseMirror view for the current model (or the detached state kept while unmounted). */
export function renderDocument(core: EditorCore): void {
	const { paper, toolbar, canvas } = core.shell;
	if (!paper) return;
	resetStylePicker(toolbar);
	core.view?.destroy();
	paper.replaceChildren();
	core.pages.refreshPageStyles();
	const state =
		core.detachedState ??
		EditorState.create({
			doc: modelToDoc(core.model),
			plugins: bodyPlugins({
				model: () => core.model,
				reviewAuthor: () => core.reviewAuthor,
				reviewDisplayMode: () => core.reviewDisplayMode,
				insertNote: (kind) => core.parts.insertNote(kind, canvas, paper),
				showSearch: () => core.shell.searchPanel?.open(),
				extraPlugins: core.inserts.plugins(),
				collaborationPlugins: [
					...(core.collab.client ? [core.collab.client.plugin] : []),
					...(core.collab.presence ? [core.collab.presence.client.plugin] : []),
				],
			}),
		});
	core.view = new EditorView(paper, {
		state,
		editable: () => !core.readOnly,
		dispatchTransaction: (transaction) => core.applyTransaction(transaction),
		nodeViews: {
			image: imageNodeView(core.imageMedia, {
				editPicture: (pos) => core.inserts.pictureDialog.open(pos),
				maxWidth: () => core.contentWidth(),
			}),
		},
		handleClick: (view, pos, event) => core.inserts.handleClick(view, pos, event),
	});
	core.detachedState = undefined;
	core.inserts.syncPaper();
	core.parts.render(canvas, paper);
	core.pages.relayout();
	core.refreshControls();
	core.scheduleCollaborationSend();
}
