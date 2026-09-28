import type { EditorCore } from './editor-core';
import type { RibbonAction } from './ribbon';
import { runRibbonCommand } from './editor-commands';
import { runListAction } from './list-commands';
import { focusView } from './focus-view';

/** Routes a ribbon action to the controller that owns it. */
export function routeRibbonAction(core: EditorCore, action: RibbonAction): void {
	const { inserts, pages, parts, shell } = core;
	if (inserts.handle(action)) return;
	const target = core.targetView();
	if (action.type === 'search') shell.searchPanel?.open();
	else if (action.type === 'thumbnails')
		core.element.toggleAttribute('show-thumbnails', !core.viewOptions.showThumbnails);
	else if (action.type === 'zoom') pages.setZoom(action.value);
	else if (action.type === 'list' && target) {
		runListAction(target, action.key, core.model);
		focusView(target);
	} else if (action.type === 'view') pages.setViewMode(action.value);
	else if (action.type === 'print') pages.print();
	else if (action.type === 'insertNote') parts.insertNote(action.kind, shell.canvas, shell.paper);
	else if (action.type === 'page') pages.changePageSetup(action.key, action.value);
	else if (action.type === 'sectionBreak') pages.insertSectionBreak(action.kind);
	else if (action.type === 'evenOddHeaders') pages.toggleEvenOddHeaders();
	else if (action.type === 'reviewDisplay') {
		core.reviewDisplayMode = action.value;
		core.view?.dispatch(core.view.state.tr);
	} else if (action.type === 'review') shell.review?.handleReview(action.key);
	else if (action.type === 'comments') shell.review?.handleComments(action.key);
	else if (target) {
		runRibbonCommand(target, action, core.collab.ids);
		focusView(target);
	}
}
