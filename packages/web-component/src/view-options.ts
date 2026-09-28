import type { EditorCore } from './editor-core';
import { findLocalizedControl } from './localization';
import type { RibbonActionId } from './ribbon-action-ids';
import { applyRibbonVisibility } from './ribbon-visibility';

/** Host-controlled UI customisation; each value maps to one element property and attribute. */
export class ViewOptions {
	showThumbnails = false;
	showToolbar = true;
	hiddenActions: readonly RibbonActionId[] = [];
}

/** Pushes the current options into the built shell; a no-op until the element connects. */
export function applyViewOptions(core: EditorCore): void {
	const { toolbar, navigator } = core.shell;
	const { showThumbnails, showToolbar, hiddenActions } = core.viewOptions;
	navigator?.setOpen(showThumbnails);
	if (!toolbar) return;
	applyRibbonVisibility(toolbar, showToolbar, hiddenActions);
	findLocalizedControl(toolbar, 'Page thumbnails')?.setAttribute(
		'aria-pressed',
		String(showThumbnails),
	);
}
