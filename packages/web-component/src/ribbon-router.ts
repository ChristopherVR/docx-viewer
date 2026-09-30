import type { EditorCore } from './editor-core';
import type { RibbonAction } from './ribbon';
import { runRibbonCommand } from './editor-commands';
import { runListAction } from './list-commands';
import { focusView } from './focus-view';
import { copyOrCut, pasteText } from './context-menu-actions';
import { translate } from './localization';
import { emit } from './events';
import { insertCoverPage } from './cover-page';
import { openTablePicker } from './table-picker';
import { openGoToPanel } from './go-to-panel';
import { selectNextObject } from './go-to';
import { closeRibbonPopover } from './ribbon-popover';
import { setHeadingLevel } from './heading-commands';
import { readAloudAvailable, textToRead, toggleReadAloud } from './read-aloud';
import { showWordCount } from './word-count-panel';
import { syncSpellingButton } from './spelling';
import { selectAll } from 'prosemirror-commands';
import { toggleFormatPainter } from './format-painter';

/** Routes a ribbon action to the controller that owns it. */
export function routeRibbonAction(core: EditorCore, action: RibbonAction): void {
	const { inserts, pages, parts, shell } = core;
	if (inserts.handle(action)) return;
	const target = core.targetView();
	if (action.type === 'clipboard') {
		if (!target || (action.key !== 'copy' && !target.editable)) return;
		const done = action.key === 'paste' ? pasteText(target) : copyOrCut(target, action.key);
		void done.then((ok) => {
			if (!ok)
				emit(core.element, 'document-warning', translate(core.locale, 'menu.clipboardDenied'));
		});
	} else if (action.type === 'search') shell.searchPanel?.open(action.focus);
	else if (action.type === 'selectAll' && target) {
		selectAll(target.state, target.dispatch);
		focusView(target);
	} else if (action.type === 'formatPainter' && target)
		toggleFormatPainter(target, (on) =>
			shell.toolbar
				?.querySelector('[aria-label="Format painter"], [data-localearialabel="Format painter"]')
				?.setAttribute('aria-pressed', String(on)),
		);
	else if (action.type === 'formatDialog') core.formatDialogs.open(action.kind);
	else if (action.type === 'pageNumber' || action.type === 'headerFooter') {
		const done =
			action.type === 'pageNumber'
				? pages.insertPageNumber(action.position, action.align)
				: pages.insertHeaderFooter(action.kind);
		if (done) parts.render(shell.canvas, shell.paper);
	} else if (action.type === 'coverPage' && target) {
		const title = Object.values(core.model.paragraphStyles?.styles ?? {}).find(
			(style) => style.id === 'Title' || style.name === 'Title',
		);
		insertCoverPage(
			target,
			{
				title: translate(core.locale, 'Document title'),
				subtitle: translate(core.locale, 'Document subtitle'),
				author: core.reviewAuthor,
				date: new Intl.DateTimeFormat(core.locale, { dateStyle: 'long' }).format(new Date()),
			},
			title?.id,
		);
		focusView(target);
	} else if (action.type === 'tablePicker' && target) {
		const anchor = shell.toolbar?.querySelector<HTMLElement>(
			'[aria-label="Insert table"], [data-localearialabel="Insert table"]',
		);
		if (anchor)
			openTablePicker(
				anchor,
				(rows, columns) => {
					runRibbonCommand(target, { type: 'table', rows, columns }, core.collab.ids);
					focusView(target);
				},
				closeRibbonPopover,
			);
	} else if (action.type === 'goTo' && target) {
		const anchor = shell.toolbar?.querySelector<HTMLElement>(
			'[aria-label="Find options"], [data-localearialabel="Find options"]',
		);
		if (anchor) openGoToPanel(anchor, target, core.model, closeRibbonPopover);
	} else if (action.type === 'selectObjects' && target) {
		selectNextObject(target);
	} else if (action.type === 'addText' && target) {
		if (!setHeadingLevel(target, core.model, action.level))
			emit(core.element, 'document-warning', 'This document has no heading styles to apply.');
		focusView(target);
	} else if (action.type === 'readAloud' && target) {
		const button = () =>
			shell.toolbar?.querySelector(
				'[aria-label="Read aloud"], [data-localearialabel="Read aloud"]',
			);
		if (!readAloudAvailable())
			emit(core.element, 'document-warning', 'This browser cannot read text aloud.');
		else
			toggleReadAloud(textToRead(target), (on) =>
				button()?.setAttribute('aria-pressed', String(on)),
			);
	} else if (action.type === 'zoomFit') pages.zoomTo(action.mode);
	else if (action.type === 'wordCount') {
		const anchor = shell.toolbar?.querySelector<HTMLElement>(
			'[aria-label="Word count"], [data-localearialabel="Word count"]',
		);
		if (anchor && target) showWordCount(anchor, target, core.locale);
	} else if (action.type === 'spelling' && target) {
		target.dom.spellcheck = !target.dom.spellcheck;
		syncSpellingButton(shell.toolbar, target);
	} else if (action.type === 'thumbnails')
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
	else if (action.type === 'pageColor') pages.setPageColor(action.value);
	else if (action.type === 'hyphenation') pages.setHyphenation(action.value);
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
