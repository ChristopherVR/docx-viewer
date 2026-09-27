import type { EditorView } from 'prosemirror-view';
import { NodeSelection } from 'prosemirror-state';
import type { DocumentModel } from '@christophervr/docx-core';
import type { EditorLocale } from './localization';
import { formatPageStatus, formatWordCount } from './localization';
import { syncFontControls, syncParagraphControls, syncFormatControls } from './ribbon-controls';
import { syncMultilingualControls } from './multilingual-ribbon';
import { syncStylePicker } from './paragraph-styles';
import { syncCharacterStylePicker } from './character-style-picker';
import { canExecuteTableCommand } from './table-commands';
import type { RibbonAction } from './ribbon';
import { countWords } from './word-count';
import { hasAnyChange, hasChangeAtCursor } from './review-commands';

export function refreshEditorControls(
	toolbar: HTMLElement | undefined,
	view: EditorView | undefined,
	model: DocumentModel,
	readOnly: boolean,
	collaboration: boolean,
	locale: EditorLocale,
	language: string,
	printPageStatus?: { current: number; total: number } | null,
	commentsOpen = false,
): { pageText: string; wordText: string } | undefined {
	toolbar
		?.querySelectorAll<HTMLButtonElement | HTMLSelectElement | HTMLInputElement>(
			'.ribbon-group button, .ribbon-group input, .ribbon-group select',
		)
		.forEach((control) => {
			const label = control.dataset.localearialabel ?? control.getAttribute('aria-label');
			control.disabled =
				readOnly &&
				!['Find and replace', 'Zoom', 'Show hidden text', 'Layout view', 'Print'].includes(
					label ?? '',
				);
		});
	if (!view) return undefined;
	const { state } = view;
	if (toolbar) {
		syncMultilingualControls(toolbar, state);
		syncStylePicker(toolbar, view, model, locale);
		syncCharacterStylePicker(toolbar, view, model, locale);
		syncFormatControls(toolbar, state);
		for (const button of toolbar.querySelectorAll<HTMLButtonElement>('button[data-action]')) {
			const action = JSON.parse(button.dataset.action!) as RibbonAction;
			if (action.type === 'tableEdit')
				button.disabled =
					readOnly || Boolean(collaboration) || !canExecuteTableCommand(view, action.key);
		}
		syncFontControls(toolbar, state);
		syncParagraphControls(toolbar, state);
		const trackButton = toolbar.querySelector<HTMLButtonElement>('[aria-label="Track changes"]');
		trackButton?.setAttribute('aria-pressed', String(Boolean(model.trackChanges)));
		const commentsButton = toolbar.querySelector<HTMLButtonElement>('[aria-label="Comments"]');
		commentsButton?.setAttribute('aria-pressed', String(commentsOpen));
		const changeAtCursor = hasChangeAtCursor(view);
		const anyChange = hasAnyChange(view);
		for (const [label, enabled] of [
			['Accept', changeAtCursor],
			['Reject', changeAtCursor],
			['Accept all', anyChange],
			['Reject all', anyChange],
			['Previous change', anyChange],
			['Next change', anyChange],
			['Add comment', !state.selection.empty],
			[
				'Format picture',
				state.selection instanceof NodeSelection && state.selection.node.type.name === 'image',
			],
		] as const) {
			const control = toolbar.querySelector<HTMLButtonElement>(
				`[data-localearialabel="${label}"], [aria-label="${label}"]`,
			);
			if (control) control.disabled = readOnly || !enabled;
		}
	}
	const content = state.doc.textBetween(0, state.doc.content.size, ' ').trim();
	const words = countWords(content, language || undefined);
	const pageText = printPageStatus
		? formatPageStatus(locale, printPageStatus.current, printPageStatus.total)
		: formatPageStatus(locale, 1, 1);
	return { pageText, wordText: formatWordCount(locale, words) };
}
