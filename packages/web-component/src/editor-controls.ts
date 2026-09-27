import type { EditorView } from 'prosemirror-view';
import type { DocumentModel } from '@christophervr/docx-core';
import type { EditorLocale } from './localization';
import { formatWordCount } from './localization';
import { syncFontControls, syncParagraphControls, syncFormatControls } from './ribbon-controls';
import { syncMultilingualControls } from './multilingual-ribbon';
import { syncStylePicker } from './paragraph-styles';
import { canExecuteTableCommand } from './table-commands';
import type { RibbonAction } from './ribbon';
import { countWords } from './word-count';

export function refreshEditorControls(
	toolbar: HTMLElement | undefined,
	view: EditorView | undefined,
	model: DocumentModel,
	readOnly: boolean,
	collaboration: boolean,
	locale: EditorLocale,
	language: string,
) {
	toolbar
		?.querySelectorAll<HTMLButtonElement | HTMLSelectElement | HTMLInputElement>(
			'.ribbon-group button, .ribbon-group input, .ribbon-group select',
		)
		.forEach((control) => {
			const label = control.dataset.localearialabel ?? control.getAttribute('aria-label');
			control.disabled = readOnly && label !== 'Find and replace' && label !== 'Zoom';
		});
	if (!view) return;
	const { state } = view;
	if (toolbar) {
		syncMultilingualControls(toolbar, state);
		syncStylePicker(toolbar, view, model, locale);
		syncFormatControls(toolbar, state);
		for (const button of toolbar.querySelectorAll<HTMLButtonElement>('button[data-action]')) {
			const action = JSON.parse(button.dataset.action!) as RibbonAction;
			if (action.type === 'tableEdit')
				button.disabled =
					readOnly || Boolean(collaboration) || !canExecuteTableCommand(view, action.key);
		}
		syncFontControls(toolbar, state);
		syncParagraphControls(toolbar, state);
	}
	const content = state.doc.textBetween(0, state.doc.content.size, ' ').trim();
	const words = countWords(content, language || undefined);
	const status = toolbar?.parentElement?.querySelector('.dve-status');
	if (status) status.textContent = `Page 1 · ${formatWordCount(locale, words)}`;
}
