/**
 * Stable, locale-independent ids of every ribbon control. `hiddenActions` is keyed on these ids;
 * the English labels below are only the keys the localizer uses to find a control's display text.
 * `element-api.test.ts` fails when a control is added to or removed from the ribbon without
 * updating this list.
 */
export const RIBBON_ACTION_IDS = [
	'font-family',
	'font-size',
	'bold',
	'italic',
	'underline',
	'strikethrough',
	'superscript',
	'subscript',
	'font-color',
	'text-highlight',
	'clear-formatting',
	'decrease-indent',
	'increase-indent',
	'spacing-after',
	'spacing-before',
	'line-spacing',
	'bulleted-list',
	'numbered-list',
	'decrease-list-level',
	'increase-list-level',
	'remove-list',
	'align-left',
	'align-center',
	'align-right',
	'justify',
	'find-and-replace',
	'insert-table',
	'insert-picture',
	'format-picture',
	'insert-link',
	'insert-page-break',
	'insert-column-break',
	'margins',
	'orientation',
	'vertical-alignment',
	'columns',
	'page-number-format',
	'page-numbering',
	'different-first-page',
	'different-odd-and-even-pages',
	'next-page-section-break',
	'continuous-section-break',
	'even-page-section-break',
	'odd-page-section-break',
	'insert-table-of-contents',
	'update-table-of-contents',
	'insert-footnote',
	'insert-endnote',
	'paragraph-direction',
	'text-language',
	'custom-text-language-tag',
	'east-asian-language',
	'custom-east-asian-language-tag',
	'complex-script-language',
	'custom-complex-script-language-tag',
	'run-direction',
	'track-changes',
	'display-for-review',
	'previous-change',
	'next-change',
	'accept',
	'reject',
	'accept-all',
	'reject-all',
	'add-comment',
	'comments',
	'show-hidden-text',
	'page-thumbnails',
	'zoom',
	'layout-view',
	'print',
	'insert-row-above',
	'insert-row-below',
	'delete-row',
	'insert-column-left',
	'insert-column-right',
	'delete-column',
	'delete-table',
] as const;

export type RibbonActionId = (typeof RIBBON_ACTION_IDS)[number];

/** English control label of each action, the lookup key into the locale strings. Display-only. */
export const RIBBON_ACTION_LABELS: Readonly<Record<RibbonActionId, string>> = {
	'font-family': 'Font family',
	'font-size': 'Font size',
	bold: 'Bold',
	italic: 'Italic',
	underline: 'Underline',
	strikethrough: 'Strikethrough',
	superscript: 'Superscript',
	subscript: 'Subscript',
	'font-color': 'Font color',
	'text-highlight': 'Text highlight',
	'clear-formatting': 'Clear formatting',
	'decrease-indent': 'Decrease indent',
	'increase-indent': 'Increase indent',
	'spacing-after': 'Spacing after',
	'spacing-before': 'Spacing before',
	'line-spacing': 'Line spacing',
	'bulleted-list': 'Bulleted list',
	'numbered-list': 'Numbered list',
	'decrease-list-level': 'Decrease list level',
	'increase-list-level': 'Increase list level',
	'remove-list': 'Remove list',
	'align-left': 'Align left',
	'align-center': 'Align center',
	'align-right': 'Align right',
	justify: 'Justify',
	'find-and-replace': 'Find and replace',
	'insert-table': 'Insert table',
	'insert-picture': 'Insert picture',
	'format-picture': 'Format picture',
	'insert-link': 'Insert link',
	'insert-page-break': 'Insert page break',
	'insert-column-break': 'Insert column break',
	margins: 'Margins',
	orientation: 'Orientation',
	'vertical-alignment': 'Vertical alignment',
	columns: 'Columns',
	'page-number-format': 'Page number format',
	'page-numbering': 'Page numbering',
	'different-first-page': 'Different first page',
	'different-odd-and-even-pages': 'Different odd and even pages',
	'next-page-section-break': 'Next page section break',
	'continuous-section-break': 'Continuous section break',
	'even-page-section-break': 'Even page section break',
	'odd-page-section-break': 'Odd page section break',
	'insert-table-of-contents': 'Insert table of contents',
	'update-table-of-contents': 'Update table of contents',
	'insert-footnote': 'Insert footnote',
	'insert-endnote': 'Insert endnote',
	'paragraph-direction': 'Paragraph direction',
	'text-language': 'Text language',
	'custom-text-language-tag': 'Custom text language tag',
	'east-asian-language': 'East Asian language',
	'custom-east-asian-language-tag': 'Custom east asian language tag',
	'complex-script-language': 'Complex script language',
	'custom-complex-script-language-tag': 'Custom complex script language tag',
	'run-direction': 'Run direction',
	'track-changes': 'Track changes',
	'display-for-review': 'Display for review',
	'previous-change': 'Previous change',
	'next-change': 'Next change',
	accept: 'Accept',
	reject: 'Reject',
	'accept-all': 'Accept all',
	'reject-all': 'Reject all',
	'add-comment': 'Add comment',
	comments: 'Comments',
	'show-hidden-text': 'Show hidden text',
	'page-thumbnails': 'Page thumbnails',
	zoom: 'Zoom',
	'layout-view': 'Layout view',
	print: 'Print',
	'insert-row-above': 'Insert row above',
	'insert-row-below': 'Insert row below',
	'delete-row': 'Delete row',
	'insert-column-left': 'Insert column left',
	'insert-column-right': 'Insert column right',
	'delete-column': 'Delete column',
	'delete-table': 'Delete table',
};

/**
 * The English control labels `hiddenActions` accepted before ids existed. Deprecated: accepted for
 * one more release, mapped to ids, and reported through a `document-warning` event.
 */
export type LegacyRibbonLabel =
	| 'Font family'
	| 'Font size'
	| 'Bold'
	| 'Italic'
	| 'Underline'
	| 'Strikethrough'
	| 'Superscript'
	| 'Subscript'
	| 'Font color'
	| 'Text highlight'
	| 'Clear formatting'
	| 'Decrease indent'
	| 'Increase indent'
	| 'Spacing after'
	| 'Spacing before'
	| 'Line spacing'
	| 'Bulleted list'
	| 'Numbered list'
	| 'Decrease list level'
	| 'Increase list level'
	| 'Remove list'
	| 'Align left'
	| 'Align center'
	| 'Align right'
	| 'Justify'
	| 'Find and replace'
	| 'Insert table'
	| 'Insert picture'
	| 'Format picture'
	| 'Insert link'
	| 'Insert page break'
	| 'Insert column break'
	| 'Margins'
	| 'Orientation'
	| 'Vertical alignment'
	| 'Columns'
	| 'Page number format'
	| 'Page numbering'
	| 'Different first page'
	| 'Different odd and even pages'
	| 'Next page section break'
	| 'Continuous section break'
	| 'Even page section break'
	| 'Odd page section break'
	| 'Insert table of contents'
	| 'Update table of contents'
	| 'Insert footnote'
	| 'Insert endnote'
	| 'Paragraph direction'
	| 'Text language'
	| 'Custom text language tag'
	| 'East Asian language'
	| 'Custom east asian language tag'
	| 'Complex script language'
	| 'Custom complex script language tag'
	| 'Run direction'
	| 'Track changes'
	| 'Display for review'
	| 'Previous change'
	| 'Next change'
	| 'Accept'
	| 'Reject'
	| 'Accept all'
	| 'Reject all'
	| 'Add comment'
	| 'Comments'
	| 'Show hidden text'
	| 'Page thumbnails'
	| 'Zoom'
	| 'Layout view'
	| 'Print'
	| 'Insert row above'
	| 'Insert row below'
	| 'Delete row'
	| 'Insert column left'
	| 'Insert column right'
	| 'Delete column'
	| 'Delete table';

/** A value `hiddenActions` accepts: a stable id, or (deprecated) the old English label. */
export type RibbonActionInput = RibbonActionId | LegacyRibbonLabel;

const labelToId = new Map<string, RibbonActionId>(
	RIBBON_ACTION_IDS.map((id) => [RIBBON_ACTION_LABELS[id], id]),
);

export function isRibbonActionId(value: string): value is RibbonActionId {
	return (RIBBON_ACTION_IDS as readonly string[]).includes(value);
}

/** The id of an English control label, or null. */
export function ribbonActionIdForLabel(label: string | null | undefined): RibbonActionId | null {
	return labelToId.get(label ?? '') ?? null;
}

/**
 * Maps `hiddenActions` input to ids. Deprecated English labels are converted and listed in
 * `legacy`; unknown values are dropped, matching how the ribbon ignored them before.
 */
export function normalizeRibbonActions(input: readonly string[]): {
	ids: RibbonActionId[];
	legacy: string[];
} {
	const ids: RibbonActionId[] = [];
	const legacy: string[] = [];
	for (const value of input) {
		if (isRibbonActionId(value)) ids.push(value);
		else {
			const id = labelToId.get(value);
			if (id) {
				ids.push(id);
				legacy.push(value);
			}
		}
	}
	return { ids, legacy };
}
