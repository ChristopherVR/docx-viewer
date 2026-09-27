import { createMultilingualControls } from './multilingual-ribbon';
import { createReviewControls } from './review-ribbon';
import { button, group, select } from './ribbon-parts';

/** The Review, Insert, Layout and View tabs. */
export function buildOtherPanels(panels: Map<string, HTMLElement>): void {
	const languageControls = createMultilingualControls();
	languageControls.setAttribute('role', 'group');
	languageControls.setAttribute('aria-label', 'Language');
	languageControls.dataset.label = 'Language';
	panels.get('Review')!.append(languageControls, ...createReviewControls());
	panels
		.get('Insert')!
		.append(
			group('Tables', button('Insert table', '▦ Table', { type: 'table' })),
			group(
				'Illustrations',
				button('Insert picture', 'Pictures', { type: 'insertPicture' }),
				button('Format picture', 'Format', { type: 'formatPicture' }),
			),
			group('Links', button('Insert link', 'Link', { type: 'link' })),
			group(
				'Footnotes',
				button('Insert footnote', 'Footnote', { type: 'insertNote', kind: 'footnote' }),
				button('Insert endnote', 'Endnote', { type: 'insertNote', kind: 'endnote' }),
			),
			group(
				'Breaks',
				button('Insert page break', 'Page break', { type: 'insertBreak', kind: 'page' }),
				button('Insert column break', 'Column break', { type: 'insertBreak', kind: 'column' }),
			),
		);
	panels.get('Layout')!.append(
		group(
			'Page setup',
			select(
				'Margins',
				[
					['normal', 'Normal'],
					['narrow', 'Narrow'],
					['wide', 'Wide'],
				],
				(value) => ({ type: 'page', key: 'margin', value }),
			),
			select(
				'Orientation',
				[
					['portrait', 'Portrait'],
					['landscape', 'Landscape'],
				],
				(value) => ({ type: 'page', key: 'orientation', value }),
			),
			select(
				'Vertical alignment',
				[
					['top', 'Top'],
					['center', 'Center'],
					['both', 'Justified'],
					['bottom', 'Bottom'],
				],
				(value) => ({ type: 'page', key: 'verticalAlign', value }),
			),
			select(
				'Columns',
				[
					['1', 'One column'],
					['2', 'Two columns'],
					['3', 'Three columns'],
				],
				(value) => ({ type: 'page', key: 'columns', value }),
			),
		),
		group(
			'Page numbers',
			select(
				'Page number format',
				[
					['decimal', '1, 2, 3'],
					['lowerRoman', 'i, ii, iii'],
					['upperRoman', 'I, II, III'],
					['lowerLetter', 'a, b, c'],
					['upperLetter', 'A, B, C'],
				],
				(value) => ({ type: 'page', key: 'numberFormat', value }),
			),
			select(
				'Page numbering',
				[
					['continue', 'Continue from previous section'],
					['restart', 'Start at 1'],
				],
				(value) => ({ type: 'page', key: 'numberStart', value }),
			),
			button('Different first page', 'Different first page', {
				type: 'page',
				key: 'titlePage',
				value: 'toggle',
			}),
			button('Different odd and even pages', 'Different odd & even pages', {
				type: 'evenOddHeaders',
			}),
		),
		group(
			'Section breaks',
			button('Next page section break', 'Next page', { type: 'sectionBreak', kind: 'nextPage' }),
			button('Continuous section break', 'Continuous', {
				type: 'sectionBreak',
				kind: 'continuous',
			}),
			button('Even page section break', 'Even page', { type: 'sectionBreak', kind: 'evenPage' }),
			button('Odd page section break', 'Odd page', { type: 'sectionBreak', kind: 'oddPage' }),
		),
	);
	panels.get('View')!.append(
		group('Show', button('Show hidden text', 'Hidden text', { type: 'showHidden' })),
		group(
			'Zoom',
			select(
				'Zoom',
				[
					['50', '50%'],
					['75', '75%'],
					['90', '90%'],
					['100', '100%'],
					['125', '125%'],
					['150', '150%'],
				],
				(value) => ({ type: 'zoom', value: Number(value) }),
			),
		),
		group(
			'Layout view',
			select(
				'Layout view',
				[
					['draft', 'Draft'],
					['print', 'Print Layout'],
				],
				(value) => ({ type: 'view', value: value === 'print' ? 'print' : 'draft' }),
			),
			button('Print', 'Print', { type: 'print' }),
		),
	);
}
