import { createMultilingualControls } from './multilingual-ribbon';
import { createReviewControls } from './review-ribbon';
import { buildParagraphSpacing } from './ribbon-home';
import { SYMBOLS } from './insert-text-commands';
import { PAGE_SIZE_OPTIONS } from './page-size';

const TOC_OPTIONS: Array<[string, string]> = [
	['3', 'Table of contents, levels 1 to 3'],
	['2', 'Table of contents, levels 1 to 2'],
	['5', 'Table of contents, levels 1 to 5'],
	['remove', 'Remove table of contents'],
];
const ADD_TEXT_OPTIONS: Array<[string, string]> = [
	['1', 'Level 1'],
	['2', 'Level 2'],
	['3', 'Level 3'],
	['0', 'Do not show in table of contents'],
];
const PAGE_NUMBER_OPTIONS: Array<[string, string]> = [
	['top:left', 'Top left'],
	['top:center', 'Top center'],
	['top:right', 'Top right'],
	['bottom:left', 'Bottom left'],
	['bottom:center', 'Bottom center'],
	['bottom:right', 'Bottom right'],
];
import { menuGallery } from './ribbon-gallery-menu';
import { colorSplit, group, launcher, menuSelect, select, stack, tool } from './ribbon-parts';

/** The Insert, Layout, References, Review and View tabs. */
export function buildOtherPanels(panels: Map<string, HTMLElement>): void {
	const languageControls = createMultilingualControls();
	languageControls.setAttribute('role', 'group');
	languageControls.setAttribute('aria-label', 'Language');
	languageControls.dataset.label = 'Language';
	panels
		.get('Review')!
		.append(
			group(
				'Proofing',
				tool('Spelling', 'spelling', { type: 'spelling' }, { large: true }),
				tool('Word count', 'wordCount', { type: 'wordCount' }, { large: true }),
				tool('Read aloud', 'readAloud', { type: 'readAloud' }, { large: true }),
			),
			...createReviewControls(),
			languageControls,
		);
	panels.get('Insert')!.append(
		group(
			'Tables',
			tool('Insert table', 'table', { type: 'tablePicker' }, { large: true, caption: 'Table' }),
		),
		group(
			'Illustrations',
			tool(
				'Insert picture',
				'picture',
				{ type: 'insertPicture' },
				{ large: true, caption: 'Pictures' },
			),
			tool(
				'Format picture',
				'formatPicture',
				{ type: 'formatPicture' },
				{ large: true, caption: 'Format' },
			),
		),
		group(
			'Links',
			tool('Insert link', 'link', { type: 'link' }, { large: true, caption: 'Link' }),
			tool('Bookmark', 'bookmark', { type: 'formatDialog', kind: 'bookmark' }, { large: true }),
		),
		group(
			'Breaks',
			tool('Cover page', 'coverPage', { type: 'coverPage' }, { large: true }),
			tool('Blank page', 'blankPage', { type: 'blankPage' }, { large: true }),
			tool(
				'Insert page break',
				'pageBreak',
				{ type: 'insertBreak', kind: 'page' },
				{ large: true, caption: 'Page break' },
			),
			tool(
				'Insert column break',
				'columnBreak',
				{ type: 'insertBreak', kind: 'column' },
				{ large: true, caption: 'Column break' },
			),
		),
		group(
			'Header & Footer',
			menuSelect(
				'Header',
				'header',
				[['blank', 'Blank']],
				() => ({ type: 'headerFooter', kind: 'header' }),
				{
					momentary: true,
				},
			),
			menuSelect(
				'Footer',
				'footer',
				[['blank', 'Blank']],
				() => ({ type: 'headerFooter', kind: 'footer' }),
				{
					momentary: true,
				},
			),
			menuSelect(
				'Page number',
				'pageNumberIcon',
				PAGE_NUMBER_OPTIONS,
				(value) => {
					const [position, align] = value.split(':') as [
						'top' | 'bottom',
						'left' | 'center' | 'right',
					];
					return { type: 'pageNumber', position, align };
				},
				{ momentary: true },
			),
		),
		group(
			'Text',
			menuSelect(
				'Drop cap',
				'dropCap',
				[
					['none', 'None'],
					['drop', 'Dropped'],
					['margin', 'In margin'],
				],
				(value) => ({ type: 'dropCap', value: value as 'none' }),
				{ momentary: true },
			),
			menuSelect(
				'Date and time',
				'dateTime',
				[
					['long', 'long'],
					['short', 'short'],
					['time', 'time'],
					['datetime', 'datetime'],
				],
				(value) => ({ type: 'insertDateTime', value: value as 'long' }),
				{ momentary: true },
			),
		),
		group(
			'Symbols',
			menuSelect(
				'Symbol',
				'symbol',
				SYMBOLS.map((symbol) => [symbol, symbol]),
				(value) => ({ type: 'insertSymbol', value }),
				{ momentary: true },
			),
		),
	);
	panels.get('References')!.append(
		group(
			'Table of Contents',
			menuSelect(
				'Insert table of contents',
				'toc',
				TOC_OPTIONS,
				(value) =>
					value === 'remove'
						? { type: 'toc', key: 'remove' }
						: { type: 'toc', key: 'insert', levels: Number(value) },
				{ momentary: true },
			),
			menuSelect(
				'Add text',
				'addText',
				ADD_TEXT_OPTIONS,
				(value) => ({ type: 'addText', level: Number(value) as 0 | 1 | 2 | 3 }),
				{ momentary: true },
			),
			tool(
				'Update table of contents',
				'updateTable',
				{ type: 'toc', key: 'update' },
				{ large: true, caption: 'Update Table' },
			),
		),
		group(
			'Captions',
			tool('Insert caption', 'caption', { type: 'formatDialog', kind: 'caption' }, { large: true }),
			tool(
				'Cross-reference',
				'crossReference',
				{ type: 'formatDialog', kind: 'crossReference' },
				{ large: true },
			),
		),
		group(
			'Footnotes',
			tool(
				'Insert footnote',
				'footnote',
				{ type: 'insertNote', kind: 'footnote' },
				{ large: true, caption: 'Footnote' },
			),
			tool(
				'Insert endnote',
				'endnote',
				{ type: 'insertNote', kind: 'endnote' },
				{ large: true, caption: 'Endnote' },
			),
		),
	);
	panels.get('Layout')!.append(
		group(
			'Page setup',
			menuGallery(
				'Margins',
				'margins',
				'margins',
				[
					['normal', 'Normal'],
					['narrow', 'Narrow'],
					['moderate', 'Moderate'],
					['wide', 'Wide'],
				],
				(value) => ({ type: 'page', key: 'margin', value }),
				{
					commands: [
						{ label: 'Custom Margins…', action: { type: 'formatDialog', kind: 'pageSetup' } },
					],
				},
			),
			menuGallery('Page size', 'pageSize', 'size', PAGE_SIZE_OPTIONS, (value) => ({
				type: 'page',
				key: 'size',
				value,
			})),
			menuGallery(
				'Orientation',
				'orientation',
				'orientation',
				[
					['portrait', 'Portrait'],
					['landscape', 'Landscape'],
				],
				(value) => ({ type: 'page', key: 'orientation', value }),
			),
			menuGallery(
				'Vertical alignment',
				'verticalAlign',
				'verticalAlign',
				[
					['top', 'Top'],
					['center', 'Center'],
					['both', 'Justified'],
					['bottom', 'Bottom'],
				],
				(value) => ({ type: 'page', key: 'verticalAlign', value }),
			),
			menuSelect(
				'Hyphenation',
				'hyphenation',
				[
					['none', 'None'],
					['auto', 'Automatic'],
				],
				(value) => ({ type: 'hyphenation', value }),
			),
			menuSelect(
				'Line numbers',
				'lineNumbers',
				[
					['none', 'None'],
					['continuous', 'Continuous'],
					['newPage', 'Restart Each Page'],
					['newSection', 'Restart Each Section'],
				],
				(value) => ({ type: 'page', key: 'lineNumbers', value }),
			),
			menuGallery(
				'Columns',
				'columns',
				'columns',
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
			menuSelect(
				'Page number format',
				'pageNumber',
				[
					['decimal', '1, 2, 3'],
					['lowerRoman', 'i, ii, iii'],
					['upperRoman', 'I, II, III'],
					['lowerLetter', 'a, b, c'],
					['upperLetter', 'A, B, C'],
				],
				(value) => ({ type: 'page', key: 'numberFormat', value }),
			),
			menuSelect(
				'Page numbering',
				'numbering2',
				[
					['continue', 'Continue from previous section'],
					['restart', 'Start at 1'],
				],
				(value) => ({ type: 'page', key: 'numberStart', value }),
			),
			stack(
				tool(
					'Different first page',
					'firstPage',
					{ type: 'page', key: 'titlePage', value: 'toggle' },
					{ inline: true },
				),
				tool(
					'Different odd and even pages',
					'oddEven',
					{ type: 'evenOddHeaders' },
					{ inline: true, caption: 'Different odd & even pages' },
				),
			),
		),
		group(
			'Section breaks',
			tool(
				'Next page section break',
				'sectionNext',
				{ type: 'sectionBreak', kind: 'nextPage' },
				{ inline: true, caption: 'Next page' },
			),
			tool(
				'Continuous section break',
				'sectionContinuous',
				{ type: 'sectionBreak', kind: 'continuous' },
				{ inline: true, caption: 'Continuous' },
			),
			tool(
				'Even page section break',
				'sectionEven',
				{ type: 'sectionBreak', kind: 'evenPage' },
				{ inline: true, caption: 'Even page' },
			),
			tool(
				'Odd page section break',
				'sectionOdd',
				{ type: 'sectionBreak', kind: 'oddPage' },
				{ inline: true, caption: 'Odd page' },
			),
		),
		group(
			'Page background',
			colorSplit(
				'Page color',
				'pageColor',
				[],
				(value) => ({ type: 'pageColor', value }),
				'#e2efd9',
				{ noneLabel: 'No Color' },
			),
		),
		...buildParagraphSpacing(),
	);
	panels.get('View')!.append(
		group(
			'Show',
			tool(
				'Show hidden text',
				'hidden',
				{ type: 'showHidden' },
				{ large: true, caption: 'Hidden text' },
			),
			tool('Ruler', 'ruler', { type: 'ruler' }, { large: true }),
			tool('Gridlines', 'gridlines', { type: 'gridlines' }, { large: true }),
			tool(
				'Page thumbnails',
				'thumbnails',
				{ type: 'thumbnails' },
				{ large: true, caption: 'Thumbnails' },
			),
		),
		group(
			'Zoom',
			tool(
				'Zoom dialog',
				'zoom',
				{ type: 'formatDialog', kind: 'zoom' },
				{ large: true, caption: 'Zoom' },
			),
			menuSelect(
				'Zoom',
				'zoom',
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
			tool(
				'Zoom to 100%',
				'zoom100',
				{ type: 'zoomFit', mode: 'actual' },
				{ large: true, caption: '100%' },
			),
			tool('One page', 'onePage', { type: 'zoomFit', mode: 'page' }, { large: true }),
			tool('Multiple pages', 'twoPages', { type: 'zoomFit', mode: 'pages' }, { large: true }),
			tool('Page width', 'pageWidth', { type: 'zoomFit', mode: 'width' }, { large: true }),
		),
		group(
			'Layout view',
			menuSelect(
				'Layout view',
				'view',
				[
					['draft', 'Draft'],
					['print', 'Print Layout'],
				],
				(value) => ({ type: 'view', value: value === 'print' ? 'print' : 'draft' }),
			),
			tool('Print', 'print', { type: 'print' }, { large: true }),
		),
	);
	panels
		.get('Layout')!
		.querySelector('[data-label="Page setup"]')
		?.append(launcher('Page setup settings', { type: 'formatDialog', kind: 'pageSetup' }));
}
