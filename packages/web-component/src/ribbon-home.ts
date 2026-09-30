import { lineSpacingOptions } from './line-spacing';
import { menuGallery } from './ribbon-gallery-menu';
import {
	comboBox,
	FONT_FAMILIES,
	FONT_SIZES,
	parseFontFamily,
	parseFontSize,
} from './ribbon-combo';
import {
	colorSplit,
	group,
	launcher,
	menuSelect,
	row,
	select,
	spinner,
	splitInline,
	stack,
	tool,
} from './ribbon-parts';

const SHADINGS: Array<[string, string]> = [
	['none', 'No color'],
	['#ffff00', 'Yellow'],
	['#00ff00', 'Green'],
	['#00ffff', 'Cyan'],
	['#ff00ff', 'Magenta'],
	['#ff0000', 'Red'],
	['#0000ff', 'Blue'],
	['#ffc000', 'Gold'],
	['#e36c09', 'Orange'],
	['#c0c0c0', 'Light gray'],
	['#808080', 'Dark gray'],
];
const BORDER_PRESETS: Array<[string, string]> = [
	['bottom', 'Bottom border'],
	['top', 'Top border'],
	['left', 'Left border'],
	['right', 'Right border'],
	['none', 'No border'],
	['all', 'All borders'],
	['outside', 'Outside borders'],
	['insideH', 'Inside horizontal border'],
	['horizontal', 'Horizontal line'],
];
const FONT_COLORS: Array<[string, string]> = [
	['#000000', 'Black'],
	['#c00000', 'Red'],
	['#e36c09', 'Orange'],
	['#ffc000', 'Gold'],
	['#70ad47', 'Green'],
	['#0070c0', 'Blue'],
	['#7030a0', 'Purple'],
];
const HIGHLIGHTS: Array<[string, string]> = [
	['none', 'No highlight'],
	['yellow', 'Yellow'],
	['green', 'Green'],
	['cyan', 'Cyan'],
	['magenta', 'Magenta'],
	['blue', 'Blue'],
	['red', 'Red'],
	['darkBlue', 'Dark blue'],
	['darkCyan', 'Dark cyan'],
	['darkGreen', 'Dark green'],
	['darkMagenta', 'Dark magenta'],
	['darkRed', 'Dark red'],
	['darkYellow', 'Dark yellow'],
	['darkGray', 'Dark gray'],
	['lightGray', 'Light gray'],
	['black', 'Black'],
	['white', 'White'],
];

/** The Home tab in Word's group order: Clipboard, Font, Paragraph, Styles (added on load), Editing. */
export function buildHomePanel(panels: Map<string, HTMLElement>): void {
	const home = panels.get('Home')!;
	home.append(
		group(
			'Clipboard',
			tool('Paste', 'paste', { type: 'clipboard', key: 'paste' }, { large: true }),
			stack(
				tool('Cut', 'cut', { type: 'clipboard', key: 'cut' }, { inline: true }),
				tool('Copy', 'copy', { type: 'clipboard', key: 'copy' }, { inline: true }),
				tool('Format painter', 'formatPainter', { type: 'formatPainter' }, { inline: true }),
			),
		),
		group(
			'Font',
			row(
				comboBox(
					'Font family',
					FONT_FAMILIES,
					parseFontFamily,
					(value) => ({ type: 'font', key: 'family', value }),
					{ preview: true, className: 'font-family-combo' },
				),
				comboBox(
					'Font size',
					FONT_SIZES.map(String),
					(text) => {
						const size = parseFontSize(text);
						return size === null ? null : String(size);
					},
					(value) => ({ type: 'font', key: 'size', value }),
					{ className: 'font-size-combo' },
				),
				tool('Grow font', 'fontGrow', { type: 'fontStep', direction: 'grow' }),
				tool('Shrink font', 'fontShrink', { type: 'fontStep', direction: 'shrink' }),
				menuSelect(
					'Change case',
					'changeCase',
					[
						['sentence', 'Sentence case'],
						['lower', 'lowercase'],
						['upper', 'UPPERCASE'],
						['title', 'Capitalize Each Word'],
						['toggle', 'tOGGLE cASE'],
					],
					(value) => ({ type: 'changeCase', value: value as 'sentence' }),
					{ compact: true, momentary: true },
				),
				tool('Clear formatting', 'clear', { type: 'clear' }),
			),
			row(
				tool('Bold', 'bold', { type: 'format', key: 'bold' }),
				tool('Italic', 'italic', { type: 'format', key: 'italic' }),
				tool('Underline', 'underline', { type: 'format', key: 'underline' }),
				tool('Strikethrough', 'strike', { type: 'format', key: 'strike' }),
				tool('Subscript', 'subscript', { type: 'format', key: 'subscript' }),
				tool('Superscript', 'superscript', { type: 'format', key: 'superscript' }),
				colorSplit(
					'Text highlight',
					'highlight',
					HIGHLIGHTS,
					(value) => ({ type: 'font', key: 'highlight', value }),
					'yellow',
				),
				colorSplit(
					'Font color',
					'fontColor',
					FONT_COLORS,
					(value) => ({ type: 'font', key: 'color', value }),
					'#c00000',
					{},
				),
			),
		),
	);
	panels
		.get('Table')!
		.append(
			group(
				'Properties',
				tool(
					'Table properties',
					'table',
					{ type: 'formatDialog', kind: 'tableProperties' },
					{ large: true, caption: 'Properties' },
				),
			),
			group(
				'Table',
				tool(
					'Insert row above',
					'rowAbove',
					{ type: 'tableEdit', key: 'rowBefore' },
					{ large: true },
				),
				tool(
					'Insert row below',
					'rowBelow',
					{ type: 'tableEdit', key: 'rowAfter' },
					{ large: true },
				),
				tool('Delete row', 'deleteRow', { type: 'tableEdit', key: 'deleteRow' }, { large: true }),
				tool(
					'Insert column left',
					'columnLeft',
					{ type: 'tableEdit', key: 'columnBefore' },
					{ large: true },
				),
				tool(
					'Insert column right',
					'columnRight',
					{ type: 'tableEdit', key: 'columnAfter' },
					{ large: true },
				),
				tool(
					'Delete column',
					'deleteColumn',
					{ type: 'tableEdit', key: 'deleteColumn' },
					{ large: true },
				),
				tool(
					'Delete table',
					'deleteTable',
					{ type: 'tableEdit', key: 'deleteTable' },
					{ large: true },
				),
			),
		);
	home.append(
		group(
			'Paragraph',
			row(
				tool('Bulleted list', 'bullets', { type: 'list', key: 'bullet' }),
				tool('Numbered list', 'numbering', { type: 'list', key: 'number' }),
				menuSelect(
					'Multilevel list',
					'multilevel',
					[
						['multilevel', '1. 1.1. 1.1.1.'],
						['outline', '1. a) i.'],
						['define', 'Define New Multilevel List…'],
					],
					(value) =>
						value === 'define'
							? { type: 'formatDialog', kind: 'multilevelList' }
							: { type: 'list', key: value as 'multilevel' },
					{ compact: true, momentary: true },
				),
				tool('Decrease list level', 'outdent', { type: 'list', key: 'decreaseLevel' }),
				tool('Increase list level', 'indent', { type: 'list', key: 'increaseLevel' }),
				tool('Remove list', 'removeList', { type: 'list', key: 'remove' }),
				tool('Decrease indent', 'outdent', { type: 'paragraph', key: 'indent', value: 'decrease' }),
				tool('Increase indent', 'indent', { type: 'paragraph', key: 'indent', value: 'increase' }),
				menuSelect(
					'Sort',
					'sort',
					[
						['ascending', 'Sort A to Z'],
						['descending', 'Sort Z to A'],
					],
					(value) => ({ type: 'sort', order: value as 'ascending' }),
					{ compact: true, momentary: true },
				),
				tool('Show paragraph marks', 'paragraphMarks', { type: 'showMarks' }),
			),
			row(
				tool('Align left', 'alignLeft', { type: 'align', value: 'left' }),
				tool('Align center', 'alignCenter', { type: 'align', value: 'center' }),
				tool('Align right', 'alignRight', { type: 'align', value: 'right' }),
				tool('Justify', 'justify', { type: 'align', value: 'justify' }),
				menuGallery(
					'Line spacing',
					'lineSpacing',
					'lineSpacing',
					lineSpacingOptions,
					(value) => ({ type: 'paragraph', key: 'lineSpacing', value }),
					{
						compact: true,
						commands: [
							{
								label: 'Line Spacing Options…',
								action: { type: 'formatDialog', kind: 'paragraph' },
							},
						],
					},
				),
				colorSplit(
					'Shading',
					'shading',
					SHADINGS,
					(value) => ({ type: 'shading', value }),
					'#ffff00',
					{ noneLabel: 'No Color' },
				),
				menuGallery(
					'Borders',
					'borders',
					'borders',
					BORDER_PRESETS,
					(value) => ({ type: 'borders', preset: value as 'none' }),
					{
						compact: true,
						momentary: true,
						commands: [
							{
								label: 'Borders and Shading…',
								action: { type: 'formatDialog', kind: 'borders' },
							},
						],
					},
				),
			),
		),
		group(
			'Editing',
			stack(
				splitInline(
					tool('Find and replace', 'find', { type: 'search' }, { inline: true, caption: 'Find' }),
					'Find options',
					[
						['find', 'Find'],
						['goto', 'Go to'],
					],
					(value) => (value === 'goto' ? { type: 'goTo' } : { type: 'search' }),
				),
				tool('Replace', 'replace', { type: 'search', focus: 'replace' }, { inline: true }),
				splitInline(
					tool('Select all', 'select', { type: 'selectAll' }, { inline: true, caption: 'Select' }),
					'Select options',
					[
						['all', 'Select all'],
						['objects', 'Select objects'],
					],
					(value) => (value === 'objects' ? { type: 'selectObjects' } : { type: 'selectAll' }),
				),
			),
		),
	);
	home
		.querySelector('[data-label="Font"]')
		?.append(launcher('Font settings', { type: 'formatDialog', kind: 'font' }));
	home
		.querySelector('[data-label="Paragraph"]')
		?.append(launcher('Paragraph settings', { type: 'formatDialog', kind: 'paragraph' }));
}

/** Layout > Paragraph: paragraph spacing before and after, as in Word's Layout tab. */
export function buildParagraphSpacing(): HTMLElement[] {
	return [
		group(
			'Indent',
			spinner('Indent left', (inches) => ({ type: 'indent', side: 'left', inches }), SPIN),
			spinner('Indent right', (inches) => ({ type: 'indent', side: 'right', inches }), SPIN),
		),
		spacingGroup(),
	];
}

const SPIN = { min: 0, max: 22, step: 0.1 };

function spacingGroup(): HTMLElement {
	return group(
		'Paragraph',
		row(
			menuSelect(
				'Spacing before',
				'spaceBefore',
				[
					['inherit', 'Before: style default'],
					['0', 'Before: no paragraph space'],
					['120', 'Before: 6 pt'],
					['240', 'Before: 12 pt'],
					['360', 'Before: 18 pt'],
				],
				(value) => ({ type: 'paragraph', key: 'spacingBefore', value }),
			),
			menuSelect(
				'Spacing after',
				'spaceAfter',
				[
					['inherit', 'After: style default'],
					['0', 'After: no paragraph space'],
					['120', 'After: 6 pt'],
					['240', 'After: 12 pt'],
					['360', 'After: 18 pt'],
				],
				(value) => ({ type: 'paragraph', key: 'spacingAfter', value }),
			),
		),
	);
}
