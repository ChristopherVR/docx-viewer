import { lineSpacingOptions } from './line-spacing';
import { button, group, row, select } from './ribbon-parts';

/** The Home tab: styles, font, paragraph, lists and alignment, and editing. */
export function buildHomePanel(panels: Map<string, HTMLElement>): void {
	const home = panels.get('Home')!;
	home.append(
		group(
			'Font',
			row(
				select(
					'Font family',
					[
						['Arial', 'Arial'],
						['Calibri', 'Calibri'],
						['Georgia', 'Georgia'],
						['Times New Roman', 'Times New Roman'],
						['Verdana', 'Verdana'],
					],
					(value) => ({ type: 'font', key: 'family', value }),
				),
				select(
					'Font size',
					['8', '9', '10', '11', '12', '14', '16', '18', '20', '24', '28', '36', '48', '72'].map(
						(v) => [v, v],
					),
					(value) => ({ type: 'font', key: 'size', value }),
				),
				button('Bold', 'B', { type: 'format', key: 'bold' }, 'tool-bold'),
				button('Italic', 'I', { type: 'format', key: 'italic' }, 'tool-italic'),
				button('Underline', 'U', { type: 'format', key: 'underline' }, 'tool-underline'),
				button('Strikethrough', 'S̶', { type: 'format', key: 'strike' }, 'tool-strike'),
			),
			row(
				button('Superscript', 'x²', { type: 'format', key: 'superscript' }),
				button('Subscript', 'x₂', { type: 'format', key: 'subscript' }),
				select(
					'Font color',
					[
						['#000000', 'Black'],
						['#c00000', 'Red'],
						['#e36c09', 'Orange'],
						['#ffc000', 'Gold'],
						['#70ad47', 'Green'],
						['#0070c0', 'Blue'],
						['#7030a0', 'Purple'],
					],
					(value) => ({ type: 'font', key: 'color', value }),
				),
				select(
					'Text highlight',
					[
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
					],
					(value) => ({ type: 'font', key: 'highlight', value }),
				),
				button('Clear formatting', 'Clear', { type: 'clear' }),
			),
		),
	);
	panels
		.get('Table')!
		.append(
			group(
				'Table',
				button('Insert row above', '↑ Row', { type: 'tableEdit', key: 'rowBefore' }),
				button('Insert row below', '↓ Row', { type: 'tableEdit', key: 'rowAfter' }),
				button('Delete row', '− Row', { type: 'tableEdit', key: 'deleteRow' }),
				button('Insert column left', '← Column', { type: 'tableEdit', key: 'columnBefore' }),
				button('Insert column right', '→ Column', { type: 'tableEdit', key: 'columnAfter' }),
				button('Delete column', '− Column', { type: 'tableEdit', key: 'deleteColumn' }),
				button('Delete table', 'Delete table', { type: 'tableEdit', key: 'deleteTable' }),
			),
		);
	home.append(
		group(
			'Paragraph',
			row(
				button('Decrease indent', '⇤', { type: 'paragraph', key: 'indent', value: 'decrease' }),
				button('Increase indent', '⇥', { type: 'paragraph', key: 'indent', value: 'increase' }),
				select(
					'Spacing after',
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
			row(
				select(
					'Spacing before',
					[
						['inherit', 'Before: style default'],
						['0', 'Before: no paragraph space'],
						['120', 'Before: 6 pt'],
						['240', 'Before: 12 pt'],
						['360', 'Before: 18 pt'],
					],
					(value) => ({ type: 'paragraph', key: 'spacingBefore', value }),
				),
				select('Line spacing', lineSpacingOptions, (value) => ({
					type: 'paragraph',
					key: 'lineSpacing',
					value,
				})),
			),
		),
	);
	home.append(
		group(
			'Alignment',
			row(
				button('Bulleted list', '•', { type: 'list', key: 'bullet' }, 'tool-list-bullet'),
				button('Numbered list', '1.', { type: 'list', key: 'number' }, 'tool-list-number'),
				button('Decrease list level', '⇤≡', { type: 'list', key: 'decreaseLevel' }),
				button('Increase list level', '⇥≡', { type: 'list', key: 'increaseLevel' }),
				button('Remove list', '✕≡', { type: 'list', key: 'remove' }),
			),
			row(
				...(
					[
						['left', 'Align left'],
						['center', 'Align center'],
						['right', 'Align right'],
						['justify', 'Justify'],
					] as const
				).map(([value, label]) =>
					button(label, value === 'center' ? '≣' : value === 'justify' ? '☰' : '≡', {
						type: 'align',
						value,
					}),
				),
			),
		),
	);
	home.append(group('Editing', button('Find and replace', 'Find and replace', { type: 'search' })));
}
