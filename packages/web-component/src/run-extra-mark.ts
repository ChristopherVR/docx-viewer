import type { TextRun } from '@christophervr/docx-core';
import type { MarkSpec } from 'prosemirror-model';

/**
 * Run properties modeled by docx-core that have no dedicated editor control yet. They travel
 * through the editor as one opaque mark so editing text never drops them from the model (and
 * so the writer never strips them from the source XML).
 */
export const extraRunFields = [
	'caps',
	'smallCaps',
	'doubleStrike',
	'vanish',
	'underlineStyle',
	'underlineColor',
	'characterSpacingTwips',
	'shadingFill',
	'shadingThemeFill',
	'colorTheme',
	'fontTheme',
] as const satisfies readonly (keyof TextRun)[];

/**
 * Toggles that have their own editor marks, carried here only when explicitly off
 * (`w:val="0"`, which cancels a style's bold, italic, strikethrough or underline).
 */
export const explicitOffFields = ['bold', 'italic', 'strike', 'underline'] as const;

export type ExtraRunProperties = Pick<
	TextRun,
	(typeof extraRunFields)[number] | (typeof explicitOffFields)[number]
>;

export function extraRunProperties(run: TextRun): ExtraRunProperties | undefined {
	const extra: Record<string, unknown> = {};
	for (const field of extraRunFields) if (run[field] !== undefined) extra[field] = run[field];
	for (const field of explicitOffFields) if (run[field] === false) extra[field] = false;
	return Object.keys(extra).length ? (extra as ExtraRunProperties) : undefined;
}

const cssColor = (value: unknown): string | undefined =>
	typeof value === 'string' && /^#[0-9a-f]{6}$/i.test(value) ? value : undefined;

function extraRunStyle(props: ExtraRunProperties): string {
	const css: string[] = [];
	if (props.caps) css.push('text-transform:uppercase');
	if (props.smallCaps) css.push('font-variant:small-caps');
	const decorations: string[] = [];
	if (props.underlineStyle) decorations.push('underline');
	if (props.doubleStrike) decorations.push('line-through');
	if (decorations.length) css.push(`text-decoration-line:${decorations.join(' ')}`);
	if (props.doubleStrike) css.push('text-decoration-style:double');
	else if (props.underlineStyle === 'double') css.push('text-decoration-style:double');
	else if (props.underlineStyle && /dot/i.test(props.underlineStyle))
		css.push('text-decoration-style:dotted');
	else if (props.underlineStyle && /dash/i.test(props.underlineStyle))
		css.push('text-decoration-style:dashed');
	else if (props.underlineStyle && /wav/i.test(props.underlineStyle))
		css.push('text-decoration-style:wavy');
	const underlineColor = cssColor(props.underlineColor);
	if (underlineColor) css.push(`text-decoration-color:${underlineColor}`);
	if (typeof props.characterSpacingTwips === 'number')
		css.push(`letter-spacing:${(props.characterSpacingTwips / 20).toFixed(2)}pt`);
	const shading = cssColor(props.shadingFill);
	if (shading) css.push(`background-color:${shading}`);
	if (props.vanish) css.push('opacity:0.45;text-decoration:underline dotted');
	return css.join(';');
}

export const runPropertiesMark: MarkSpec = {
	attrs: { props: { default: null } },
	parseDOM: [
		{
			tag: 'span[data-run-props]',
			getAttrs: (el) => {
				try {
					return { props: JSON.parse((el as HTMLElement).dataset.runProps || 'null') };
				} catch {
					return { props: null };
				}
			},
		},
	],
	toDOM: (mark) => {
		const props = (mark.attrs.props ?? {}) as ExtraRunProperties;
		const style = extraRunStyle(props);
		return [
			'span',
			{
				'data-run-props': JSON.stringify(props),
				...(props.vanish ? { class: 'dve-hidden-text' } : {}),
				...(style ? { style } : {}),
			},
			0,
		];
	},
};
