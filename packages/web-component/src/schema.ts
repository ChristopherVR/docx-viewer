import { Schema } from 'prosemirror-model';

function parseFontSize(value: string): number | null {
	const match = /^\s*(\d+(?:\.\d+)?)\s*(pt|px)?\s*$/i.exec(value);
	if (!match) return null;
	const size = Number(match[1]) * (match[2]?.toLowerCase() === 'px' ? 0.75 : 1);
	return Number.isFinite(size) && size > 0 ? size : null;
}

const safeCssValue = (value: unknown): string => String(value ?? '').replace(/[;{}]/g, '');
const twipsCss = (value: unknown): string | null =>
	Number.isSafeInteger(value) ? `${Number(value) / 15}px` : null;

function paragraphStyle(attrs: Record<string, unknown>): string {
	const declarations = [
		attrs.align && `text-align:${safeCssValue(attrs.align)}`,
		twipsCss(attrs.spacingBeforeTwips) && `margin-top:${twipsCss(attrs.spacingBeforeTwips)}`,
		twipsCss(attrs.spacingAfterTwips) && `margin-bottom:${twipsCss(attrs.spacingAfterTwips)}`,
		twipsCss(attrs.indentLeftTwips) && `margin-left:${twipsCss(attrs.indentLeftTwips)}`,
		twipsCss(attrs.indentRightTwips) && `margin-right:${twipsCss(attrs.indentRightTwips)}`,
		twipsCss(attrs.indentStartTwips) && `margin-inline-start:${twipsCss(attrs.indentStartTwips)}`,
		twipsCss(attrs.indentEndTwips) && `margin-inline-end:${twipsCss(attrs.indentEndTwips)}`,
		twipsCss(attrs.firstLineTwips) && `text-indent:${twipsCss(attrs.firstLineTwips)}`,
		twipsCss(attrs.hangingTwips) && `text-indent:-${twipsCss(attrs.hangingTwips)}`,
		Number.isSafeInteger(attrs.lineSpacingTwips) &&
			`line-height:${
				attrs.lineSpacingRule === 'auto'
					? Number(attrs.lineSpacingTwips) / 240
					: attrs.lineSpacingRule === 'atLeast'
						? `max(1.35em, ${Number(attrs.lineSpacingTwips) / 15}px)`
						: `${Number(attrs.lineSpacingTwips) / 15}px`
			}`,
	].filter(Boolean);
	return declarations.join(';');
}

export const schema = new Schema({
	nodes: {
		doc: {
			content: 'block+',
			attrs: {
				pageWidth: { default: 816 },
				pageHeight: { default: 1056 },
				marginTop: { default: 96 },
				marginRight: { default: 96 },
				marginBottom: { default: 96 },
				marginLeft: { default: 96 },
			},
		},
		paragraph: {
			content: 'inline*',
			group: 'block',
			attrs: {
				align: { default: 'left' },
				id: { default: '' },
				style: { default: '' },
				spacingBeforeTwips: { default: null },
				spacingAfterTwips: { default: null },
				lineSpacingTwips: { default: null },
				lineSpacingRule: { default: null },
				indentLeftTwips: { default: null },
				indentRightTwips: { default: null },
				indentStartTwips: { default: null },
				indentEndTwips: { default: null },
				firstLineTwips: { default: null },
				hangingTwips: { default: null },
			},
			parseDOM: [
				{
					tag: 'p',
					getAttrs: (el) => ({
						align: (el as HTMLElement).style.textAlign || 'left',
						id: (el as HTMLElement).dataset.id || '',
						style: '',
						spacingBeforeTwips: null,
						spacingAfterTwips: null,
						lineSpacingTwips: null,
						lineSpacingRule: null,
						indentLeftTwips: null,
						indentRightTwips: null,
						indentStartTwips: null,
						indentEndTwips: null,
						firstLineTwips: null,
						hangingTwips: null,
					}),
				},
			],
			toDOM: (node) => ['p', { style: paragraphStyle(node.attrs), 'data-id': node.attrs.id }, 0],
		},
		text: { group: 'inline' },
		table: {
			content: 'tableRow+',
			group: 'block',
			attrs: { id: { default: '' } },
			parseDOM: [{ tag: 'table' }],
			toDOM: () => ['table', ['tbody', 0]],
		},
		tableRow: { content: 'tableCell+', parseDOM: [{ tag: 'tr' }], toDOM: () => ['tr', 0] },
		tableCell: {
			content: 'paragraph+',
			parseDOM: [{ tag: 'td' }, { tag: 'th' }],
			toDOM: () => ['td', 0],
		},
	},
	marks: {
		bold: { parseDOM: [{ tag: 'strong' }, { tag: 'b' }], toDOM: () => ['strong', 0] },
		italic: { parseDOM: [{ tag: 'em' }, { tag: 'i' }], toDOM: () => ['em', 0] },
		underline: { parseDOM: [{ tag: 'u' }], toDOM: () => ['u', 0] },
		font: {
			attrs: { family: { default: null }, size: { default: null }, color: { default: null } },
			parseDOM: [
				{
					tag: 'span',
					getAttrs: (el) => {
						const style = (el as HTMLElement).style;
						return {
							family: style.fontFamily.trim() || null,
							size: style.fontSize ? parseFontSize(style.fontSize) : null,
							color: style.color || null,
						};
					},
				},
			],
			toDOM: (mark) => [
				'span',
				{
					style: [
						mark.attrs.family && `font-family:${safeCssValue(mark.attrs.family)}`,
						Number.isFinite(mark.attrs.size) &&
							mark.attrs.size > 0 &&
							`font-size:${mark.attrs.size}pt`,
						mark.attrs.color && `color:${safeCssValue(mark.attrs.color)}`,
					]
						.filter(Boolean)
						.join(';'),
				},
				0,
			],
		},
	},
});
