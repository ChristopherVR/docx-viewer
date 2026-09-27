import { Schema, type DOMOutputSpec } from 'prosemirror-model';
import { isWordHighlightToken, type WordHighlightToken } from '@christophervr/docx-core';
import { noteReferenceNodeSpec, pageBreakNodeSpec } from './break-note-schema';
import { reviewMarks } from './review-schema';
import { runPropertiesMark } from './run-extra-mark';
import { tableStyle, tableCellStyle } from './table-render';
import { imageNodeSpec, linkMarkSpec } from './inline-content-schema';

function parseFontSize(value: string): number | null {
	const match = /^\s*(\d+(?:\.\d+)?)\s*(pt|px)?\s*$/i.exec(value);
	if (!match) return null;
	const size = Number(match[1]) * (match[2]?.toLowerCase() === 'px' ? 0.75 : 1);
	return Number.isFinite(size) && size > 0 ? size : null;
}

const safeCssValue = (value: unknown): string => String(value ?? '').replace(/[;{}]/g, '');
const twipsCss = (value: unknown): string | null =>
	Number.isSafeInteger(value) ? `${Number(value) / 15}px` : null;

/** Word's highlight palette tokens mapped to stable browser colors. */
export const wordHighlightColors = {
	black: '#000000',
	blue: '#0000ff',
	cyan: '#00ffff',
	green: '#00ff00',
	magenta: '#ff00ff',
	red: '#ff0000',
	yellow: '#ffff00',
	white: '#ffffff',
	darkBlue: '#000080',
	darkCyan: '#008080',
	darkGreen: '#008000',
	darkMagenta: '#800080',
	darkRed: '#800000',
	darkYellow: '#808000',
	darkGray: '#808080',
	lightGray: '#c0c0c0',
} satisfies Record<Exclude<WordHighlightToken, 'none'>, string>;

export function paragraphStyle(attrs: Record<string, unknown>): string {
	const declarations = [
		attrs.align
			? `text-align:${safeCssValue(attrs.align)}`
			: attrs.direction === 'rtl'
				? null
				: 'text-align:left',
		attrs.direction && `direction:${safeCssValue(attrs.direction)}`,
		twipsCss(attrs.spacingBeforeTwips) && `margin-top:${twipsCss(attrs.spacingBeforeTwips)}`,
		twipsCss(attrs.spacingAfterTwips) && `margin-bottom:${twipsCss(attrs.spacingAfterTwips)}`,
		twipsCss(attrs.indentLeftTwips) && `margin-left:${twipsCss(attrs.indentLeftTwips)}`,
		twipsCss(attrs.indentRightTwips) && `margin-right:${twipsCss(attrs.indentRightTwips)}`,
		twipsCss(attrs.indentStartTwips) && `margin-inline-start:${twipsCss(attrs.indentStartTwips)}`,
		twipsCss(attrs.indentEndTwips) && `margin-inline-end:${twipsCss(attrs.indentEndTwips)}`,
		twipsCss(attrs.firstLineTwips ?? attrs.listFirstLineTwips) &&
			`text-indent:${twipsCss(attrs.firstLineTwips ?? attrs.listFirstLineTwips)}`,
		twipsCss(attrs.hangingTwips ?? attrs.listHangingTwips) &&
			`text-indent:-${twipsCss(attrs.hangingTwips ?? attrs.listHangingTwips)}`,
		twipsCss(attrs.indentLeftTwips ?? attrs.listIndentLeftTwips) &&
			attrs.indentLeftTwips == null &&
			`margin-left:${twipsCss(attrs.listIndentLeftTwips)}`,
		Number.isSafeInteger(attrs.lineSpacingTwips) &&
			`line-height:${
				attrs.lineSpacingRule == null || attrs.lineSpacingRule === 'auto'
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
				/** Section layout (no header/footer content) as JSON, so page setup is undoable. */
				sections: { default: null },
			},
		},
		paragraph: {
			content: 'inline*',
			group: 'block',
			attrs: {
				align: { default: null },
				direction: { default: null },
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
				numId: { default: null },
				ilvl: { default: null },
				listLabelText: { default: null },
				listSuffix: { default: null },
				listIndentLeftTwips: { default: null },
				listHangingTwips: { default: null },
				listFirstLineTwips: { default: null },
				pageBreakBefore: { default: false },
				/** Read-only bookmark names starting in this paragraph; not user-editable. */
				bookmarks: { default: [] },
			},
			parseDOM: [
				{
					tag: 'p',
					getAttrs: (el) => ({
						align: (el as HTMLElement).style.textAlign || null,
						direction: (el as HTMLElement).dir || (el as HTMLElement).style.direction || null,
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
						numId: (el as HTMLElement).dataset.numId
							? Number((el as HTMLElement).dataset.numId)
							: null,
						ilvl: (el as HTMLElement).dataset.ilvl
							? Number((el as HTMLElement).dataset.ilvl)
							: null,
						listLabelText: null,
						listSuffix: null,
						listIndentLeftTwips: null,
						listHangingTwips: null,
						listFirstLineTwips: null,
						pageBreakBefore: (el as HTMLElement).dataset.pageBreakBefore === 'true',
						bookmarks: (el as HTMLElement).dataset.bookmarks
							? (el as HTMLElement).dataset.bookmarks!.split(',')
							: [],
					}),
				},
			],
			toDOM: (node) => [
				'p',
				{
					style: paragraphStyle(node.attrs),
					dir: node.attrs.direction || null,
					'data-id': node.attrs.id,
					...(node.attrs.numId != null ? { 'data-num-id': String(node.attrs.numId) } : {}),
					...(node.attrs.ilvl != null ? { 'data-ilvl': String(node.attrs.ilvl) } : {}),
					...(node.attrs.listLabelText != null
						? {
								'data-list-label': `${safeCssValue(String(node.attrs.listLabelText))}${node.attrs.listSuffix === 'space' ? ' ' : node.attrs.listSuffix === 'none' ? '' : '\t'}`,
							}
						: {}),
					...(node.attrs.pageBreakBefore ? { 'data-page-break-before': 'true' } : {}),
					...(Array.isArray(node.attrs.bookmarks) && node.attrs.bookmarks.length
						? { 'data-bookmarks': node.attrs.bookmarks.join(',') }
						: {}),
				},
				0,
			],
		},
		text: { group: 'inline' },
		hardBreak: {
			group: 'inline',
			inline: true,
			atom: true,
			selectable: false,
			leafText: () => '\n',
			parseDOM: [{ tag: 'br' }],
			toDOM: () => ['br'],
		},
		pageBreak: pageBreakNodeSpec,
		noteReference: noteReferenceNodeSpec,
		image: imageNodeSpec,
		table: {
			content: 'tableRow+',
			group: 'block',
			attrs: {
				id: { default: '' },
				structureEditable: { default: true },
				widthTwips: { default: null },
				alignment: { default: null },
				indentTwips: { default: null },
				borders: { default: null },
			},
			parseDOM: [{ tag: 'table' }],
			toDOM: (node) => ['table', { style: tableStyle(node.attrs) }, ['tbody', 0]],
		},
		tableRow: { content: 'tableCell+', parseDOM: [{ tag: 'tr' }], toDOM: () => ['tr', 0] },
		tableCell: {
			content: '(paragraph | nestedTablePreview)+',
			attrs: {
				/** First paragraph id in the source cell; reconciles edits back without relying on colspan/rowspan. */
				sourceCellKey: { default: '' },
				colspan: { default: 1 },
				rowspan: { default: 1 },
				widthTwips: { default: null },
				verticalAlign: { default: null },
				shadingFill: { default: null },
				borders: { default: null },
				/** Cell margins (`w:tcMar`) in twips as JSON; display only. */
				margins: { default: null },
			},
			parseDOM: [{ tag: 'td' }, { tag: 'th' }],
			toDOM: (node) => [
				'td',
				{
					colspan: String(node.attrs.colspan || 1),
					rowspan: String(node.attrs.rowspan || 1),
					style: tableCellStyle(node.attrs),
				},
				0,
			],
		},
		/** A nested table's read-only text preview; edit the source document for nested table content. */
		nestedTablePreview: {
			atom: true,
			selectable: false,
			attrs: { rowsJson: { default: '[]' } },
			toDOM: (node): DOMOutputSpec => {
				let rows: { text: string }[][] = [];
				try {
					rows = JSON.parse(String(node.attrs.rowsJson));
				} catch {
					rows = [];
				}
				const body: DOMOutputSpec = [
					'tbody',
					...rows.map((row): DOMOutputSpec => [
						'tr',
						...row.map((cell): DOMOutputSpec => ['td', cell.text]),
					]),
				];
				return ['table', { class: 'dve-nested-preview', contenteditable: 'false' }, body];
			},
		},
	},
	marks: {
		bold: { parseDOM: [{ tag: 'strong' }, { tag: 'b' }], toDOM: () => ['strong', 0] },
		italic: { parseDOM: [{ tag: 'em' }, { tag: 'i' }], toDOM: () => ['em', 0] },
		underline: { parseDOM: [{ tag: 'u' }], toDOM: () => ['u', 0] },
		strike: {
			parseDOM: [{ tag: 's' }, { tag: 'del' }, { style: 'text-decoration=line-through' }],
			toDOM: () => ['s', 0],
		},
		highlight: {
			attrs: { color: { default: null } },
			parseDOM: [
				{
					tag: 'span[style*="background-color"]',
					getAttrs: (el) => {
						const color = (el as HTMLElement).style.backgroundColor.toLowerCase();
						const entry = Object.entries(wordHighlightColors).find(([, css]) => {
							const hex = css.slice(1);
							const rgb = `rgb(${parseInt(hex.slice(0, 2), 16)}, ${parseInt(hex.slice(2, 4), 16)}, ${parseInt(hex.slice(4, 6), 16)})`;
							return css.toLowerCase() === color || rgb === color;
						});
						return entry ? { color: entry[0] } : false;
					},
				},
			],
			toDOM: (mark) => {
				const value = String(mark.attrs.color || '');
				const token = isWordHighlightToken(value) && value !== 'none' ? value : null;
				return token
					? ['span', { style: `background-color:${wordHighlightColors[token]}` }, 0]
					: ['span', 0];
			},
		},
		verticalAlign: {
			attrs: { value: { default: 'superscript' } },
			excludes: 'verticalAlign',
			parseDOM: [
				{ tag: 'sup', attrs: { value: 'superscript' } },
				{ tag: 'sub', attrs: { value: 'subscript' } },
			],
			toDOM: (mark) => [mark.attrs.value === 'subscript' ? 'sub' : 'sup', 0],
		},
		language: {
			attrs: {
				language: { default: null },
				eastAsiaLanguage: { default: null },
				bidiLanguage: { default: null },
			},
			parseDOM: [
				{
					tag: 'span[lang], span[data-docx-east-asia-language], span[data-docx-bidi-language]',
					getAttrs: (el) => ({
						language: (el as HTMLElement).getAttribute('lang'),
						eastAsiaLanguage: (el as HTMLElement).dataset.docxEastAsiaLanguage || null,
						bidiLanguage: (el as HTMLElement).dataset.docxBidiLanguage || null,
					}),
				},
			],
			toDOM: (mark) => {
				const attrs: Record<string, string> = {};
				if (mark.attrs.language) attrs.lang = mark.attrs.language;
				if (mark.attrs.eastAsiaLanguage)
					attrs['data-docx-east-asia-language'] = mark.attrs.eastAsiaLanguage;
				if (mark.attrs.bidiLanguage) attrs['data-docx-bidi-language'] = mark.attrs.bidiLanguage;
				return ['span', attrs, 0];
			},
		},
		runRtl: {
			attrs: { value: { default: true } },
			excludes: 'runRtl',
			parseDOM: [
				{ tag: 'span[dir="rtl"]', attrs: { value: true } },
				{ tag: 'span[dir="ltr"]', attrs: { value: false } },
			],
			toDOM: (mark) => ['span', { dir: mark.attrs.value ? 'rtl' : 'ltr' }, 0],
		},
		...reviewMarks,
		runProperties: runPropertiesMark,
		link: linkMarkSpec,
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
		/** Direct `w:rStyle` character style reference; the toolbar picker edits it, decorations render it. */
		characterStyle: {
			attrs: { id: { default: null } },
			parseDOM: [
				{
					tag: 'span[data-docx-style]',
					getAttrs: (el) => ({ id: (el as HTMLElement).dataset.docxStyle || null }),
				},
			],
			toDOM: (mark) => ['span', { 'data-docx-style': mark.attrs.id }, 0],
		},
	},
});
