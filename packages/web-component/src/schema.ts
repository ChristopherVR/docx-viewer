import { Schema, type DOMOutputSpec } from 'prosemirror-model';
import { isWordHighlightToken, type WordHighlightToken } from '@christophervr/docx-core';
import { fieldMarkerNodeSpec, noteReferenceNodeSpec, pageBreakNodeSpec } from './break-note-schema';
import { tableStyle, tableCellStyle } from './table-render';
import { imageNodeSpec } from './inline-content-schema';
import { markSpecs } from './schema-marks';

export { wordHighlightColors } from './schema-marks';

const safeCssValue = (value: unknown): string => String(value ?? '').replace(/[;{}]/g, '');
const twipsCss = (value: unknown): string | null =>
	Number.isSafeInteger(value) ? `${Number(value) / 15}px` : null;

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
				/** Footnote/endnote number formats (`w:numFmt`), for reference labels. */
				evenAndOddHeaders: { default: false },
				footnoteNumFmt: { default: null },
				endnoteNumFmt: { default: null },
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
				/** Custom tab stops (`w:tabs`) as an array; kept for saving, not rendered at their positions. */
				tabStops: { default: null },
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
						tabStops: null,
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
		fieldMarker: fieldMarkerNodeSpec,
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
	marks: markSpecs,
});
