import {
	displayListLabel,
	resolveRunFormatting,
	resolveThemeColorReference,
	type DocumentModel,
	type Paragraph,
	type ParagraphListLabel,
	type TextRun,
} from '@christophervr/docx-core';
import { cssFontStack, type TextMeasurer } from '@christophervr/docx-layout';
import type { Node } from 'prosemirror-model';
import { appendInlineNode } from './run-adapter';
import { themeFontOf } from './run-styles';

/** Position the generated marker separately from editable text; no synthetic characters enter the document. */
export function listMarkerDisplay(
	node: Node,
	paragraph: Paragraph,
	label: ParagraphListLabel | undefined,
	model: DocumentModel,
	measurer: TextMeasurer,
): { css: string; attributes: Record<string, string> } {
	if (!label || paragraph.direction === 'rtl')
		return { css: '', attributes: { 'data-list-marker': '' } };
	const runs: TextRun[] = [];
	if (node.firstChild) appendInlineNode(runs, node.firstChild);
	const format = resolveRunFormatting(runs[0] ?? { text: '' }, {
		runCatalog: model.characterStyles,
		paragraphCatalog: model.paragraphStyles,
		paragraphStyleId: paragraph.style,
	});
	const font = {
		family: format.fontFamily ?? themeFontOf(format, model.theme) ?? 'Calibri',
		sizePx: ((format.fontSize ?? 11) * 4) / 3,
		...(format.bold ? { bold: true } : {}),
		...(format.italic ? { italic: true } : {}),
	};
	const text = displayListLabel(label.text);
	const color =
		format.color ??
		(format.colorTheme && model.theme
			? resolveThemeColorReference(format.colorTheme, model.theme)
			: undefined);
	const safeColor = color && /^#[0-9a-f]{6}$/i.test(color) ? color : 'inherit';
	const width = measurer.widthOf(text, font);
	const indent =
		(paragraph.indentStartTwips ?? paragraph.indentLeftTwips ?? label.indentLeftTwips ?? 0) / 15;
	const ownFirst = paragraph.hangingTwips !== undefined || paragraph.firstLineTwips !== undefined;
	const hanging = ownFirst ? paragraph.hangingTwips : label.hangingTwips;
	const first = ownFirst ? paragraph.firstLineTwips : label.firstLineTwips;
	const anchor = indent + (hanging !== undefined ? -hanging / 15 : (first ?? 0) / 15);
	const factor = label.alignment === 'right' ? 1 : label.alignment === 'center' ? 0.5 : 0;
	const end = anchor + width * (1 - factor);
	let body = end;
	if (label.suffix === 'space') body += measurer.widthOf(' ', font);
	if (label.suffix === 'tab') {
		const custom = paragraph.tabStops
			?.filter(
				(stop) => stop.align !== 'clear' && stop.align !== 'bar' && stop.posTwips / 15 > end + 0.5,
			)
			.sort((a, b) => a.posTwips - b.posTwips)[0];
		body = custom
			? custom.posTwips / 15
			: indent > end + 0.5
				? indent
				: (Math.floor(end / 48) + 1) * 48;
	}
	return {
		attributes: { 'data-list-marker': text },
		css: `position:relative;text-indent:${body - indent}px;--dve-marker-left:${anchor - width * factor - indent}px;--dve-marker-family:${cssFontStack(font.family)};--dve-marker-size:${font.sizePx}px;--dve-marker-weight:${font.bold ? 'bold' : 'normal'};--dve-marker-style:${font.italic ? 'italic' : 'normal'};--dve-marker-color:${safeColor}`,
	};
}
