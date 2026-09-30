import type { Paragraph } from './model.js';
import { writeParagraphDecoration } from './write-paragraph-decoration.js';
import { children, first, makeW, type XmlDocument, type XmlElement, WORD_NS } from './xml.js';

function setAttribute(element: XmlElement, local: string, value: string): void {
	element.setAttributeNS(WORD_NS, `w:${local}`, value);
}

function setWordValue(
	element: XmlElement,
	local: string,
	value: number | string | undefined,
): void {
	if (value === undefined) element.removeAttributeNS(WORD_NS, local);
	else setAttribute(element, local, String(value));
}

function updateDirection(
	doc: XmlDocument,
	props: XmlElement,
	paragraph: Paragraph,
	base?: Paragraph,
): void {
	if (base && paragraph.direction === base.direction) return;
	let bidi = first(props, 'bidi');
	if (paragraph.direction === undefined) {
		for (const element of children(props, 'bidi')) props.removeChild(element);
	} else {
		if (!bidi) {
			bidi = makeW(doc, 'bidi');
			props.appendChild(bidi);
		}
		setAttribute(bidi, 'val', paragraph.direction === 'rtl' ? '1' : '0');
	}
}

function updateSpacing(
	doc: XmlDocument,
	props: XmlElement,
	paragraph: Paragraph,
	base?: Paragraph,
): void {
	let spacing = first(props, 'spacing');
	if (!spacing) {
		spacing = makeW(doc, 'spacing');
		props.appendChild(spacing);
	}
	if (!base || paragraph.spacingBeforeTwips !== base.spacingBeforeTwips)
		setWordValue(spacing, 'before', paragraph.spacingBeforeTwips);
	if (!base || paragraph.spacingAfterTwips !== base.spacingAfterTwips)
		setWordValue(spacing, 'after', paragraph.spacingAfterTwips);
	if (
		!base ||
		paragraph.lineSpacingTwips !== base.lineSpacingTwips ||
		paragraph.lineSpacingRule !== base.lineSpacingRule
	) {
		setWordValue(spacing, 'line', paragraph.lineSpacingTwips);
		setWordValue(
			spacing,
			'lineRule',
			paragraph.lineSpacingRule ?? (paragraph.lineSpacingTwips === undefined ? undefined : 'auto'),
		);
	}
	if (!spacing.attributes.length && !spacing.childNodes.length) props.removeChild(spacing);
}

function updateIndent(
	doc: XmlDocument,
	props: XmlElement,
	paragraph: Paragraph,
	base?: Paragraph,
): void {
	let indent = first(props, 'ind');
	if (!indent) {
		indent = makeW(doc, 'ind');
		props.appendChild(indent);
	}
	if (!base || paragraph.indentLeftTwips !== base.indentLeftTwips)
		setWordValue(indent, 'left', paragraph.indentLeftTwips);
	if (!base || paragraph.indentRightTwips !== base.indentRightTwips)
		setWordValue(indent, 'right', paragraph.indentRightTwips);
	if (!base || paragraph.indentStartTwips !== base.indentStartTwips)
		setWordValue(indent, 'start', paragraph.indentStartTwips);
	if (!base || paragraph.indentEndTwips !== base.indentEndTwips)
		setWordValue(indent, 'end', paragraph.indentEndTwips);
	if (!base || paragraph.firstLineTwips !== base.firstLineTwips)
		setWordValue(indent, 'firstLine', paragraph.firstLineTwips);
	if (!base || paragraph.hangingTwips !== base.hangingTwips)
		setWordValue(indent, 'hanging', paragraph.hangingTwips);
	if (!indent.attributes.length) props.removeChild(indent);
}

function updatePageBreakBefore(
	doc: XmlDocument,
	props: XmlElement,
	paragraph: Paragraph,
	base?: Paragraph,
): void {
	if (base && paragraph.pageBreakBefore === base.pageBreakBefore) return;
	for (const element of children(props, 'pageBreakBefore')) props.removeChild(element);
	if (paragraph.pageBreakBefore) props.appendChild(makeW(doc, 'pageBreakBefore'));
}

/** Pagination toggles: on, explicitly off (cancelling a style), or absent. */
function updateKeepOptions(
	doc: XmlDocument,
	props: XmlElement,
	paragraph: Paragraph,
	base?: Paragraph,
): void {
	for (const key of ['keepNext', 'keepLines', 'widowControl', 'contextualSpacing'] as const) {
		if (base && paragraph[key] === base[key]) continue;
		for (const element of children(props, key)) props.removeChild(element);
		const value = paragraph[key];
		if (value === undefined) continue;
		const element = makeW(doc, key);
		if (!value) setAttribute(element, 'val', '0');
		props.appendChild(element);
	}
}

/** `w:framePr` for a drop cap; other frame attributes of a loaded frame are kept as they are. */
function updateDropCap(
	doc: XmlDocument,
	props: XmlElement,
	paragraph: Paragraph,
	base?: Paragraph,
): void {
	if (base && JSON.stringify(paragraph.dropCap) === JSON.stringify(base.dropCap)) return;
	let frame = first(props, 'framePr');
	const cap = paragraph.dropCap;
	if (!cap) {
		if (frame) props.removeChild(frame);
		return;
	}
	if (!frame) {
		frame = makeW(doc, 'framePr');
		props.appendChild(frame);
		setAttribute(frame, 'wrap', 'around');
		setAttribute(frame, 'vAnchor', 'text');
		setAttribute(frame, 'hAnchor', 'text');
	}
	setAttribute(frame, 'dropCap', cap.style);
	setAttribute(frame, 'lines', String(cap.lines));
}

export function writeParagraphProperties(
	doc: XmlDocument,
	props: XmlElement,
	paragraph: Paragraph,
	base?: Paragraph,
): void {
	updatePageBreakBefore(doc, props, paragraph, base);
	updateKeepOptions(doc, props, paragraph, base);
	updateDropCap(doc, props, paragraph, base);
	updateDirection(doc, props, paragraph, base);
	writeParagraphDecoration(doc, props, paragraph, base);
	const spacingKeys = [
		'spacingBeforeTwips',
		'spacingAfterTwips',
		'lineSpacingTwips',
		'lineSpacingRule',
	] as const;
	if (!base || spacingKeys.some((key) => paragraph[key] !== base[key]))
		updateSpacing(doc, props, paragraph, base);
	const indentKeys = [
		'indentLeftTwips',
		'indentRightTwips',
		'indentStartTwips',
		'indentEndTwips',
		'firstLineTwips',
		'hangingTwips',
	] as const;
	if (!base || indentKeys.some((key) => paragraph[key] !== base[key]))
		updateIndent(doc, props, paragraph, base);
}
