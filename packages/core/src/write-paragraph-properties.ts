import type { Paragraph } from './model.js';
import { first, makeW, type XmlDocument, type XmlElement, WORD_NS } from './xml.js';

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
	if (!base || paragraph.indentLeftTwips !== base.indentLeftTwips) {
		setWordValue(indent, 'left', paragraph.indentLeftTwips);
	}
	if (!base || paragraph.indentRightTwips !== base.indentRightTwips) {
		setWordValue(indent, 'right', paragraph.indentRightTwips);
	}
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

export function writeParagraphProperties(
	doc: XmlDocument,
	props: XmlElement,
	paragraph: Paragraph,
	base?: Paragraph,
): void {
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
