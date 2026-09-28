import type {
	Paragraph,
	ParagraphFormatting,
	ParagraphStyleCatalog,
	ParagraphStyleDefinition,
} from './model.js';
import { first, getW, named, parseXml, type XmlDocument, type XmlElement } from './xml.js';
import { parseParagraphBorders, parseShadingFill } from './table-borders.js';
import { alignFromJustification, parseJustification } from './paragraph-alignment.js';
import { onOffElement, parseInteger, parseOnOff, parseSignedTwips } from './simple-types.js';

const integer = parseSignedTwips;
const enabled = onOffElement;

/** Pagination toggles whose element names match their model keys. */
export const PAGINATION_KEYS = [
	'keepNext',
	'keepLines',
	'widowControl',
	'contextualSpacing',
] as const;

/** Parses `w:pPr` formatting (shared by styles, docDefaults and direct paragraph properties). */
export function parseFormatting(pPr: XmlElement | undefined): ParagraphFormatting {
	const result: ParagraphFormatting = {};
	const { align, justification } = parseJustification(pPr);
	if (justification) result.justification = justification;
	if (align) result.align = align;
	const bidi = enabled(first(pPr, 'bidi'));
	if (bidi !== undefined) result.direction = bidi ? 'rtl' : 'ltr';
	for (const key of PAGINATION_KEYS) {
		const value = enabled(first(pPr, key));
		if (value !== undefined) result[key] = value;
	}
	const borders = parseParagraphBorders(first(pPr, 'pBdr'));
	if (borders) result.borders = borders;
	const shading = parseShadingFill(first(pPr, 'shd'));
	if (shading) result.shadingFill = shading;
	const spacing = first(pPr, 'spacing');
	const before = integer(getW(spacing, 'before'));
	const after = integer(getW(spacing, 'after'));
	const line = integer(getW(spacing, 'line'));
	if (before !== undefined) result.spacingBeforeTwips = before;
	if (after !== undefined) result.spacingAfterTwips = after;
	if (line !== undefined) result.lineSpacingTwips = line;
	const rule = getW(spacing, 'lineRule');
	if (rule === 'auto' || rule === 'exact' || rule === 'atLeast') result.lineSpacingRule = rule;
	else if (line !== undefined) result.lineSpacingRule = 'auto';
	const ind = first(pPr, 'ind');
	for (const [xml, key] of [
		['left', 'indentLeftTwips'],
		['right', 'indentRightTwips'],
		['start', 'indentStartTwips'],
		['end', 'indentEndTwips'],
		['firstLine', 'firstLineTwips'],
		['hanging', 'hangingTwips'],
	] as const) {
		const value = integer(getW(ind, xml));
		if (value !== undefined) result[key] = value;
	}
	return result;
}

function parseStyleNumbering(
	pPr: XmlElement | undefined,
): { numId: number; level: number } | undefined {
	const numPr = first(pPr, 'numPr');
	if (!numPr) return undefined;
	const numId = parseInteger(getW(first(numPr, 'numId'), 'val'));
	if (numId === undefined) return undefined;
	return { numId, level: parseInteger(getW(first(numPr, 'ilvl'), 'val')) ?? 0 };
}

function styleElements(document: XmlDocument): XmlElement[] {
	return Array.from(document.getElementsByTagName('*')).filter(
		(element): element is XmlElement =>
			named(element, 'style') && getW(element, 'type') === 'paragraph',
	);
}

export function parseParagraphStyleCatalog(xml: string): ParagraphStyleCatalog {
	const document = parseXml(xml);
	const defaults = first(first(document.documentElement, 'docDefaults'), 'pPrDefault');
	const pDefaults = first(defaults, 'pPr');
	const styles = Object.create(null) as Record<string, ParagraphStyleDefinition>;
	for (const element of styleElements(document)) {
		const id = getW(element, 'styleId');
		if (!id) continue;
		const basedOn = getW(first(element, 'basedOn'), 'val');
		const isDefault = parseOnOff(getW(element, 'default'));
		const name = getW(first(element, 'name'), 'val');
		const numbering = parseStyleNumbering(first(element, 'pPr'));
		styles[id] = {
			id,
			...(name ? { name } : {}),
			...(basedOn ? { basedOn } : {}),
			...(isDefault === undefined ? {} : { isDefault }),
			...(numbering ? { numbering } : {}),
			formatting: parseFormatting(first(element, 'pPr')),
		};
	}
	const warnings: string[] = [];
	for (const style of Object.values(styles)) {
		const seen = new Set<string>([style.id]);
		let parent = style.basedOn;
		while (parent && styles[parent]) {
			if (seen.has(parent)) {
				warnings.push(
					`Paragraph style inheritance cycle detected at “${parent}”; cyclic inheritance is ignored.`,
				);
				break;
			}
			seen.add(parent);
			parent = styles[parent]?.basedOn;
		}
	}
	return { docDefaults: parseFormatting(pDefaults), styles, warnings: [...new Set(warnings)] };
}

function applyStyle(
	styleId: string | undefined,
	catalog: ParagraphStyleCatalog,
	out: ParagraphFormatting,
): void {
	const chain: ParagraphStyleDefinition[] = [];
	const visited = new Set<string>();
	let current = styleId;
	while (current && !visited.has(current)) {
		visited.add(current);
		const style = catalog.styles[current];
		if (!style) break;
		chain.push(style);
		current = style.basedOn;
	}
	for (const style of chain.reverse()) Object.assign(out, style.formatting);
}

/** Resolves defaults and style ancestry without modifying paragraph direct properties. */
export function resolveParagraphFormatting(
	paragraph: Paragraph,
	catalog: ParagraphStyleCatalog,
): ParagraphFormatting {
	const result: ParagraphFormatting = { ...catalog.docDefaults };
	const explicitStyle = paragraph.style;
	const defaultStyle = Object.values(catalog.styles).find((style) => style.isDefault)?.id;
	applyStyle(explicitStyle || defaultStyle, catalog, result);
	const direct: ParagraphFormatting = {};
	const keys: (keyof ParagraphFormatting)[] = [
		'align',
		'justification',
		'direction',
		'spacingBeforeTwips',
		'spacingAfterTwips',
		'lineSpacingTwips',
		'lineSpacingRule',
		'indentLeftTwips',
		'indentRightTwips',
		'indentStartTwips',
		'indentEndTwips',
		'firstLineTwips',
		'hangingTwips',
		...PAGINATION_KEYS,
		'borders',
		'shadingFill',
	];
	for (const key of keys) {
		const value = paragraph[key];
		if (value !== undefined) Object.assign(direct, { [key]: value });
	}
	Object.assign(result, direct);
	// A direct alignment without its own `w:jc` value supersedes any inherited exact value.
	if (paragraph.align !== undefined && paragraph.justification === undefined)
		delete result.justification;
	// `start`/`end` follow the resolved direction, which may be inherited from a style.
	if (
		(result.justification === 'start' || result.justification === 'end') &&
		result.direction === 'rtl' &&
		(paragraph.align === undefined || paragraph.justification !== undefined) &&
		result.align === alignFromJustification(result.justification, false)
	) {
		const align = alignFromJustification(result.justification, true);
		if (align === undefined) delete result.align;
		else result.align = align;
	}
	return result;
}

/** Resolves numbering inherited through `pStyle` when a paragraph has no direct `w:numPr`. */
export function resolveStyleNumbering(
	styleId: string | undefined,
	catalog: ParagraphStyleCatalog,
): { numId: number; level: number } | undefined {
	const chain: ParagraphStyleDefinition[] = [];
	const visited = new Set<string>();
	let current = styleId;
	while (current && !visited.has(current)) {
		visited.add(current);
		const style = catalog.styles[current];
		if (!style) break;
		chain.push(style);
		current = style.basedOn;
	}
	return chain.find((style) => style.numbering)?.numbering;
}
