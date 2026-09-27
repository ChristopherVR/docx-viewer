import type {
	Paragraph,
	ParagraphFormatting,
	ParagraphStyleCatalog,
	ParagraphStyleDefinition,
} from './model.js';
import { first, getW, named, parseXml, type XmlDocument, type XmlElement } from './xml.js';

const integer = (value: string | undefined): number | undefined => {
	if (value === undefined || !/^-?\d+$/.test(value)) return undefined;
	const number = Number(value);
	return Number.isSafeInteger(number) ? number : undefined;
};
const enabled = (element: XmlElement | undefined): boolean | undefined => {
	if (!element) return undefined;
	return !['0', 'false', 'off', 'no', 'none'].includes((getW(element, 'val') ?? '').toLowerCase());
};

function parseFormatting(pPr: XmlElement | undefined): ParagraphFormatting {
	const result: ParagraphFormatting = {};
	const alignment = getW(first(pPr, 'jc'), 'val');
	if (alignment === 'left' || alignment === 'center' || alignment === 'right')
		result.align = alignment;
	if (alignment === 'both' || alignment === 'distribute') result.align = 'justify';
	const bidi = enabled(first(pPr, 'bidi'));
	if (bidi !== undefined) result.direction = bidi ? 'rtl' : 'ltr';
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
		const defaultValue = getW(element, 'default');
		const isDefault =
			defaultValue === undefined
				? undefined
				: !['0', 'false', 'off', 'no', 'none'].includes(defaultValue.toLowerCase());
		const name = getW(first(element, 'name'), 'val');
		styles[id] = {
			id,
			...(name ? { name } : {}),
			...(basedOn ? { basedOn } : {}),
			...(isDefault === undefined ? {} : { isDefault }),
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
			parent = styles[parent].basedOn;
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
	for (let index = chain.length - 1; index >= 0; index--)
		Object.assign(out, chain[index].formatting);
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
	];
	for (const key of keys) {
		const value = paragraph[key];
		if (value !== undefined) Object.assign(direct, { [key]: value });
	}
	return Object.assign(result, direct);
}
