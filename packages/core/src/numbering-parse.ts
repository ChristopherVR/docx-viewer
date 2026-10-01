// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type {
	AbstractNumDefinition,
	NumberingCatalog,
	NumberingLevelDefinition,
	NumberingMarkerFormat,
	NumDefinition,
	NumLevelOverride,
} from './numbering-model.js';
import { children, first, getW, named, parseXml, type XmlElement } from './xml.js';
import { isStNumberFormat } from './generated/wml-simple-types.js';
import { enumValue } from './parse-diagnostics.js';
import { onOffElement, parseInteger, parseSignedTwips, parseTwips } from './simple-types.js';
import { signedTwips, twips } from './units.js';

const integer = (value: string | undefined, fallback: number): number =>
	parseInteger(value) ?? fallback;
const flag = (element: XmlElement | undefined): boolean => onOffElement(element) === true;
function byName(root: XmlElement, local: string): XmlElement[] {
	return Array.from(root.getElementsByTagName('*')).filter(
		(node): node is XmlElement => node.nodeType === 1 && named(node as XmlElement, local),
	);
}

function parseMarkerFormat(rPr: XmlElement | undefined): NumberingMarkerFormat | undefined {
	if (!rPr) return undefined;
	const format: NumberingMarkerFormat = {};
	const family = getW(first(rPr, 'rFonts'), 'ascii');
	if (family) format.fontFamily = family;
	const size = parseInteger(getW(first(rPr, 'sz'), 'val'));
	if (size !== undefined && size > 0) format.fontSizeHalfPoints = size;
	if (onOffElement(first(rPr, 'b')) === true) format.bold = true;
	if (onOffElement(first(rPr, 'i')) === true) format.italic = true;
	const color = getW(first(rPr, 'color'), 'val');
	if (color && /^[0-9a-f]{6}$/i.test(color)) format.color = `#${color.toLowerCase()}`;
	return Object.keys(format).length ? format : undefined;
}

export function parseNumberingLevel(lvl: XmlElement): NumberingLevelDefinition {
	const level = integer(getW(lvl, 'ilvl'), 0);
	const start = integer(getW(first(lvl, 'start'), 'val'), 1);
	const numFmt =
		enumValue(isStNumberFormat, getW(first(lvl, 'numFmt'), 'val'), 'w:numFmt') ?? 'decimal';
	const lvlText = getW(first(lvl, 'lvlText'), 'val') ?? '';
	const jc = getW(first(lvl, 'lvlJc'), 'val');
	const ind = first(first(lvl, 'pPr'), 'ind');
	const suffRaw = getW(first(lvl, 'suff'), 'val');
	const result: NumberingLevelDefinition = { level, start, numFmt, lvlText };
	const paragraphStyleId = getW(first(lvl, 'pStyle'), 'val');
	if (paragraphStyleId) result.paragraphStyleId = paragraphStyleId;
	if (jc === 'left' || jc === 'center' || jc === 'right') result.lvlJc = jc;
	const left = getW(ind, 'left') ?? getW(ind, 'start');
	if (left !== undefined) result.indentLeftTwips = parseSignedTwips(left) ?? signedTwips(0);
	const hanging = getW(ind, 'hanging');
	if (hanging !== undefined) result.hangingTwips = parseTwips(hanging) ?? twips(0);
	const firstLine = getW(ind, 'firstLine');
	if (firstLine !== undefined) result.firstLineTwips = parseTwips(firstLine) ?? twips(0);
	if (flag(first(lvl, 'isLgl'))) result.isLgl = true;
	const restart = getW(first(lvl, 'lvlRestart'), 'val');
	if (restart !== undefined) result.lvlRestart = integer(restart, 0);
	const marker = parseMarkerFormat(first(lvl, 'rPr'));
	if (marker) result.markerFormat = marker;
	result.suffix = suffRaw === 'space' ? 'space' : suffRaw === 'nothing' ? 'none' : 'tab';
	return result;
}

function parseAbstractNum(element: XmlElement): AbstractNumDefinition | undefined {
	const id = getW(element, 'abstractNumId');
	if (!id) return undefined;
	const levels: Record<number, NumberingLevelDefinition> = {};
	for (const lvl of children(element, 'lvl')) {
		const parsed = parseNumberingLevel(lvl);
		levels[parsed.level] = parsed;
	}
	return { id, levels };
}

function parseNum(element: XmlElement): NumDefinition | undefined {
	const id = getW(element, 'numId');
	if (!id) return undefined;
	const abstractNumId = getW(first(element, 'abstractNumId'), 'val') ?? '';
	const levelOverrides: Record<number, NumLevelOverride> = {};
	for (const override of children(element, 'lvlOverride')) {
		const level = integer(getW(override, 'ilvl'), 0);
		const entry: NumLevelOverride = {};
		const startOverride = getW(first(override, 'startOverride'), 'val');
		if (startOverride !== undefined) entry.startOverride = integer(startOverride, 1);
		const lvl = first(override, 'lvl');
		if (lvl) entry.lvl = parseNumberingLevel(lvl);
		if (Object.keys(entry).length) levelOverrides[level] = entry;
	}
	return {
		id,
		abstractNumId,
		...(Object.keys(levelOverrides).length ? { levelOverrides } : {}),
	};
}

/** Parses `word/numbering.xml` into a read-only catalog; editing existing entries is not supported. */
export function parseNumberingCatalog(xml: string): NumberingCatalog {
	const document = parseXml(xml);
	const root = document.documentElement;
	const abstractNums: Record<string, AbstractNumDefinition> = {};
	const warnings: string[] = [];
	for (const element of byName(root, 'abstractNum')) {
		const parsed = parseAbstractNum(element);
		if (!parsed) continue;
		abstractNums[parsed.id] = parsed;
		const link = getW(first(element, 'numStyleLink'), 'val');
		if (link)
			warnings.push(
				`Numbering definition abstractNum ${parsed.id} links to style "${link}"; style-linked numbering levels are not resolved.`,
			);
		if (byName(element, 'lvlPicBulletId').length)
			warnings.push(
				`Numbering definition abstractNum ${parsed.id} uses a picture bullet; picture bullets are not rendered.`,
			);
	}
	const nums: Record<string, NumDefinition> = {};
	for (const element of byName(root, 'num')) {
		const parsed = parseNum(element);
		if (parsed) nums[parsed.id] = parsed;
	}
	return { abstractNums, nums, warnings: [...new Set(warnings)] };
}

/** Resolves the effective level definition for a `numId`, applying any `lvlOverride`. */
export function resolveNumberingLevel(
	catalog: NumberingCatalog,
	numId: string,
	level: number,
): NumberingLevelDefinition | undefined {
	const num = catalog.nums[numId];
	if (!num) return undefined;
	const override = num.levelOverrides?.[level];
	const base = override?.lvl ?? catalog.abstractNums[num.abstractNumId]?.levels[level];
	if (!base) return undefined;
	return override?.startOverride !== undefined ? { ...base, start: override.startOverride } : base;
}
