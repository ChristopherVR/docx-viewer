// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import { parseXml, WORD_NS, getW, type XmlElement } from './xml.js';
import type { ThemeCatalog, ThemeColorSlot, ThemeFontSet } from './theme-model.js';

const DRAWING_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const SCHEME_ORDER: ThemeColorSlot[] = [
	'dk1',
	'lt1',
	'dk2',
	'lt2',
	'accent1',
	'accent2',
	'accent3',
	'accent4',
	'accent5',
	'accent6',
	'hlink',
	'folHlink',
];

function drawingChild(parent: Element | undefined, local: string): Element | undefined {
	if (!parent) return undefined;
	for (const node of Array.from(parent.childNodes)) {
		const element = node as Element;
		if (
			element.nodeType === 1 &&
			element.localName === local &&
			element.namespaceURI === DRAWING_NS
		)
			return element;
	}
	return undefined;
}

function colorOf(slot: Element | undefined): string | undefined {
	const srgb = drawingChild(slot, 'srgbClr');
	if (srgb) return srgb.getAttribute('val')?.toUpperCase() || undefined;
	const sys = drawingChild(slot, 'sysClr');
	if (sys) return sys.getAttribute('lastClr')?.toUpperCase() || undefined;
	return undefined;
}

function fontSetOf(element: Element | undefined): ThemeFontSet {
	if (!element) return {};
	const latin = drawingChild(element, 'latin')?.getAttribute('typeface') || undefined;
	const ea = drawingChild(element, 'ea')?.getAttribute('typeface') || undefined;
	const cs = drawingChild(element, 'cs')?.getAttribute('typeface') || undefined;
	return {
		...(latin ? { latin } : {}),
		...(ea ? { eastAsia: ea } : {}),
		...(cs ? { complexScript: cs } : {}),
	};
}

/** Parses `word/theme/theme1.xml`'s color scheme and major/minor font scheme. */
export function parseTheme(xml: string): ThemeCatalog {
	const document = parseXml(xml);
	const root = document.documentElement;
	const clrScheme = Array.from(root.getElementsByTagNameNS(DRAWING_NS, 'clrScheme'))[0] as
		| Element
		| undefined;
	const colors: Partial<Record<ThemeColorSlot, string>> = {};
	for (const slotName of SCHEME_ORDER) {
		const slot = drawingChild(clrScheme, slotName);
		const value = colorOf(slot);
		if (value) colors[slotName] = value;
	}
	const fontScheme = Array.from(root.getElementsByTagNameNS(DRAWING_NS, 'fontScheme'))[0] as
		| Element
		| undefined;
	const major = drawingChild(fontScheme, 'majorFont');
	const minor = drawingChild(fontScheme, 'minorFont');
	return { colors, colorMapping: {}, fonts: { major: fontSetOf(major), minor: fontSetOf(minor) } };
}

const MAPPING_KEYS = ['bg1', 'tx1', 'bg2', 'tx2'] as const;
const TOKEN_TO_SLOT: Record<string, ThemeColorSlot> = {
	dark1: 'dk1',
	light1: 'lt1',
	dark2: 'dk2',
	light2: 'lt2',
};

/** Parses `word/settings.xml`'s `w:clrSchemeMapping`, mapping logical bg/tx slots to scheme slots. */
export function parseColorSchemeMapping(xml: string): ThemeCatalog['colorMapping'] {
	const document = parseXml(xml);
	const mapping = Array.from(document.getElementsByTagNameNS(WORD_NS, 'clrSchemeMapping'))[0] as
		| XmlElement
		| undefined;
	const result: ThemeCatalog['colorMapping'] = {};
	if (!mapping) return result;
	for (const key of MAPPING_KEYS) {
		const value = getW(mapping, key);
		if (!value) continue;
		result[key] = TOKEN_TO_SLOT[value] ?? (value as ThemeColorSlot);
	}
	return result;
}
