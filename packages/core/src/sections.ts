// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { Block, SectionColumns, SectionProperties } from './model.js';
import { children, first, getW, isElement, named, type XmlElement } from './xml.js';
import { getRelationshipId } from './relationships.js';
import { isStNumberFormat, isStSectionMark, isStVerticalJc } from './generated/wml-simple-types.js';
import { enumValue } from './parse-diagnostics.js';
import { twips } from './units.js';
import { parseOnOff, parseSignedTwips, parseTwips, parseUnsignedInteger } from './simple-types.js';

const twipInt = parseTwips;
const signedTwipInt = parseSignedTwips;

export interface RawHeaderFooterRef {
	slot: 'default' | 'even' | 'first';
	rId: string;
}
/** A parsed `w:sectPr` before its header/footer relationship ids are resolved into content. */
export interface RawSection extends Omit<SectionProperties, 'headers' | 'footers'> {
	headerRefs: RawHeaderFooterRef[];
	footerRefs: RawHeaderFooterRef[];
}

function parseColumns(section: XmlElement): SectionColumns {
	const cols = first(section, 'cols');
	if (!cols) return { count: 1, equalWidth: true };
	const count = parseUnsignedInteger(getW(cols, 'num')) ?? 1;
	const spacingTwips = twipInt(getW(cols, 'space'));
	const equalWidthValue = getW(cols, 'equalWidth');
	const equalWidth = equalWidthValue === undefined || parseOnOff(equalWidthValue) !== false;
	const separator = parseOnOff(getW(cols, 'sep')) === true;
	const colChildren = children(cols, 'col');
	const widths = colChildren.length
		? colChildren.map((column) => {
				const space = twipInt(getW(column, 'space'));
				return {
					widthTwips: twipInt(getW(column, 'w')) ?? twips(0),
					...(space !== undefined && { spacingTwips: space }),
				};
			})
		: undefined;
	return {
		count,
		...(spacingTwips !== undefined ? { spacingTwips } : {}),
		equalWidth,
		...(separator ? { separator } : {}),
		...(widths ? { widths } : {}),
	};
}

function parseHeaderFooterRefs(section: XmlElement, localName: string): RawHeaderFooterRef[] {
	const refs: RawHeaderFooterRef[] = [];
	for (const element of children(section, localName)) {
		const slotValue = getW(element, 'type');
		const slot: RawHeaderFooterRef['slot'] =
			slotValue === 'even' || slotValue === 'first' ? slotValue : 'default';
		const rId = getRelationshipId(element);
		if (rId) refs.push({ slot, rId });
	}
	return refs;
}

/** `lineNumbering` and its values from `w:lnNumType`. */
function lineNumbering(
	element: XmlElement | undefined,
): Pick<RawSection, 'lineNumbering' | 'lineNumberSettings'> {
	if (!element) return {};
	const restart = getW(element, 'restart');
	const distance = twipInt(getW(element, 'distance'));
	return {
		lineNumbering: true,
		lineNumberSettings: {
			countBy: Math.max(1, parseUnsignedInteger(getW(element, 'countBy')) ?? 1),
			start: parseUnsignedInteger(getW(element, 'start')) ?? 1,
			restart: restart === 'newPage' || restart === 'newSection' ? restart : 'continuous',
			...(distance !== undefined ? { distanceTwips: distance } : {}),
		},
	};
}

function parseOneSection(section: XmlElement, endsAtBlockId: string): RawSection {
	const size = first(section, 'pgSz');
	const margins = first(section, 'pgMar');
	const typeValue = enumValue(isStSectionMark, getW(first(section, 'type'), 'val'), 'w:type');
	const pgNum = first(section, 'pgNumType');
	const vAlignValue = enumValue(isStVerticalJc, getW(first(section, 'vAlign'), 'val'), 'w:vAlign');
	const headerDistance = twipInt(getW(margins, 'header'));
	const footerDistance = twipInt(getW(margins, 'footer'));
	const gutter = twipInt(getW(margins, 'gutter'));
	const pageNumberingStart = parseUnsignedInteger(getW(pgNum, 'start'));
	const pageNumberingFormat = enumValue(isStNumberFormat, getW(pgNum, 'fmt'), 'w:pgNumType/@w:fmt');
	return {
		endsAtBlockId,
		type: typeValue ?? 'nextPage',
		pageWidthTwips: twipInt(getW(size, 'w')) ?? twips(12240),
		pageHeightTwips: twipInt(getW(size, 'h')) ?? twips(15840),
		orientation: getW(size, 'orient') === 'landscape' ? 'landscape' : 'portrait',
		marginTopTwips: signedTwipInt(getW(margins, 'top')) ?? twips(1440),
		marginRightTwips: twipInt(getW(margins, 'right')) ?? twips(1440),
		marginBottomTwips: signedTwipInt(getW(margins, 'bottom')) ?? twips(1440),
		marginLeftTwips: twipInt(getW(margins, 'left')) ?? twips(1440),
		...(headerDistance !== undefined ? { headerDistanceTwips: headerDistance } : {}),
		...(footerDistance !== undefined ? { footerDistanceTwips: footerDistance } : {}),
		...(gutter !== undefined ? { gutterTwips: gutter } : {}),
		columns: parseColumns(section),
		...(first(section, 'titlePg') ? { titlePage: true } : {}),
		...(vAlignValue ? { verticalAlign: vAlignValue } : {}),
		...(pgNum
			? {
					pageNumbering: {
						...(pageNumberingStart !== undefined ? { start: pageNumberingStart } : {}),
						...(pageNumberingFormat ? { format: pageNumberingFormat } : {}),
					},
				}
			: {}),
		...lineNumbering(first(section, 'lnNumType')),
		...(first(section, 'pgBorders') ? { pageBorders: true } : {}),
		headerRefs: parseHeaderFooterRefs(section, 'headerReference'),
		footerRefs: parseHeaderFooterRefs(section, 'footerReference'),
	};
}

/**
 * Parses every section in document order: one per paragraph whose `pPr` carries a `sectPr`
 * (that paragraph ends the section), followed by the final body-level `sectPr`. Header/footer
 * relationship ids are returned unresolved; see `document-parts.ts` for content resolution.
 */
export function parseRawSections(body: XmlElement, blocks: Block[]): RawSection[] {
	const sections: RawSection[] = [];
	let position = 0;
	for (const node of Array.from(body.childNodes).filter(isElement)) {
		if (named(node, 'p')) {
			const sectPr = first(first(node, 'pPr'), 'sectPr');
			if (sectPr) {
				const blockId = blocks[position]?.id;
				if (blockId) sections.push(parseOneSection(sectPr, blockId));
			}
			position++;
		} else if (named(node, 'tbl')) {
			position++;
		}
	}
	const finalSectPr = children(body, 'sectPr').at(-1);
	if (finalSectPr) {
		const lastBlockId = blocks.at(-1)?.id;
		if (lastBlockId) sections.push(parseOneSection(finalSectPr, lastBlockId));
	}
	return sections;
}
