// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { Block, SectionColumns, SectionProperties } from './model.js';
import { children, first, getW, isElement, named, type XmlElement } from './xml.js';
import { getRelationshipId } from './relationships.js';

const twipInt = (value: string | undefined): number | undefined => {
	if (value === undefined || !/^-?\d+$/.test(value)) return undefined;
	const parsed = Number(value);
	return Number.isSafeInteger(parsed) ? parsed : undefined;
};
const flag = (value: string | undefined): boolean =>
	value !== undefined && !['0', 'false', 'off', 'no', 'none'].includes(value.toLowerCase());

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
	const count = twipInt(getW(cols, 'num')) ?? 1;
	const spacingTwips = twipInt(getW(cols, 'space'));
	const equalWidthValue = getW(cols, 'equalWidth');
	const equalWidth = equalWidthValue === undefined || flag(equalWidthValue);
	const separator = flag(getW(cols, 'sep'));
	const colChildren = children(cols, 'col');
	const widths = colChildren.length
		? colChildren.map((column) => ({
				widthTwips: twipInt(getW(column, 'w')) ?? 0,
				...(twipInt(getW(column, 'space')) !== undefined
					? { spacingTwips: twipInt(getW(column, 'space')) }
					: {}),
			}))
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

const SECTION_TYPES = new Set(['nextPage', 'continuous', 'evenPage', 'oddPage', 'nextColumn']);
const VALIGN_VALUES = new Set(['top', 'center', 'both', 'bottom']);

function parseOneSection(section: XmlElement, endsAtBlockId: string): RawSection {
	const size = first(section, 'pgSz');
	const margins = first(section, 'pgMar');
	const typeValue = getW(first(section, 'type'), 'val');
	const pgNum = first(section, 'pgNumType');
	const vAlignValue = getW(first(section, 'vAlign'), 'val');
	const headerDistance = twipInt(getW(margins, 'header'));
	const footerDistance = twipInt(getW(margins, 'footer'));
	const gutter = twipInt(getW(margins, 'gutter'));
	const pageNumberingStart = twipInt(getW(pgNum, 'start'));
	const pageNumberingFormat = getW(pgNum, 'fmt');
	return {
		endsAtBlockId,
		type: (typeValue && SECTION_TYPES.has(typeValue)
			? typeValue
			: 'nextPage') as SectionProperties['type'],
		pageWidthTwips: twipInt(getW(size, 'w')) ?? 12240,
		pageHeightTwips: twipInt(getW(size, 'h')) ?? 15840,
		orientation: getW(size, 'orient') === 'landscape' ? 'landscape' : 'portrait',
		marginTopTwips: twipInt(getW(margins, 'top')) ?? 1440,
		marginRightTwips: twipInt(getW(margins, 'right')) ?? 1440,
		marginBottomTwips: twipInt(getW(margins, 'bottom')) ?? 1440,
		marginLeftTwips: twipInt(getW(margins, 'left')) ?? 1440,
		...(headerDistance !== undefined ? { headerDistanceTwips: headerDistance } : {}),
		...(footerDistance !== undefined ? { footerDistanceTwips: footerDistance } : {}),
		...(gutter !== undefined ? { gutterTwips: gutter } : {}),
		columns: parseColumns(section),
		...(first(section, 'titlePg') ? { titlePage: true } : {}),
		...(vAlignValue && VALIGN_VALUES.has(vAlignValue)
			? { verticalAlign: vAlignValue as SectionProperties['verticalAlign'] }
			: {}),
		...(pgNum
			? {
					pageNumbering: {
						...(pageNumberingStart !== undefined ? { start: pageNumberingStart } : {}),
						...(pageNumberingFormat ? { format: pageNumberingFormat } : {}),
					},
				}
			: {}),
		...(first(section, 'lnNumType') ? { lineNumbering: true } : {}),
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
