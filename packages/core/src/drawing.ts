// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Parses w:drawing (wp:inline / wp:anchor pictures) and legacy w:pict (VML) into InlineImage.
// Adapted layout/EMU-conversion approach from pptx-viewer-new packages/core/src/core (DrawingML
// picture parsing); reimplemented here against @xmldom/xmldom instead of that package's DOM layer.
import type { InlineImage, PicturePlacement } from './model.js';
import {
	contentTypeForPart,
	resolveInternalTarget,
	type ContentTypes,
	type Relationship,
} from './package-parts.js';
import { definedProps } from './defined-props.js';
import { getR, isElement, type XmlElement } from './xml.js';
import { isStRelFromH, isStRelFromV } from './generated/wml-simple-types.js';
import { enumValue } from './parse-diagnostics.js';

const WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const VML_NS = 'urn:schemas-microsoft-com:vml';
const PICTURE_GRAPHIC_URI = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
/** 914400 EMU per inch, 96 CSS px per inch. */
export const EMU_PER_PIXEL = 9525;

export interface DrawingContext {
	rels: ReadonlyMap<string, Relationship>;
	contentTypes: ContentTypes;
	mediaParts: ReadonlySet<string>;
}

function firstNS(
	parent: XmlElement | undefined,
	ns: string,
	local: string,
): XmlElement | undefined {
	if (!parent) return undefined;
	for (const child of Array.from(parent.childNodes))
		if (isElement(child) && child.namespaceURI === ns && child.localName === local) return child;
	return undefined;
}
function descendantNS(
	parent: XmlElement | undefined,
	ns: string,
	local: string,
): XmlElement | undefined {
	if (!parent) return undefined;
	const stack: XmlElement[] = [parent];
	while (stack.length) {
		const node = stack.shift()!;
		for (const child of Array.from(node.childNodes)) {
			if (!isElement(child)) continue;
			if (child.namespaceURI === ns && child.localName === local) return child;
			stack.push(child);
		}
	}
	return undefined;
}
function emuToPx(value: string | null): number {
	const n = value ? Number(value) : NaN;
	return Number.isFinite(n) ? Math.round(n / EMU_PER_PIXEL) : 0;
}
function guessContentType(partName: string): string {
	const extension = partName.split('.').pop()?.toLowerCase();
	const known: Record<string, string> = {
		png: 'image/png',
		jpg: 'image/jpeg',
		jpeg: 'image/jpeg',
		gif: 'image/gif',
		bmp: 'image/bmp',
		svg: 'image/svg+xml',
		tif: 'image/tiff',
		tiff: 'image/tiff',
		emf: 'image/x-emf',
		wmf: 'image/x-wmf',
	};
	return (extension && known[extension]) || 'application/octet-stream';
}
function resolvePicture(
	relId: string,
	context: DrawingContext,
): Pick<InlineImage, 'relId' | 'partName' | 'contentType'> | undefined {
	const rel = context.rels.get(relId);
	if (!rel || rel.mode === 'External') return undefined;
	const partName = resolveInternalTarget(rel.target);
	if (!context.mediaParts.has(partName)) return undefined;
	return {
		relId,
		partName,
		contentType: contentTypeForPart(context.contentTypes, partName) ?? guessContentType(partName),
	};
}
const SVG_NS = 'http://schemas.microsoft.com/office/drawing/2016/SVG/main';
/** The relationship id of an `asvg:svgBlip` in the blip's extension list (Word 2016+ SVG). */
function svgBlipRelId(blip: XmlElement): string | undefined {
	const svgBlip = descendantNS(blip, SVG_NS, 'svgBlip');
	return (svgBlip && getR(svgBlip, 'embed')) || undefined;
}
function unsupportedKindLabel(uri: string): string {
	if (uri.includes('/chart')) return 'Chart';
	if (uri.includes('/diagram')) return 'SmartArt';
	if (uri.includes('/oleObject') || uri.includes('/ole')) return 'Embedded object';
	return 'Drawing object';
}

/** Parses a `w:drawing` (inline or floating) element into an inline image model, or an honest placeholder. */
function parseModernDrawing(node: XmlElement, context: DrawingContext): InlineImage {
	const inline = firstNS(node, WP_NS, 'inline');
	const anchor = inline ?? firstNS(node, WP_NS, 'anchor');
	const extent = firstNS(anchor, WP_NS, 'extent');
	const docPr = firstNS(anchor, WP_NS, 'docPr');
	const widthPx = emuToPx(extent?.getAttribute('cx') ?? null);
	const heightPx = emuToPx(extent?.getAttribute('cy') ?? null);
	const altText = docPr?.getAttribute('descr') || undefined;
	const title = docPr?.getAttribute('title') || undefined;
	const anchored = Boolean(anchor && anchor !== inline) || undefined;
	const placement = anchored && anchor ? parsePlacement(anchor) : undefined;
	const graphicData = descendantNS(anchor, A_NS, 'graphicData');
	const uri = graphicData?.getAttribute('uri') ?? '';
	const blip = descendantNS(graphicData, A_NS, 'blip');
	const relId = (blip && getR(blip, 'embed')) || undefined;
	const picture = uri === PICTURE_GRAPHIC_URI && relId ? resolvePicture(relId, context) : undefined;
	const svgRelId = blip && svgBlipRelId(blip);
	const svg = picture && svgRelId ? resolvePicture(svgRelId, context) : undefined;
	if (picture)
		return {
			...picture,
			widthPx,
			heightPx,
			...definedProps({ altText, title, anchored }),
			...(placement ? { placement } : {}),
			...(svg ? { svgPartName: svg.partName } : {}),
		};
	return {
		relId: relId ?? '',
		partName: '',
		contentType: 'application/octet-stream',
		widthPx,
		heightPx,
		...definedProps({ altText, title, anchored }),
		unsupported: unsupportedKindLabel(uri),
	};
}

const WRAPS: [string, PicturePlacement['wrap']][] = [
	['wrapSquare', 'square'],
	['wrapTight', 'tight'],
	['wrapThrough', 'through'],
	['wrapTopAndBottom', 'topAndBottom'],
	['wrapNone', 'none'],
];

/** Wrapping and position of a floating `wp:anchor` picture. */
function parsePlacement(anchor: XmlElement): PicturePlacement {
	const wrap = WRAPS.find(([local]) => firstNS(anchor, WP_NS, local))?.[1] ?? 'none';
	const placement: PicturePlacement = { wrap };
	if (anchor.getAttribute('behindDoc') === '1') placement.behindText = true;
	const positionH = firstNS(anchor, WP_NS, 'positionH');
	const relativeFrom = enumValue(
		isStRelFromH,
		positionH?.getAttribute('relativeFrom'),
		'wp:positionH/@relativeFrom',
	);
	if (relativeFrom) placement.relativeFrom = relativeFrom;
	const align = firstNS(positionH, WP_NS, 'align')?.textContent?.trim();
	if (
		align === 'left' ||
		align === 'center' ||
		align === 'right' ||
		align === 'inside' ||
		align === 'outside'
	)
		placement.align = align;
	const offset = firstNS(positionH, WP_NS, 'posOffset')?.textContent?.trim();
	if (offset && /^-?\d+$/.test(offset)) placement.offsetXPx = emuToPx(offset);
	const positionV = firstNS(anchor, WP_NS, 'positionV');
	const relativeFromV = enumValue(
		isStRelFromV,
		positionV?.getAttribute('relativeFrom'),
		'wp:positionV/@relativeFrom',
	);
	if (relativeFromV) placement.relativeFromV = relativeFromV;
	const alignV = firstNS(positionV, WP_NS, 'align')?.textContent?.trim();
	if (
		alignV === 'top' ||
		alignV === 'center' ||
		alignV === 'bottom' ||
		alignV === 'inside' ||
		alignV === 'outside'
	)
		placement.alignV = alignV;
	const offsetV = firstNS(positionV, WP_NS, 'posOffset')?.textContent?.trim();
	if (offsetV && /^-?\d+$/.test(offsetV)) placement.offsetYPx = emuToPx(offsetV);
	return placement;
}

/** Best-effort parse of a legacy VML `w:pict` picture (`v:shape/v:imagedata`). */
function parseVmlPicture(node: XmlElement, context: DrawingContext): InlineImage {
	const imagedata = descendantNS(node, VML_NS, 'imagedata');
	const shape = descendantNS(node, VML_NS, 'shape');
	const style = shape?.getAttribute('style') ?? '';
	const width = /width:\s*([\d.]+)pt/i.exec(style);
	const height = /height:\s*([\d.]+)pt/i.exec(style);
	const widthPx = width ? Math.round(Number(width[1]) * (96 / 72)) : 0;
	const heightPx = height ? Math.round(Number(height[1]) * (96 / 72)) : 0;
	const relId = (imagedata && getR(imagedata, 'id')) || undefined;
	const picture = relId ? resolvePicture(relId, context) : undefined;
	if (picture) return { ...picture, widthPx, heightPx };
	return {
		relId: relId ?? '',
		partName: '',
		contentType: 'application/octet-stream',
		widthPx,
		heightPx,
		unsupported: 'Legacy VML drawing',
	};
}

/** Parses a `w:drawing` or `w:pict` run child into an inline image model. */
export function parseDrawing(node: XmlElement, context: DrawingContext): InlineImage {
	return node.localName === 'pict'
		? parseVmlPicture(node, context)
		: parseModernDrawing(node, context);
}
