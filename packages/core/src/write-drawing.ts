// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Writes inline picture runs: reuses original wp:inline/pic:pic XML verbatim (only patching
// extent/alt text) for unchanged pictures, and synthesizes a minimal valid drawing for insertions.
import type { InlineImage } from './model.js';
import { EMU_PER_PIXEL } from './drawing.js';
import type { RelationshipAllocator } from './relationship-allocator.js';
import {
	isElement,
	makeNS,
	makeW,
	named,
	REL_NS,
	type XmlDocument,
	type XmlElement,
} from './xml.js';

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
const PICTURE_GRAPHIC_URI = 'http://schemas.openxmlformats.org/drawingml/2006/picture';

export function findDrawingChild(run: XmlElement | undefined): XmlElement | undefined {
	if (!run) return undefined;
	for (const child of Array.from(run.childNodes))
		if (isElement(child) && named(child, 'drawing')) return child;
	return undefined;
}

function sameUnderlyingImage(a: InlineImage, b: InlineImage): boolean {
	return Boolean(a.partName) && a.relId === b.relId && a.partName === b.partName;
}

function walk(root: XmlElement, visit: (node: XmlElement) => void): void {
	const stack: XmlElement[] = [root];
	while (stack.length) {
		const node = stack.shift()!;
		visit(node);
		for (const child of Array.from(node.childNodes)) if (isElement(child)) stack.push(child);
	}
}

function updateExtent(drawing: XmlElement, widthPx: number, heightPx: number): void {
	const cx = String(Math.max(1, Math.round(widthPx * EMU_PER_PIXEL)));
	const cy = String(Math.max(1, Math.round(heightPx * EMU_PER_PIXEL)));
	walk(drawing, (node) => {
		const isExtent = node.namespaceURI === WP_NS && node.localName === 'extent';
		// Only a:xfrm/a:ext is a size; a:extLst/a:ext carries a uri and extension content.
		const parent = node.parentNode as XmlElement | null;
		const isExt =
			node.namespaceURI === A_NS &&
			node.localName === 'ext' &&
			parent?.namespaceURI === A_NS &&
			parent.localName === 'xfrm';
		if (isExtent || isExt) {
			node.setAttribute('cx', cx);
			node.setAttribute('cy', cy);
		}
	});
}

function updateDocPr(
	drawing: XmlElement,
	altText: string | undefined,
	title: string | undefined,
): void {
	walk(drawing, (node) => {
		if (node.namespaceURI !== WP_NS || node.localName !== 'docPr') return;
		if (altText) node.setAttribute('descr', altText);
		else node.removeAttribute('descr');
		if (title) node.setAttribute('title', title);
		else node.removeAttribute('title');
	});
}

function nextDocPrId(doc: XmlDocument): string {
	let max = 0;
	if (doc.documentElement)
		walk(doc.documentElement, (node) => {
			if (node.namespaceURI !== WP_NS || node.localName !== 'docPr') return;
			const id = Number(node.getAttribute('id'));
			if (Number.isFinite(id) && id > max) max = id;
		});
	return String(max + 1);
}

function el(doc: XmlDocument, ns: string, qualified: string): XmlElement {
	return makeNS(doc, ns, qualified);
}
/** @xmldom/xmldom elements only implement `appendChild`, not the newer `append(...)` convenience. */
function appendAll(parent: XmlElement, ...items: XmlElement[]): void {
	for (const item of items) parent.appendChild(item);
}

const SVG_EXTENSION_URI = '{96DAC541-7B7A-43D3-8B79-37D633B846F1}';
const SVG_NS = 'http://schemas.microsoft.com/office/drawing/2016/SVG/main';

/** Word 2016+ SVG pictures: the SVG rides in an a:blip extension beside its PNG fallback. */
function svgExtension(doc: XmlDocument, svgRelId: string): XmlElement {
	const extLst = el(doc, A_NS, 'a:extLst');
	const ext = el(doc, A_NS, 'a:ext');
	ext.setAttribute('uri', SVG_EXTENSION_URI);
	const svgBlip = doc.createElementNS(SVG_NS, 'asvg:svgBlip');
	svgBlip.setAttributeNS(REL_NS, 'r:embed', svgRelId);
	ext.appendChild(svgBlip);
	extLst.appendChild(ext);
	return extLst;
}

function buildInlineDrawing(
	doc: XmlDocument,
	image: InlineImage,
	relId: string,
	svgRelId?: string,
): XmlElement {
	const drawing = makeW(doc, 'drawing');
	const inline = el(doc, WP_NS, 'wp:inline');
	for (const side of ['distT', 'distB', 'distL', 'distR']) inline.setAttribute(side, '0');
	const cx = String(Math.max(1, Math.round(image.widthPx * EMU_PER_PIXEL)));
	const cy = String(Math.max(1, Math.round(image.heightPx * EMU_PER_PIXEL)));
	const extent = el(doc, WP_NS, 'wp:extent');
	extent.setAttribute('cx', cx);
	extent.setAttribute('cy', cy);
	const effectExtent = el(doc, WP_NS, 'wp:effectExtent');
	for (const side of ['l', 't', 'r', 'b']) effectExtent.setAttribute(side, '0');
	const docPrId = nextDocPrId(doc);
	const docPr = el(doc, WP_NS, 'wp:docPr');
	docPr.setAttribute('id', docPrId);
	docPr.setAttribute('name', `Picture ${docPrId}`);
	if (image.altText) docPr.setAttribute('descr', image.altText);
	if (image.title) docPr.setAttribute('title', image.title);
	const frameLocks = el(doc, WP_NS, 'wp:cNvGraphicFramePr');
	const graphicFrameLocks = el(doc, A_NS, 'a:graphicFrameLocks');
	graphicFrameLocks.setAttribute('noChangeAspect', '1');
	frameLocks.appendChild(graphicFrameLocks);
	const cNvPr = el(doc, PIC_NS, 'pic:cNvPr');
	cNvPr.setAttribute('id', '0');
	cNvPr.setAttribute('name', `Picture ${docPrId}`);
	const nvPicPr = el(doc, PIC_NS, 'pic:nvPicPr');
	appendAll(nvPicPr, cNvPr, el(doc, PIC_NS, 'pic:cNvPicPr'));
	const blip = el(doc, A_NS, 'a:blip');
	blip.setAttributeNS(REL_NS, 'r:embed', relId);
	if (svgRelId) blip.appendChild(svgExtension(doc, svgRelId));
	const stretch = el(doc, A_NS, 'a:stretch');
	stretch.appendChild(el(doc, A_NS, 'a:fillRect'));
	const blipFill = el(doc, PIC_NS, 'pic:blipFill');
	appendAll(blipFill, blip, stretch);
	const xfrm = el(doc, A_NS, 'a:xfrm');
	const off = el(doc, A_NS, 'a:off');
	off.setAttribute('x', '0');
	off.setAttribute('y', '0');
	const ext = el(doc, A_NS, 'a:ext');
	ext.setAttribute('cx', cx);
	ext.setAttribute('cy', cy);
	appendAll(xfrm, off, ext);
	const prstGeom = el(doc, A_NS, 'a:prstGeom');
	prstGeom.setAttribute('prst', 'rect');
	prstGeom.appendChild(el(doc, A_NS, 'a:avLst'));
	const spPr = el(doc, PIC_NS, 'pic:spPr');
	appendAll(spPr, xfrm, prstGeom);
	const pic = el(doc, PIC_NS, 'pic:pic');
	appendAll(pic, nvPicPr, blipFill, spPr);
	const graphicData = el(doc, A_NS, 'a:graphicData');
	graphicData.setAttribute('uri', PICTURE_GRAPHIC_URI);
	graphicData.appendChild(pic);
	const graphic = el(doc, A_NS, 'a:graphic');
	graphic.appendChild(graphicData);
	appendAll(inline, extent, effectExtent, docPr, frameLocks, graphic);
	drawing.appendChild(inline);
	return drawing;
}

function clearNonProperties(node: XmlElement, keep?: XmlElement): void {
	for (const child of Array.from(node.childNodes)) {
		if (isElement(child) && (named(child, 'rPr') || child === keep)) continue;
		node.removeChild(child);
	}
}

/**
 * Builds or reuses the run for an inline image. Reuses the original drawing XML verbatim
 * (patching only extent/alt text) when the run is recognizably the same picture; synthesizes a
 * minimal, valid `wp:inline` drawing for a newly inserted picture. Refuses to swap one picture's
 * relationship for another in place rather than silently dropping the original XML.
 */
export function createImageRun(
	doc: XmlDocument,
	image: InlineImage,
	base: InlineImage | undefined,
	old: XmlElement | undefined,
	allocator?: RelationshipAllocator,
): XmlElement {
	const node = old ?? makeW(doc, 'r');
	const oldDrawing = findDrawingChild(old);
	if (!base && oldDrawing)
		throw new Error(
			'Cannot safely insert a picture here: this position lines up with a different existing picture. Edit further away from surrounding pictures, then save.',
		);
	if (base && oldDrawing && sameUnderlyingImage(base, image)) {
		if (image.widthPx !== base.widthPx || image.heightPx !== base.heightPx)
			updateExtent(oldDrawing, image.widthPx, image.heightPx);
		if (image.altText !== base.altText || image.title !== base.title)
			updateDocPr(oldDrawing, image.altText, image.title);
		clearNonProperties(node, oldDrawing);
		return node;
	}
	if (base)
		throw new Error(
			'Cannot replace an existing picture with a different one in place; remove it and insert a new picture instead.',
		);
	if (!allocator) throw new Error('Cannot insert a new picture without a relationship allocator.');
	const relId = allocator.addImage(image.partName);
	const svgRelId = image.svgPartName ? allocator.addImage(image.svgPartName) : undefined;
	const drawing = buildInlineDrawing(doc, image, relId, svgRelId);
	clearNonProperties(node);
	node.appendChild(drawing);
	return node;
}
