// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Inline text boxes (`wp:inline` > `wps:wsp` with `wps:txbx`) the editor creates and edits. Only the
// simple form written here is editable: plain paragraphs of plain runs. Any other text box keeps
// its XML untouched.
import type { InlineImage } from './model.js';
import { EMU_PER_PIXEL } from './drawing.js';
import { isElement, makeNS, makeW, type XmlDocument, type XmlElement } from './xml.js';

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const WPS_NS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const SHAPE_GRAPHIC_URI = WPS_NS;

const el = (doc: XmlDocument, ns: string, name: string) => makeNS(doc, ns, name);
const add = (parent: XmlElement, ...items: XmlElement[]) => {
	for (const item of items) parent.appendChild(item);
	return parent;
};
const emu = (px: number) => String(Math.max(1, Math.round(px * EMU_PER_PIXEL)));
const descendants = (root: XmlElement, ns: string, local: string) =>
	Array.from(root.getElementsByTagNameNS(ns, local));

/** `w:p` elements holding one plain run each, one per line (an empty line is an empty paragraph). */
function paragraphsFor(doc: XmlDocument, lines: readonly string[]): XmlElement[] {
	return (lines.length ? lines : ['']).map((line) => {
		const paragraph = makeW(doc, 'p');
		if (line) {
			const run = makeW(doc, 'r');
			const text = makeW(doc, 't');
			if (/^\s|\s$/.test(line))
				text.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
			text.textContent = line;
			add(run, text);
			add(paragraph, run);
		}
		return paragraph;
	});
}

function outline(doc: XmlDocument, border: boolean): XmlElement {
	const ln = el(doc, A_NS, 'a:ln');
	if (!border) return add(ln, el(doc, A_NS, 'a:noFill'));
	ln.setAttribute('w', '9525');
	const color = el(doc, A_NS, 'a:srgbClr');
	color.setAttribute('val', '000000');
	return add(ln, add(el(doc, A_NS, 'a:solidFill'), color));
}

/** A new inline text box drawing with `image.textBoxText` as its paragraphs. */
export function buildTextBoxDrawing(
	doc: XmlDocument,
	image: InlineImage,
	docPrId: string,
): XmlElement {
	const drawing = makeW(doc, 'drawing');
	const inline = el(doc, WP_NS, 'wp:inline');
	for (const side of ['distT', 'distB', 'distL', 'distR']) inline.setAttribute(side, '0');
	const extent = el(doc, WP_NS, 'wp:extent');
	extent.setAttribute('cx', emu(image.widthPx));
	extent.setAttribute('cy', emu(image.heightPx));
	const effect = el(doc, WP_NS, 'wp:effectExtent');
	for (const side of ['l', 't', 'r', 'b']) effect.setAttribute(side, '0');
	const docPr = el(doc, WP_NS, 'wp:docPr');
	docPr.setAttribute('id', docPrId);
	docPr.setAttribute('name', `Text Box ${docPrId}`);
	if (image.altText) docPr.setAttribute('descr', image.altText);
	const wsp = el(doc, WPS_NS, 'wps:wsp');
	const cNvSpPr = el(doc, WPS_NS, 'wps:cNvSpPr');
	cNvSpPr.setAttribute('txBox', '1');
	const xfrm = el(doc, A_NS, 'a:xfrm');
	const off = el(doc, A_NS, 'a:off');
	off.setAttribute('x', '0');
	off.setAttribute('y', '0');
	const ext = el(doc, A_NS, 'a:ext');
	ext.setAttribute('cx', emu(image.widthPx));
	ext.setAttribute('cy', emu(image.heightPx));
	const geometry = el(doc, A_NS, 'a:prstGeom');
	geometry.setAttribute('prst', 'rect');
	add(geometry, el(doc, A_NS, 'a:avLst'));
	const fillColor = el(doc, A_NS, 'a:srgbClr');
	fillColor.setAttribute('val', 'FFFFFF');
	const spPr = add(
		el(doc, WPS_NS, 'wps:spPr'),
		add(xfrm, off, ext),
		geometry,
		add(el(doc, A_NS, 'a:solidFill'), fillColor),
		outline(doc, image.textBoxBorder !== false),
	);
	const content = add(
		el(doc, W_NS, 'w:txbxContent'),
		...paragraphsFor(doc, image.textBoxText ?? []),
	);
	const body = el(doc, WPS_NS, 'wps:bodyPr');
	for (const [name, value] of [
		['rot', '0'],
		['vert', 'horz'],
		['wrap', 'square'],
		['lIns', '91440'],
		['tIns', '45720'],
		['rIns', '91440'],
		['bIns', '45720'],
		['anchor', 't'],
		['anchorCtr', '0'],
	] as const)
		body.setAttribute(name, value);
	add(body, el(doc, A_NS, 'a:noAutofit'));
	add(wsp, cNvSpPr, spPr, add(el(doc, WPS_NS, 'wps:txbx'), content), body);
	const graphicData = el(doc, A_NS, 'a:graphicData');
	graphicData.setAttribute('uri', SHAPE_GRAPHIC_URI);
	add(
		inline,
		extent,
		effect,
		docPr,
		el(doc, WP_NS, 'wp:cNvGraphicFramePr'),
		add(el(doc, A_NS, 'a:graphic'), add(graphicData, wsp)),
	);
	return add(drawing, inline);
}

/**
 * Whether a text box is the simple form this module writes: not floating, with paragraphs that hold
 * only plain runs of text. Such a box can have its text, size and outline edited without losing XML.
 */
export function isSimpleTextBox(graphicData: XmlElement | undefined, anchored: boolean): boolean {
	if (!graphicData || anchored) return false;
	const content = descendants(graphicData, W_NS, 'txbxContent')[0];
	if (!content) return false;
	const plain = (node: XmlElement, allowed: string[]) =>
		Array.from(node.childNodes)
			.filter(isElement)
			.every((child) => child.namespaceURI === W_NS && allowed.includes(child.localName));
	return (
		plain(content, ['p']) &&
		descendants(content, W_NS, 'p').every(
			(paragraph) =>
				plain(paragraph, ['r']) &&
				descendants(paragraph, W_NS, 'r').every((run) => plain(run, ['t'])),
		)
	);
}

/** Whether the shape draws an outline (`a:ln` that is not `a:noFill`). */
export function textBoxHasBorder(graphicData: XmlElement | undefined): boolean {
	const spPr = graphicData && descendants(graphicData, WPS_NS, 'spPr')[0];
	const ln = spPr && Array.from(spPr.getElementsByTagNameNS(A_NS, 'ln'))[0];
	return Boolean(ln && !ln.getElementsByTagNameNS(A_NS, 'noFill').length);
}

/** Applies edited text, size and outline to an existing simple text box drawing in place. */
export function patchTextBox(
	doc: XmlDocument,
	drawing: XmlElement,
	image: InlineImage,
	base: InlineImage,
): void {
	if (image.widthPx !== base.widthPx || image.heightPx !== base.heightPx)
		for (const node of [
			...descendants(drawing, WP_NS, 'extent'),
			...descendants(drawing, A_NS, 'ext').filter(
				(ext) => (ext.parentNode as XmlElement | null)?.localName === 'xfrm',
			),
		]) {
			node.setAttribute('cx', emu(image.widthPx));
			node.setAttribute('cy', emu(image.heightPx));
		}
	if (JSON.stringify(image.textBoxText) !== JSON.stringify(base.textBoxText)) {
		const content = descendants(drawing, W_NS, 'txbxContent')[0];
		if (content) {
			for (const child of Array.from(content.childNodes)) content.removeChild(child);
			add(content, ...paragraphsFor(doc, image.textBoxText ?? []));
		}
	}
	if ((image.textBoxBorder !== false) !== (base.textBoxBorder !== false)) {
		const spPr = descendants(drawing, WPS_NS, 'spPr')[0];
		if (spPr) {
			for (const old of Array.from(spPr.childNodes).filter(
				(node): node is XmlElement => isElement(node) && node.localName === 'ln',
			))
				spPr.removeChild(old);
			spPr.appendChild(outline(doc, image.textBoxBorder !== false));
		}
	}
}
