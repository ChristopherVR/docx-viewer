// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Text watermarks: the VML `PowerPlusWaterMarkObject` shape Word puts in a header paragraph.
import type { InlineImage, WatermarkSpec } from './inline-model.js';
import { makeNS, makeW, type XmlDocument, type XmlElement } from './xml.js';

const VML_NS = 'urn:schemas-microsoft-com:vml';
const OFFICE_NS = 'urn:schemas-microsoft-com:office:office';
const WORD10_NS = 'urn:schemas-microsoft-com:office:word';
const NAMED: Record<string, string> = {
	silver: '#c0c0c0',
	gray: '#808080',
	grey: '#808080',
	black: '#000000',
	white: '#ffffff',
	red: '#ff0000',
	blue: '#0000ff',
	green: '#008000',
	yellow: '#ffff00',
};
const POINT_TO_PX = 96 / 72;

export const WATERMARK_PREFIX = 'PowerPlusWaterMarkObject';
export const DEFAULT_WATERMARK_COLOR = '#c0c0c0';

function colorOf(value: string | null): string {
	const text = (value ?? '').trim().toLowerCase().split(/\s/)[0] ?? '';
	if (NAMED[text]) return NAMED[text]!;
	const hex = /^#?([0-9a-f]{6})$/.exec(text);
	return hex ? `#${hex[1]}` : DEFAULT_WATERMARK_COLOR;
}

/** The watermark a `v:shape` describes, or undefined when it is not a Word text watermark. */
export function parseWatermarkShape(shape: XmlElement | undefined): InlineImage | undefined {
	if (!shape || !(shape.getAttribute('id') ?? '').startsWith(WATERMARK_PREFIX)) return undefined;
	const path = Array.from(shape.getElementsByTagNameNS(VML_NS, 'textpath'))[0];
	if (!path) return undefined;
	const style = shape.getAttribute('style') ?? '';
	const fill = Array.from(shape.getElementsByTagNameNS(VML_NS, 'fill'))[0];
	const opacity = fill?.getAttribute('opacity');
	const family = /font-family:\s*(?:&quot;|"|')?([^;"'&]+)/i.exec(
		path.getAttribute('style') ?? '',
	)?.[1];
	const width = /width:\s*([\d.]+)pt/i.exec(style);
	const height = /height:\s*([\d.]+)pt/i.exec(style);
	const rotation = /rotation:\s*(-?[\d.]+)/i.exec(style);
	const spec: WatermarkSpec = {
		text: path.getAttribute('string') ?? '',
		color: colorOf(shape.getAttribute('fillcolor')),
		semitransparent: opacity !== null && opacity !== '' && opacity !== '1',
		layout: rotation && Math.abs(Number(rotation[1]) % 360) > 1 ? 'diagonal' : 'horizontal',
		...(family ? { fontFamily: family.trim() } : {}),
	};
	return {
		relId: '',
		partName: '',
		contentType: 'application/octet-stream',
		widthPx: width ? Math.round(Number(width[1]) * POINT_TO_PX) : 0,
		heightPx: height ? Math.round(Number(height[1]) * POINT_TO_PX) : 0,
		unsupported: 'Watermark',
		watermark: spec,
	};
}

/** The shape's size in points: text-length driven, as Word fits the text to the shape. */
export function watermarkSizePt(spec: WatermarkSpec): { width: number; height: number } {
	const chars = Math.max(1, [...spec.text].length);
	const width = Math.min(spec.layout === 'diagonal' ? 527 : 468, Math.max(150, chars * 44));
	return { width, height: Math.round((width / (chars * 0.55)) * 100) / 100 };
}

const FORMULAS = [
	'sum #0 0 10800',
	'prod #0 2 1',
	'sum 21600 0 @1',
	'sum 0 0 @2',
	'sum 21600 0 @3',
	'if @0 @3 0',
	'if @0 21600 @1',
	'if @0 0 @2',
	'if @0 @4 21600',
	'mid @5 @6',
	'mid @8 @5',
	'mid @7 @8',
	'mid @6 @7',
	'sum @6 0 @5',
];

/** A `w:r` holding Word's text-watermark VML for `spec`, with a unique shape id. */
export function buildWatermarkRun(
	doc: XmlDocument,
	spec: WatermarkSpec,
	uniqueId: string,
): XmlElement {
	const v = (name: string) => makeNS(doc, VML_NS, `v:${name}`);
	const run = makeW(doc, 'r');
	const props = makeW(doc, 'rPr');
	props.appendChild(makeW(doc, 'noProof'));
	run.appendChild(props);
	const pict = makeW(doc, 'pict');
	const type = v('shapetype');
	type.setAttribute('id', '_x0000_t136');
	type.setAttribute('coordsize', '21600,21600');
	type.setAttributeNS(OFFICE_NS, 'o:spt', '136');
	type.setAttribute('adj', '10800');
	type.setAttribute('path', 'm@7,l@8,m@5,21600l@6,21600e');
	const formulas = v('formulas');
	for (const eqn of FORMULAS) {
		const f = v('f');
		f.setAttribute('eqn', eqn);
		formulas.appendChild(f);
	}
	const path = v('path');
	path.setAttribute('textpathok', 't');
	path.setAttributeNS(OFFICE_NS, 'o:connecttype', 'custom');
	path.setAttributeNS(OFFICE_NS, 'o:connectlocs', '@9,0;@10,10800;@11,21600;@12,10800');
	path.setAttributeNS(OFFICE_NS, 'o:connectangles', '270,180,90,0');
	const typePath = v('textpath');
	typePath.setAttribute('on', 't');
	typePath.setAttribute('fitshape', 't');
	const handles = v('handles');
	const handle = v('h');
	handle.setAttribute('position', '#0,bottomRight');
	handle.setAttribute('xrange', '6629,14971');
	handles.appendChild(handle);
	const lock = makeNS(doc, OFFICE_NS, 'o:lock');
	lock.setAttributeNS(VML_NS, 'v:ext', 'edit');
	lock.setAttribute('text', 't');
	lock.setAttribute('shapetype', 't');
	for (const child of [formulas, path, typePath, handles, lock]) type.appendChild(child);
	const size = watermarkSizePt(spec);
	const shape = v('shape');
	shape.setAttribute('id', `${WATERMARK_PREFIX}${uniqueId}`);
	shape.setAttributeNS(OFFICE_NS, 'o:spid', `_x0000_s${2048 + (Number(uniqueId) || 1)}`);
	shape.setAttribute('type', '#_x0000_t136');
	shape.setAttribute(
		'style',
		`position:absolute;margin-left:0;margin-top:0;width:${size.width}pt;height:${size.height}pt;${spec.layout === 'diagonal' ? 'rotation:315;' : ''}z-index:-251658752;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin`,
	);
	shape.setAttributeNS(OFFICE_NS, 'o:allowincell', 'f');
	shape.setAttribute('fillcolor', spec.color.replace(/^#/, '').toLowerCase());
	shape.setAttribute('stroked', 'f');
	const fill = v('fill');
	fill.setAttribute('opacity', spec.semitransparent ? '.5' : '1');
	const text = v('textpath');
	text.setAttribute('style', `font-family:"${spec.fontFamily ?? 'Calibri'}";font-size:1pt`);
	text.setAttribute('string', spec.text);
	const wrap = makeNS(doc, WORD10_NS, 'w10:wrap');
	wrap.setAttribute('anchorx', 'margin');
	wrap.setAttribute('anchory', 'margin');
	for (const child of [fill, text, wrap]) shape.appendChild(child);
	pict.appendChild(type);
	pict.appendChild(shape);
	run.appendChild(pict);
	return run;
}
