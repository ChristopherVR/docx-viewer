// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type JSZip from 'jszip';
import type { DocumentModel } from './model.js';
import type {
	AbstractNumDefinition,
	NumberingLevelDefinition,
	NumDefinition,
} from './numbering-model.js';
import {
	buildXml,
	children,
	makeW,
	parseXml,
	type XmlDocument,
	type XmlElement,
	WORD_NS,
} from './xml.js';
import { registerNumberingPart } from './numbering-parts.js';

function setAttribute(element: XmlElement, local: string, value: string): void {
	element.setAttributeNS(WORD_NS, `w:${local}`, value);
}
function valueElement(doc: XmlDocument, local: string, value: string): XmlElement {
	const element = makeW(doc, local);
	setAttribute(element, 'val', value);
	return element;
}

function buildLevelElement(doc: XmlDocument, def: NumberingLevelDefinition): XmlElement {
	const lvl = makeW(doc, 'lvl');
	setAttribute(lvl, 'ilvl', String(def.level));
	// CT_Lvl order: start, numFmt, lvlRestart, pStyle, isLgl, suff, lvlText, lvlPicBulletId, legacy, lvlJc, pPr, rPr.
	lvl.appendChild(valueElement(doc, 'start', String(def.start)));
	lvl.appendChild(valueElement(doc, 'numFmt', def.numFmt));
	if (def.lvlRestart !== undefined)
		lvl.appendChild(valueElement(doc, 'lvlRestart', String(def.lvlRestart)));
	if (def.paragraphStyleId) lvl.appendChild(valueElement(doc, 'pStyle', def.paragraphStyleId));
	if (def.isLgl) lvl.appendChild(makeW(doc, 'isLgl'));
	lvl.appendChild(
		valueElement(
			doc,
			'suff',
			def.suffix === 'space' ? 'space' : def.suffix === 'none' ? 'nothing' : 'tab',
		),
	);
	lvl.appendChild(valueElement(doc, 'lvlText', def.lvlText));
	if (def.lvlJc) lvl.appendChild(valueElement(doc, 'lvlJc', def.lvlJc));
	if (
		def.indentLeftTwips !== undefined ||
		def.hangingTwips !== undefined ||
		def.firstLineTwips !== undefined
	) {
		const pPr = makeW(doc, 'pPr');
		const ind = makeW(doc, 'ind');
		if (def.indentLeftTwips !== undefined) setAttribute(ind, 'left', String(def.indentLeftTwips));
		if (def.hangingTwips !== undefined) setAttribute(ind, 'hanging', String(def.hangingTwips));
		if (def.firstLineTwips !== undefined)
			setAttribute(ind, 'firstLine', String(def.firstLineTwips));
		pPr.appendChild(ind);
		lvl.appendChild(pPr);
	}
	const marker = def.markerFormat;
	if (marker && Object.keys(marker).length) {
		// CT_RPr order: rFonts, b, i, color, sz.
		const rPr = makeW(doc, 'rPr');
		if (marker.fontFamily) {
			const fonts = makeW(doc, 'rFonts');
			setAttribute(fonts, 'ascii', marker.fontFamily);
			setAttribute(fonts, 'hAnsi', marker.fontFamily);
			rPr.appendChild(fonts);
		}
		if (marker.bold) rPr.appendChild(makeW(doc, 'b'));
		if (marker.italic) rPr.appendChild(makeW(doc, 'i'));
		if (marker.color) rPr.appendChild(valueElement(doc, 'color', marker.color.slice(1).toUpperCase()));
		if (marker.fontSizeHalfPoints)
			rPr.appendChild(valueElement(doc, 'sz', String(marker.fontSizeHalfPoints)));
		lvl.appendChild(rPr);
	}
	return lvl;
}

function buildAbstractNumElement(doc: XmlDocument, def: AbstractNumDefinition): XmlElement {
	const element = makeW(doc, 'abstractNum');
	setAttribute(element, 'abstractNumId', def.id);
	element.appendChild(valueElement(doc, 'multiLevelType', 'hybridMultilevel'));
	for (const level of Object.values(def.levels).sort((a, b) => a.level - b.level))
		element.appendChild(buildLevelElement(doc, level));
	return element;
}

function buildNumElement(doc: XmlDocument, def: NumDefinition): XmlElement {
	const element = makeW(doc, 'num');
	setAttribute(element, 'numId', def.id);
	element.appendChild(valueElement(doc, 'abstractNumId', def.abstractNumId));
	for (const [level, override] of Object.entries(def.levelOverrides ?? {})) {
		const lvlOverride = makeW(doc, 'lvlOverride');
		setAttribute(lvlOverride, 'ilvl', level);
		if (override.startOverride !== undefined)
			lvlOverride.appendChild(valueElement(doc, 'startOverride', String(override.startOverride)));
		if (override.lvl) lvlOverride.appendChild(buildLevelElement(doc, override.lvl));
		element.appendChild(lvlOverride);
	}
	return element;
}

/**
 * Applies additive changes to the numbering catalog onto the package's `word/numbering.xml`.
 * Existing `abstractNum`/`num` entries are read-only: any difference from the loaded baseline is
 * rejected. New entries are appended in schema order and the part/relationship/content-type are
 * created the first time a document gains numbering.
 */
export async function applyNumberingCatalog(
	zip: JSZip,
	model: DocumentModel,
	binding: { base: DocumentModel } | undefined,
): Promise<void> {
	const next = model.numberingCatalog;
	const prior = binding?.base.numberingCatalog;
	if (JSON.stringify(next ?? null) === JSON.stringify(prior ?? null)) return;
	if (!binding)
		throw new Error(
			'Creating or editing numbering definitions is not supported by the standalone DOCX writer.',
		);
	if (!next)
		throw new Error(
			'Removing the numbering catalog is not supported; the source numbering.xml is preserved unchanged.',
		);
	for (const id of Object.keys(prior?.abstractNums ?? {}))
		if (JSON.stringify(next.abstractNums[id]) !== JSON.stringify(prior?.abstractNums[id]))
			throw new Error(
				`Cannot edit numbering definition abstractNum "${id}"; the original numbering.xml is preserved unchanged.`,
			);
	for (const id of Object.keys(prior?.nums ?? {}))
		if (JSON.stringify(next.nums[id]) !== JSON.stringify(prior?.nums[id]))
			throw new Error(
				`Cannot edit numbering definition num "${id}"; the original numbering.xml is preserved unchanged.`,
			);
	const newAbstractIds = Object.keys(next.abstractNums).filter((id) => !prior?.abstractNums[id]);
	const newNumIds = Object.keys(next.nums).filter((id) => !prior?.nums[id]);
	if (!newAbstractIds.length && !newNumIds.length) return;
	const existingFile = zip.file('word/numbering.xml');
	const doc: XmlDocument = existingFile
		? parseXml(await existingFile.async('string'))
		: parseXml(`<w:numbering xmlns:w="${WORD_NS}"/>`);
	const root = doc.documentElement;
	const cleanup = children(root, 'numIdMacAtCleanup')[0] ?? null;
	// CT_Numbering: abstractNum* precede num*, which precede numIdMacAtCleanup.
	const firstNum = children(root, 'num')[0] ?? cleanup;
	for (const id of newAbstractIds) {
		const abstractNum = next.abstractNums[id];
		if (abstractNum) root.insertBefore(buildAbstractNumElement(doc, abstractNum), firstNum);
	}
	for (const id of newNumIds) {
		const num = next.nums[id];
		if (num) root.insertBefore(buildNumElement(doc, num), cleanup);
	}
	zip.file('word/numbering.xml', buildXml(doc));
	if (!existingFile) await registerNumberingPart(zip);
}
