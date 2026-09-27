// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type JSZip from 'jszip';
import { buildXml, children, getW, makeW, parseXml, type XmlElement, WORD_NS } from './xml.js';
import { ensureContentTypeOverride, ensureDocumentRelationship } from './zip-parts.js';

const SETTINGS_PATH = 'word/settings.xml';
const SETTINGS_CONTENT_TYPE =
	'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml';
const SETTINGS_RELATIONSHIP_TYPE =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings';

function on(element: XmlElement | undefined): boolean {
	if (!element) return false;
	const value = getW(element, 'val')?.toLowerCase();
	return !['0', 'false', 'off', 'no', 'none'].includes(value ?? '');
}

/** Reads settings.xml `w:trackChanges` (Word's "Track Changes" toggle). */
export function parseTrackChangesSetting(xml: string): boolean {
	const doc = parseXml(xml);
	const root = doc.documentElement as XmlElement | null;
	return on(root ? children(root, 'trackChanges')[0] : undefined);
}

/** Creates settings.xml (and its part/relationship/content type) the first time it's needed. */
export async function applyTrackChangesSetting(zip: JSZip, enabled: boolean): Promise<void> {
	const file = zip.file(SETTINGS_PATH);
	const xml =
		(await file?.async('string')) ??
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="${WORD_NS}"></w:settings>`;
	const doc = parseXml(xml);
	const root = doc.documentElement as XmlElement;
	for (const element of children(root, 'trackChanges')) root.removeChild(element);
	if (enabled) root.insertBefore(makeW(doc, 'trackChanges'), root.firstChild);
	zip.file(SETTINGS_PATH, buildXml(doc));
	if (!file) {
		await ensureContentTypeOverride(zip, SETTINGS_PATH, SETTINGS_CONTENT_TYPE);
		await ensureDocumentRelationship(zip, SETTINGS_RELATIONSHIP_TYPE, 'settings.xml');
	}
}
