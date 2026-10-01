// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type JSZip from 'jszip';
import { ensureContentTypeOverride } from './zip-parts.js';
import { buildXml, elements, parseXml, textContent, type XmlElement } from './xml.js';

/** Direct child by local name; core properties live in the cp/dc namespaces, not `w:`. */
const child = (parent: XmlElement, local: string): XmlElement | undefined =>
	elements(parent).find((element) => element.localName === local);

/** Editable package properties (`docProps/core.xml`); absent fields are unset. */
export interface DocumentProperties {
	title?: string;
	subject?: string;
	/** `dc:creator`: Word's Author. */
	creator?: string;
	/** `cp:keywords`: Word's Tags. */
	keywords?: string;
	/** `dc:description`: Word's Comments. */
	description?: string;
	lastModifiedBy?: string;
}
export const PROPERTY_FIELDS = [
	'title',
	'subject',
	'creator',
	'keywords',
	'description',
	'lastModifiedBy',
] as const satisfies readonly (keyof DocumentProperties)[];

const CORE_PATH = 'docProps/core.xml';
const CP = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';
const DC = 'http://purl.org/dc/elements/1.1/';
const PREFIX: Record<(typeof PROPERTY_FIELDS)[number], readonly [string, string]> = {
	title: ['dc', DC],
	subject: ['dc', DC],
	creator: ['dc', DC],
	keywords: ['cp', CP],
	description: ['dc', DC],
	lastModifiedBy: ['cp', CP],
};
const CORE_TYPE = 'application/vnd.openxmlformats-package.core-properties+xml';
const CORE_REL =
	'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';

export function parseCoreProperties(xml: string): DocumentProperties | undefined {
	const root = parseXml(xml).documentElement;
	const result: DocumentProperties = {};
	for (const field of PROPERTY_FIELDS) {
		const element = child(root, field);
		const value = element ? textContent(element).trim() : '';
		if (value) result[field] = value;
	}
	return Object.keys(result).length ? result : undefined;
}

const same = (a: DocumentProperties | undefined, b: DocumentProperties | undefined) =>
	PROPERTY_FIELDS.every((field) => (a?.[field] ?? '') === (b?.[field] ?? ''));

/** Writes changed properties into `docProps/core.xml`, keeping every other element it holds. */
export async function applyCoreProperties(
	zip: JSZip,
	next: DocumentProperties | undefined,
	prior: DocumentProperties | undefined,
): Promise<void> {
	if (same(next, prior)) return;
	const existing = zip.file(CORE_PATH);
	const doc = parseXml(
		existing
			? await existing.async('string')
			: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="${CP}" xmlns:dc="${DC}"/>`,
	);
	const root = doc.documentElement;
	for (const field of PROPERTY_FIELDS) {
		if ((next?.[field] ?? '') === (prior?.[field] ?? '')) continue;
		const [prefix, ns] = PREFIX[field];
		let element: XmlElement | undefined = child(root, field);
		const value = next?.[field] ?? '';
		if (!value) {
			if (element) root.removeChild(element);
			continue;
		}
		if (!element) {
			element = doc.createElementNS(ns, `${prefix}:${field}`);
			root.appendChild(element);
		}
		element.textContent = value;
	}
	zip.file(CORE_PATH, buildXml(doc));
	if (existing) return;
	await ensureContentTypeOverride(zip, CORE_PATH, CORE_TYPE);
	const rels = zip.file('_rels/.rels');
	const relsXml =
		(await rels?.async('string')) ??
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
	if (!relsXml.includes(CORE_REL)) {
		const ids = [...relsXml.matchAll(/Id="rId(\d+)"/g)].map((match) => Number(match[1]));
		const id = `rId${Math.max(0, ...ids) + 1}`;
		zip.file(
			'_rels/.rels',
			relsXml.replace(
				'</Relationships>',
				`<Relationship Id="${id}" Type="${CORE_REL}" Target="${CORE_PATH}"/></Relationships>`,
			),
		);
	}
}
