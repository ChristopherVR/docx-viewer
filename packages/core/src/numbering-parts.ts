// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type JSZip from 'jszip';
import { buildXml, parseXml, type XmlElement } from './xml.js';

const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const NUMBERING_RELATIONSHIP_TYPE =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering';
const NUMBERING_CONTENT_TYPE =
	'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml';

/** Adds the content-type override and word-part relationship for a newly created `word/numbering.xml`. */
export async function registerNumberingPart(zip: JSZip): Promise<void> {
	const contentTypesPath = '[Content_Types].xml';
	const contentTypesFile = zip.file(contentTypesPath);
	const contentTypesXml = contentTypesFile
		? await contentTypesFile.async('string')
		: `<?xml version="1.0"?><Types xmlns="${CT_NS}"/>`;
	const contentTypesDoc = parseXml(contentTypesXml);
	const hasOverride = Array.from(contentTypesDoc.getElementsByTagNameNS(CT_NS, 'Override')).some(
		(element) => (element as XmlElement).getAttribute('PartName') === '/word/numbering.xml',
	);
	if (!hasOverride) {
		const override = contentTypesDoc.createElementNS(CT_NS, 'Override');
		override.setAttribute('PartName', '/word/numbering.xml');
		override.setAttribute('ContentType', NUMBERING_CONTENT_TYPE);
		contentTypesDoc.documentElement.appendChild(override);
		zip.file(contentTypesPath, buildXml(contentTypesDoc));
	}
	const relsPath = 'word/_rels/document.xml.rels';
	const relsFile = zip.file(relsPath);
	const relsXml = relsFile
		? await relsFile.async('string')
		: `<?xml version="1.0"?><Relationships xmlns="${REL_NS}"/>`;
	const relsDoc = parseXml(relsXml);
	const existingIds = new Set(
		Array.from(relsDoc.getElementsByTagNameNS(REL_NS, 'Relationship')).map(
			(element) => (element as XmlElement).getAttribute('Id') ?? '',
		),
	);
	let serial = 1;
	while (existingIds.has(`rId${serial}`)) serial++;
	const relationship = relsDoc.createElementNS(REL_NS, 'Relationship');
	relationship.setAttribute('Id', `rId${serial}`);
	relationship.setAttribute('Type', NUMBERING_RELATIONSHIP_TYPE);
	relationship.setAttribute('Target', 'numbering.xml');
	relsDoc.documentElement.appendChild(relationship);
	zip.file(relsPath, buildXml(relsDoc));
}
