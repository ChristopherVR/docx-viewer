// Test-only: validates WordprocessingML parts the writer produces against the ECMA-376
// Transitional schemas (see schemas/ecma-376-transitional/README.md).
import { readdirSync, readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';
import { validateXML } from 'xmllint-wasm';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';

const SCHEMA_DIR = join(dirname(fileURLToPath(import.meta.url)), '../../schemas/ecma-376-transitional');

/** Namespaces the ECMA-376 Transitional schemas define; anything else is an extension. */
const STANDARD_NAMESPACES = new Set([
	'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
	'http://schemas.openxmlformats.org/officeDocument/2006/math',
	'http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes',
	'http://schemas.openxmlformats.org/drawingml/2006/main',
	'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
	'http://schemas.openxmlformats.org/drawingml/2006/picture',
	'http://schemas.openxmlformats.org/drawingml/2006/chart',
	'http://schemas.openxmlformats.org/drawingml/2006/diagram',
	'urn:schemas-microsoft-com:vml',
	'urn:schemas-microsoft-com:office:office',
	'urn:schemas-microsoft-com:office:word',
	'http://www.w3.org/XML/1998/namespace',
	'http://www.w3.org/2000/xmlns/',
]);
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

let schemaFiles: { fileName: string; contents: string }[] | undefined;
function schemas() {
	schemaFiles ??= readdirSync(SCHEMA_DIR)
		.filter((name) => name.endsWith('.xsd'))
		.map((fileName) => ({
			fileName,
			contents: readFileSync(join(SCHEMA_DIR, fileName), 'utf8')
				.replace(/schemaLocation="http:\/\/www\.w3\.org\/(?:2001|2009\/01)\/xml\.xsd"/g, 'schemaLocation="xml.xsd"')
				.replace(
					/<xsd:import namespace="http:\/\/www\.w3\.org\/XML\/1998\/namespace"\s*\/>/g,
					'<xsd:import namespace="http://www.w3.org/XML/1998/namespace" schemaLocation="xml.xsd"/>',
				),
		}));
	return schemaFiles;
}

/**
 * Applies Markup Compatibility processing as a consumer of only the ECMA-376 vocabulary does:
 * `mc:AlternateContent` becomes its `mc:Fallback` content, and elements or attributes in other
 * (extension) namespaces, such as Word 2010+'s `w14:`, are ignored.
 */
export function withoutExtensions(xml: string): string {
	const doc = new DOMParser().parseFromString(xml, 'text/xml');
	const visit = (element: Element) => {
		for (const attribute of Array.from(element.attributes)) {
			const ns = attribute.namespaceURI;
			if ((ns && !STANDARD_NAMESPACES.has(ns)) || (attribute.prefix === 'xmlns' && !STANDARD_NAMESPACES.has(attribute.value)))
				element.removeAttributeNode(attribute);
		}
		for (const child of Array.from(element.childNodes)) {
			if (child.nodeType !== 1) continue;
			const item = child as Element;
			if (item.namespaceURI === MC && item.localName === 'AlternateContent') {
				const fallback = Array.from(item.childNodes).find(
					(node) => node.nodeType === 1 && (node as Element).localName === 'Fallback',
				);
				for (const node of Array.from(fallback?.childNodes ?? [])) element.insertBefore(node, item);
				element.removeChild(item);
			} else if (!item.namespaceURI || !STANDARD_NAMESPACES.has(item.namespaceURI)) element.removeChild(item);
		}
		for (const child of Array.from(element.childNodes)) if (child.nodeType === 1) visit(child as Element);
	};
	visit(doc.documentElement as unknown as Element);
	return new XMLSerializer().serializeToString(doc);
}

/** Schema errors for one WordprocessingML part (empty when it is valid). */
export async function schemaErrors(xml: string): Promise<string[]> {
	const files = schemas();
	const result = await validateXML({
		xml: [{ fileName: 'part.xml', contents: withoutExtensions(xml) }],
		schema: files.find((file) => file.fileName === 'wml.xsd')!,
		preload: files.filter((file) => file.fileName !== 'wml.xsd'),
	});
	return result.valid ? [] : result.errors.map((error) => error.rawMessage);
}

/** Every WordprocessingML part in a saved package, with its schema errors. */
export async function packageSchemaErrors(bytes: Uint8Array): Promise<Record<string, string[]>> {
	const zip = await JSZip.loadAsync(bytes);
	const parts = Object.keys(zip.files).filter((name) =>
		/^word\/(document|styles|settings|numbering|comments|footnotes|endnotes|header\d*|footer\d*)\.xml$/.test(name),
	);
	const result: Record<string, string[]> = {};
	for (const name of parts) {
		const errors = await schemaErrors(await zip.file(name)!.async('string'));
		if (errors.length) result[name] = errors;
	}
	return result;
}
