// Test-only: validates WordprocessingML parts the writer produces against the ECMA-376
// Transitional schemas (see schemas/ecma-376-transitional/README.md).
import { readdirSync, readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';
import { validateXML } from 'xmllint-wasm';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';

const SCHEMA_DIR = join(
	dirname(fileURLToPath(import.meta.url)),
	'../../schemas/ecma-376-transitional',
);

const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

let schemaFiles: { fileName: string; contents: string }[] | undefined;
function schemas() {
	schemaFiles ??= readdirSync(SCHEMA_DIR)
		.filter((name) => name.endsWith('.xsd'))
		.map((fileName) => ({
			fileName,
			contents: readFileSync(join(SCHEMA_DIR, fileName), 'utf8')
				.replace(
					/schemaLocation="http:\/\/www\.w3\.org\/(?:2001|2009\/01)\/xml\.xsd"/g,
					'schemaLocation="xml.xsd"',
				)
				.replace(
					/<xsd:import namespace="http:\/\/www\.w3\.org\/XML\/1998\/namespace"\s*\/>/g,
					'<xsd:import namespace="http://www.w3.org/XML/1998/namespace" schemaLocation="xml.xsd"/>',
				),
		}));
	return schemaFiles;
}

export interface PreprocessResult {
	xml: string;
	/** Markup Compatibility violations found while preprocessing (these fail the part). */
	errors: string[];
}

/**
 * Applies Markup Compatibility processing the way a consumer that does not understand an
 * extension namespace must: only elements/attributes whose namespace prefix is declared in
 * `mc:Ignorable` are removed, and `mc:AlternateContent` becomes its `mc:Fallback` content (an
 * error is recorded when there is none). Anything else that is not in the schema is left in
 * place so validation reports it.
 */
export function withoutExtensions(xml: string): PreprocessResult {
	const doc = new DOMParser().parseFromString(xml, 'text/xml');
	const root = doc.documentElement as unknown as Element;
	const errors: string[] = [];
	const ignorable = new Set<string>();
	const declared = root.getAttributeNS(MC, 'Ignorable');
	for (const prefix of declared?.split(/\s+/).filter(Boolean) ?? []) {
		const ns = root.lookupNamespaceURI(prefix);
		if (ns) ignorable.add(ns);
		else errors.push(`mc:Ignorable names prefix "${prefix}" which is not declared`);
	}
	const visit = (element: Element) => {
		for (const attribute of Array.from(element.attributes)) {
			const ns = attribute.namespaceURI;
			const isIgnorableDeclaration = attribute.prefix === 'xmlns' && ignorable.has(attribute.value);
			const isIgnorableAttribute = ns !== null && ignorable.has(ns);
			const isIgnorableList = ns === MC && attribute.localName === 'Ignorable';
			if (isIgnorableDeclaration || isIgnorableAttribute || isIgnorableList)
				element.removeAttributeNode(attribute);
		}
		for (const child of Array.from(element.childNodes)) {
			if (child.nodeType !== 1) continue;
			const item = child as Element;
			if (item.namespaceURI === MC && item.localName === 'AlternateContent') {
				const fallback = Array.from(item.childNodes).find(
					(node) =>
						node.nodeType === 1 &&
						(node as Element).namespaceURI === MC &&
						(node as Element).localName === 'Fallback',
				);
				if (!fallback) errors.push('mc:AlternateContent has no mc:Fallback');
				for (const node of Array.from(fallback?.childNodes ?? [])) element.insertBefore(node, item);
				element.removeChild(item);
			} else if (item.namespaceURI && ignorable.has(item.namespaceURI)) element.removeChild(item);
		}
		for (const child of Array.from(element.childNodes))
			if (child.nodeType === 1) visit(child as Element);
	};
	visit(root);
	return { xml: new XMLSerializer().serializeToString(doc), errors };
}

/** Schema errors for one WordprocessingML part (empty when it is valid). */
export async function schemaErrors(xml: string): Promise<string[]> {
	const files = schemas();
	const processed = withoutExtensions(xml);
	const result = await validateXML({
		xml: [{ fileName: 'part.xml', contents: processed.xml }],
		schema: files.find((file) => file.fileName === 'wml.xsd')!,
		preload: files.filter((file) => file.fileName !== 'wml.xsd'),
	});
	return [
		...processed.errors,
		...(result.valid ? [] : result.errors.map((error) => error.rawMessage)),
	];
}

export interface PackageValidation {
	/** Names of the parts that were checked, so a test can prove a part was not skipped. */
	validated: string[];
	/** Schema errors keyed by part name; empty for a valid package. */
	errors: Record<string, string[]>;
}

/** Every WordprocessingML part in a saved package, with its schema errors. */
export async function packageSchemaErrors(bytes: Uint8Array): Promise<PackageValidation> {
	const zip = await JSZip.loadAsync(bytes);
	const parts = Object.keys(zip.files)
		.filter((name) =>
			/^word\/(document|styles|settings|numbering|comments|footnotes|endnotes|header\d*|footer\d*)\.xml$/.test(
				name,
			),
		)
		.sort();
	const errors: Record<string, string[]> = {};
	for (const name of parts) {
		const found = await schemaErrors(await zip.file(name)!.async('string'));
		if (found.length) errors[name] = found;
	}
	return { validated: parts, errors };
}
