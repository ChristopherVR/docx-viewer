// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { CharacterStyleDefinition, RunStyleCatalog } from './run-style-model.js';
import { first, getW, named, parseXml, type XmlDocument, type XmlElement } from './xml.js';
import { parseRunProperties } from './run-properties.js';

function styleElements(document: XmlDocument): XmlElement[] {
	return Array.from(document.getElementsByTagName('*')).filter(
		(element): element is XmlElement =>
			named(element, 'style') &&
			(getW(element, 'type') === 'character' || getW(element, 'type') === 'paragraph'),
	);
}

/** Parses `styles.xml`'s `rPrDefault` plus every character/paragraph style's `w:rPr`. */
export function parseRunStyleCatalog(xml: string): RunStyleCatalog {
	const document = parseXml(xml);
	const defaults = first(first(document.documentElement, 'docDefaults'), 'rPrDefault');
	const styles = Object.create(null) as Record<string, CharacterStyleDefinition>;
	for (const element of styleElements(document)) {
		const id = getW(element, 'styleId');
		if (!id) continue;
		const type = getW(element, 'type') === 'character' ? 'character' : 'paragraph';
		const basedOn = getW(first(element, 'basedOn'), 'val');
		const linkedStyle = getW(first(element, 'link'), 'val');
		const defaultValue = getW(element, 'default');
		const isDefault =
			defaultValue === undefined
				? undefined
				: !['0', 'false', 'off', 'no', 'none'].includes(defaultValue.toLowerCase());
		const name = getW(first(element, 'name'), 'val');
		styles[id] = {
			id,
			type,
			...(name ? { name } : {}),
			...(basedOn ? { basedOn } : {}),
			...(linkedStyle ? { linkedStyle } : {}),
			...(isDefault === undefined ? {} : { isDefault }),
			formatting: parseRunProperties(first(element, 'rPr')),
		};
	}
	const warnings: string[] = [];
	for (const style of Object.values(styles)) {
		const seen = new Set<string>([style.id]);
		let parent = style.basedOn;
		while (parent && styles[parent]) {
			if (seen.has(parent)) {
				warnings.push(
					`Character style inheritance cycle detected at "${parent}"; cyclic inheritance is ignored.`,
				);
				break;
			}
			seen.add(parent);
			parent = styles[parent].basedOn;
		}
	}
	return {
		docDefaults: parseRunProperties(first(defaults, 'rPr')),
		styles,
		warnings: [...new Set(warnings)],
	};
}
