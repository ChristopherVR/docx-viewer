// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Bookmarks are modeled read-only at paragraph granularity: the model only exposes which
// names start in a paragraph, not their exact character range within it.
import { children, getW, isElement, named, type XmlElement } from './xml.js';

/** Bookmark names starting directly inside this paragraph (`w:bookmarkStart/@w:name`, direct children only). */
export function paragraphBookmarkNames(node: XmlElement): string[] {
	return children(node, 'bookmarkStart')
		.map((el) => getW(el, 'name'))
		.filter((name): name is string => Boolean(name));
}

/** Clones a paragraph's direct `w:bookmarkStart`/`w:bookmarkEnd` children so they can be re-inserted after a rewrite. */
export function captureBookmarkNodes(node: XmlElement): {
	starts: XmlElement[];
	ends: XmlElement[];
} {
	const starts: XmlElement[] = [];
	const ends: XmlElement[] = [];
	for (const child of Array.from(node.childNodes)) {
		if (!isElement(child)) continue;
		if (named(child, 'bookmarkStart')) starts.push(child.cloneNode(true) as XmlElement);
		else if (named(child, 'bookmarkEnd')) ends.push(child.cloneNode(true) as XmlElement);
	}
	return { starts, ends };
}
