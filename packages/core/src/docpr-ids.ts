// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Package-level allocator for `wp:docPr/@id`. Word expects drawing ids to be unique across the
// document, headers, footers and notes, so one allocator is shared by every part in a save.
import type JSZip from 'jszip';
import type { XmlDocument } from './xml.js';

const WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const DOC_PR_ID = /<(?:[\w.-]+:)?docPr\b[^>]*?\sid\s*=\s*["'](\d+)["']/g;

export class DocPrIdAllocator {
	private max = 0;

	reserve(id: number): void {
		if (Number.isSafeInteger(id) && id > this.max) this.max = id;
	}

	/** Reserves every `wp:docPr` id in an already parsed part. */
	reserveFromDocument(doc: XmlDocument): void {
		for (const node of Array.from(doc.getElementsByTagNameNS(WP_NS, 'docPr')))
			this.reserve(Number(node.getAttribute('id')));
	}

	/** Reserves every `docPr` id found in serialized XML. */
	reserveFromXml(xml: string): void {
		for (const match of xml.matchAll(DOC_PR_ID)) this.reserve(Number(match[1]));
	}

	/** Reserves the ids of every XML part in a package (document, headers, footers, notes...). */
	async reserveFromPackage(zip: JSZip): Promise<void> {
		for (const file of zip.file(/^word\/[^/]+\.xml$/))
			this.reserveFromXml(await file.async('string'));
	}

	next(): string {
		this.max += 1;
		return String(this.max);
	}
}
