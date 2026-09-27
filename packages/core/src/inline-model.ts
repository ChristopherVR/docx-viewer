// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Inline pictures and hyperlink targets carried by text runs.

/** An inline drawing (`w:drawing` or legacy `w:pict`) modeled at run granularity. */
export interface InlineImage {
	/** Relationship id in `word/_rels/document.xml.rels` pointing at the media part. */
	relId: string;
	/** Package part name holding the image bytes, e.g. `word/media/image1.png`. */
	partName: string;
	contentType: string;
	/** An SVG original (`asvg:svgBlip`) shown instead of the raster `partName`, which is its PNG fallback. */
	svgPartName?: string;
	widthPx: number;
	heightPx: number;
	/** From `wp:docPr/@descr`. */
	altText?: string;
	/** From `wp:docPr/@title` or `@name`. */
	title?: string;
	/** `wp:anchor` (floating) drawings render as sized inline placeholders; wrapping/position is lost. */
	anchored?: boolean;
	/** Floating (`wp:anchor`) wrapping and horizontal position, read-only; the anchor XML is preserved. */
	placement?: PicturePlacement;
	/** Set for non-picture drawings (chart, SmartArt, shape, unresolved legacy VML): rendered as a labeled placeholder with no editable bytes. */
	unsupported?: string;
}
/** How a floating picture sits relative to text. */
export interface PicturePlacement {
	wrap: 'square' | 'tight' | 'through' | 'topAndBottom' | 'none';
	/** `behindDoc`: with `wrap: 'none'`, the picture is behind rather than in front of text. */
	behindText?: boolean;
	/** `wp:positionH/wp:align`, when the picture is aligned rather than offset. */
	align?: 'left' | 'center' | 'right' | 'inside' | 'outside';
	/** `wp:positionH/wp:posOffset` in CSS pixels, relative to `relativeFrom`. */
	offsetXPx?: number;
	relativeFrom?: string;
}
/** A `w:hyperlink` target, resolved from its relationship (external) or `w:anchor` (internal). */
export interface HyperlinkInfo {
	href?: string;
	anchor?: string;
	tooltip?: string;
}
