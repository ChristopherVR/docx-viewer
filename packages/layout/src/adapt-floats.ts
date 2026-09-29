import type { Paragraph } from '@christophervr/docx-core';
import type { LayoutFloat } from './input.js';

/** Floating pictures anchored in a paragraph, with their `wp:positionH`/`wp:positionV` placement. */
export function paragraphFloats(paragraph: Paragraph): LayoutFloat[] {
	const floats: LayoutFloat[] = [];
	for (const run of paragraph.runs) {
		const image = run.image;
		if (!image?.anchored || !image.partName) continue;
		const placement = image.placement;
		floats.push({
			partName: image.partName,
			contentType: image.contentType,
			widthPx: image.widthPx,
			heightPx: image.heightPx,
			...(placement?.relativeFrom ? { relativeFromH: placement.relativeFrom } : {}),
			...(placement?.align ? { alignH: placement.align } : {}),
			...(placement?.offsetXPx !== undefined ? { offsetXPx: placement.offsetXPx } : {}),
			...(placement?.relativeFromV ? { relativeFromV: placement.relativeFromV } : {}),
			...(placement?.alignV ? { alignV: placement.alignV } : {}),
			...(placement?.offsetYPx !== undefined ? { offsetYPx: placement.offsetYPx } : {}),
			...(placement?.behindText ? { behindText: true } : {}),
			...(placement?.wrap ? { wrap: placement.wrap } : {}),
		});
	}
	return floats;
}
export function floatsOf(paragraph: Paragraph): { floats?: LayoutFloat[] } {
	const floats = paragraphFloats(paragraph);
	return floats.length ? { floats } : {};
}
