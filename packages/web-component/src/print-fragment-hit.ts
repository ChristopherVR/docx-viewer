import type { LayoutFragment } from 'ooxml-core/docx/layout';

export interface PrintFragmentHit {
	element: HTMLElement;
	fragment: LayoutFragment;
}

const graphemes = new Intl.Segmenter(undefined, { granularity: 'grapheme' });

/** Measure the displayed glyph boxes, including CSS transforms, rather than the column width. */
function textOffset(element: HTMLElement, clientX: number): number | undefined {
	const text = element.firstChild;
	if (!text || text.nodeType !== Node.TEXT_NODE) return undefined;
	const range = element.ownerDocument.createRange();
	if (typeof range.getBoundingClientRect !== 'function') return undefined;
	const rtl = getComputedStyle(element).direction === 'rtl';
	let nearest: { distance: number; offset: number } | undefined;
	for (const segment of graphemes.segment(text.textContent ?? '')) {
		range.setStart(text, segment.index);
		range.setEnd(text, segment.index + segment.segment.length);
		const rect = range.getBoundingClientRect();
		if (!rect.width) continue;
		for (const [x, offset] of [
			[rtl ? rect.right : rect.left, segment.index],
			[rtl ? rect.left : rect.right, segment.index + segment.segment.length],
		] as const) {
			const distance = Math.abs(clientX - x);
			if (!nearest || distance < nearest.distance) nearest = { distance, offset };
		}
	}
	return nearest?.offset;
}

/** Returns undefined for old layout results that do not include fragment source ranges. */
export function resolveFragmentClick(
	fragments: PrintFragmentHit[],
	clientX: number,
): number | undefined {
	let nearest: { hit: PrintFragmentHit; distance: number } | undefined;
	for (const hit of fragments) {
		if (hit.fragment.sourceStart === undefined || hit.fragment.sourceEnd === undefined) continue;
		const rect = hit.element.getBoundingClientRect();
		const distance = Math.max(rect.left - clientX, clientX - rect.right, 0);
		if (!nearest || distance < nearest.distance) nearest = { hit, distance };
	}
	if (!nearest) return undefined;
	const { element, fragment } = nearest.hit;
	const start = fragment.sourceStart!;
	const end = fragment.sourceEnd!;
	if (start === end) return start;
	if (fragment.text.length === end - start && !fragment.leader && !fragment.object) {
		const measured = textOffset(element, clientX);
		if (measured !== undefined) return Math.min(end, start + measured);
	}
	const rect = element.getBoundingClientRect();
	const ratio = rect.width ? Math.min(1, Math.max(0, (clientX - rect.left) / rect.width)) : 0;
	if (fragment.text.length === end - start && !fragment.leader && !fragment.object) {
		const desired = (end - start) * ratio;
		const boundaries = [
			0,
			...[...graphemes.segment(fragment.text)].map((part) => part.index + part.segment.length),
		];
		return (
			start +
			boundaries.reduce(
				(nearest, boundary) =>
					Math.abs(boundary - desired) < Math.abs(nearest - desired) ? boundary : nearest,
				0,
			)
		);
	}
	return start + Math.round((end - start) * ratio);
}
