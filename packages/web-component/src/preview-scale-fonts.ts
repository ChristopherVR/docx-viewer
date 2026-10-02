import type { RunFormatting } from 'docx-core';
import { scaledSegments, scaleMeasurer } from './run-scale';

interface PreviewScale {
	element: WeakRef<HTMLElement>;
	text: string;
	formatting: RunFormatting;
	family: string | undefined;
}
const registries = new WeakMap<FontFaceSet, Set<PreviewScale>>();

/** Closed previews have no editor lifecycle. Keep only weak element references while fonts load. */
export function refreshPreviewScaleWithFonts(
	element: HTMLElement,
	text: string,
	formatting: RunFormatting,
	family: string | undefined,
): void {
	const fonts = element.ownerDocument.fonts;
	if (!fonts?.addEventListener) return;
	let entries = registries.get(fonts);
	if (!entries) {
		entries = new Set();
		registries.set(fonts, entries);
		const registry = entries;
		const refresh = () => {
			const measurer = scaleMeasurer();
			for (const entry of registry) {
				const target = entry.element.deref();
				if (!target?.isConnected) {
					registry.delete(entry);
					continue;
				}
				const segment = scaledSegments(entry.text, entry.formatting, entry.family, measurer)[0];
				if (segment) target.style.cssText = segment.css;
			}
		};
		fonts.addEventListener('loadingdone', refresh);
		void fonts.ready?.then(refresh);
	}
	entries.add({ element: new WeakRef(element), text, formatting, family });
}
