import {
	createFakeMeasurer,
	cssFontStack,
	ligatureCss,
	type LayoutFontSpec,
	type TextMeasurer,
} from '@christophervr/ooxml-core/docx/layout';

/**
 * Browser measurer backed by an offscreen `<canvas>` 2D context. Falls back to
 * the fake measurer's formula if no canvas is available (SSR, jsdom without
 * canvas support), so callers never need to branch on environment.
 */
export function createCanvasMeasurer(): TextMeasurer {
	let context: CanvasRenderingContext2D | null | undefined;
	const widthCache = new Map<string, number>();
	const fallback = createFakeMeasurer();
	const ctx = (): CanvasRenderingContext2D | null => {
		if (context !== undefined) return context;
		try {
			context =
				typeof document === 'undefined' ? null : document.createElement('canvas').getContext('2d');
		} catch {
			context = null;
		}
		return context;
	};
	const fontString = (font: LayoutFontSpec) =>
		`${font.italic ? 'italic ' : ''}${font.bold ? 'bold ' : ''}${font.sizePx}px ${cssFontStack(font.family)}`;
	return {
		widthOf(text, font) {
			const c = ctx();
			if (!c) return fallback.widthOf(text, font);
			const key = `${fontString(font)}\u0000${font.kerning ?? 'auto'}\u0000${font.ligatures ?? 'auto'}\u0000${font.smallCaps ?? false}\u0000${text}`;
			const cached = widthCache.get(key);
			if (cached !== undefined) return cached;
			c.font = fontString(font);
			c.fontKerning = font.kerning ?? 'auto';
			let width = c.measureText(text).width;
			// Canvas has no ligature feature selector. Measure explicit selections with the same
			// browser shaping used to draw Print Layout, before applying glyph scale and spacing.
			if ((font.ligatures !== undefined || font.smallCaps) && document.body) {
				const span = document.createElement('span');
				span.style.cssText =
					'position:fixed;left:-100000px;top:0;visibility:hidden;display:inline-block;white-space:pre;padding:0;border:0;margin:0;letter-spacing:0;text-transform:none';
				span.style.font = fontString(font);
				span.style.fontKerning = font.kerning ?? 'auto';
				if (font.ligatures !== undefined)
					span.style.fontVariantLigatures = ligatureCss(font.ligatures);
				if (font.smallCaps) span.style.fontVariantCaps = 'small-caps';
				span.textContent = text;
				document.body.append(span);
				try {
					width = span.getBoundingClientRect().width;
				} finally {
					span.remove();
				}
			}
			if (widthCache.size > 20000) widthCache.clear();
			widthCache.set(key, width);
			return width;
		},
		ascentOf(font) {
			const c = ctx();
			if (!c) return fallback.lineHeightOf(font) * 0.8;
			c.font = fontString(font);
			return c.measureText('Mg').fontBoundingBoxAscent ?? font.sizePx * 0.95;
		},
		lineHeightOf(font) {
			const c = ctx();
			if (!c) return fallback.lineHeightOf(font);
			c.font = fontString(font);
			const metrics = c.measureText('Mg');
			// Word's single line is the font's ascent plus descent; paragraph line spacing multiplies it.
			const ascent = metrics.fontBoundingBoxAscent;
			const descent = metrics.fontBoundingBoxDescent;
			return ascent !== undefined && descent !== undefined ? ascent + descent : font.sizePx * 1.22;
		},
	};
}
