import { cssFontStack } from './fonts.js';
/** A font as the measurer needs it; sizes are already in CSS pixels. */
export interface LayoutFontSpec {
	family: string;
	sizePx: number;
	bold?: boolean;
	italic?: boolean;
}

/** Injectable text measurer: canvas in the browser, a deterministic fake in tests. */
export interface TextMeasurer {
	/** Width of `text` set in `font`, in CSS pixels. Must be additive-ish (no kerning assumed). */
	widthOf(text: string, font: LayoutFontSpec): number;
	/** Natural single-spaced line height for `font`, in CSS pixels (ascent + descent + leading). */
	lineHeightOf(font: LayoutFontSpec): number;
	/** Height above the baseline within that line height; about 80% when not provided. */
	ascentOf?(font: LayoutFontSpec): number;
}

/** A font's ascent and descent (together its single line height), for baseline alignment. */
export function fontMetrics(
	measurer: TextMeasurer,
	font: LayoutFontSpec,
): { ascent: number; descent: number } {
	const height = measurer.lineHeightOf(font);
	const ascent = Math.min(height, measurer.ascentOf?.(font) ?? height * 0.8);
	return { ascent, descent: height - ascent };
}

/**
 * Deterministic measurer for unit tests: fixed per-character advance and line
 * height derived only from font size, so line-breaking and pagination tests
 * never depend on host font metrics or canvas availability.
 */
export function createFakeMeasurer(
	options: { charWidthFactor?: number; lineHeightFactor?: number } = {},
): TextMeasurer {
	const charWidthFactor = options.charWidthFactor ?? 0.6;
	const lineHeightFactor = options.lineHeightFactor ?? 1.15;
	return {
		widthOf(text, font) {
			let width = 0;
			for (const ch of text) {
				if (ch === '\t' || ch === '\n') continue;
				width += font.sizePx * charWidthFactor * (font.bold ? 1.08 : 1) * (ch === ' ' ? 0.6 : 1);
			}
			return width;
		},
		lineHeightOf(font) {
			return font.sizePx * lineHeightFactor;
		},
	};
}

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
			const key = `${fontString(font)}\u0000${text}`;
			const cached = widthCache.get(key);
			if (cached !== undefined) return cached;
			c.font = fontString(font);
			const width = c.measureText(text).width;
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
