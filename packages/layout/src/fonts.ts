// Metric-compatible substitutes for Word's core fonts, so pagination measured on systems without
// Microsoft fonts (Linux, ChromeOS) lines up with Word. Carlito, Caladea, Arimo, Tinos and Cousine
// share advance widths with Calibri, Cambria, Arial, Times New Roman and Courier New.
import { DEFAULT_FONT_FAMILY } from './units.js';

const SUBSTITUTES: Record<string, { fallbacks: string[]; generic: string }> = {
	calibri: { fallbacks: ['Carlito'], generic: 'sans-serif' },
	'calibri light': { fallbacks: ['Carlito'], generic: 'sans-serif' },
	cambria: { fallbacks: ['Caladea'], generic: 'serif' },
	arial: { fallbacks: ['Arimo', 'Liberation Sans'], generic: 'sans-serif' },
	helvetica: { fallbacks: ['Arimo', 'Liberation Sans'], generic: 'sans-serif' },
	'times new roman': { fallbacks: ['Tinos', 'Liberation Serif'], generic: 'serif' },
	times: { fallbacks: ['Tinos', 'Liberation Serif'], generic: 'serif' },
	'courier new': { fallbacks: ['Cousine', 'Liberation Mono'], generic: 'monospace' },
	courier: { fallbacks: ['Cousine', 'Liberation Mono'], generic: 'monospace' },
	georgia: { fallbacks: ['Gelasio'], generic: 'serif' },
	consolas: { fallbacks: ['Inconsolata'], generic: 'monospace' },
};

const quote = (family: string) => `"${family.replace(/["\\]/g, '')}"`;

/** A CSS `font-family` list for a Word font name: the font, its metric-compatible substitutes, a generic. */
export function cssFontStack(family: string | undefined): string {
	const name = family?.trim() || DEFAULT_FONT_FAMILY;
	const known = SUBSTITUTES[name.toLowerCase()];
	return [quote(name), ...(known?.fallbacks ?? []).map(quote), known?.generic ?? 'sans-serif'].join(
		', ',
	);
}
