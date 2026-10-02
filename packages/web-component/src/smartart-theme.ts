import { diagramColorTheme, resolveDiagramColor, type DocxDiagram } from '@christophervr/docx-core';
import type { ThemeCatalog } from '@christophervr/docx-core';

type Drawing = NonNullable<DocxDiagram['drawing']>;

/** Scheme colour names a cached drawing may reference (plus the mapped `tx1`/`bg1` aliases). */
const SCHEME_NAMES = [
	'dk1',
	'lt1',
	'dk2',
	'lt2',
	'accent1',
	'accent2',
	'accent3',
	'accent4',
	'accent5',
	'accent6',
	'hlink',
	'folHlink',
	'tx1',
	'bg1',
	'tx2',
	'bg2',
];

export interface ThemedDrawing {
	drawing: Drawing;
	/** Resolved theme colours for any scheme reference left in the drawing. */
	schemeColors: Record<string, string>;
	/** Colours whose transforms (lumMod, shade...) the resolver could not apply, by name. */
	unappliedTransforms: string[];
	/** The renderer draws opaque colours: some resolved colours carry transparency. */
	alphaDropped: boolean;
}

const isColor = (value: unknown): value is { kind: string; value: string; transforms: unknown } =>
	typeof value === 'object' &&
	value !== null &&
	typeof (value as { kind?: unknown }).kind === 'string' &&
	typeof (value as { value?: unknown }).value === 'string' &&
	Array.isArray((value as { transforms?: unknown }).transforms);

/**
 * Resolves every colour of a diagram's cached drawing against the document theme with
 * `resolveDiagramColor` (scheme references and their transforms become plain sRGB), so the shared
 * SVG renderer, which only knows base scheme colours, draws what the producing application
 * cached. Returns a copy; the model is not changed.
 */
export function themeDrawing(drawing: Drawing, theme: ThemeCatalog | undefined): ThemedDrawing {
	const unapplied = new Set<string>();
	let alphaDropped = false;
	const walk = (value: unknown): unknown => {
		if (Array.isArray(value)) return value.map(walk);
		if (typeof value !== 'object' || value === null) return value;
		if (isColor(value) && value.kind !== 'srgb') {
			const resolved = resolveDiagramColor(value as never, theme);
			if (resolved) {
				for (const name of resolved.unapplied) unapplied.add(name);
				if (resolved.alpha < 1) alphaDropped = true;
				return { kind: 'srgb', value: resolved.hex.replace(/^#/u, ''), transforms: [] };
			}
		}
		return Object.fromEntries(Object.entries(value).map(([key, entry]) => [key, walk(entry)]));
	};
	const colors = diagramColorTheme(theme);
	const schemeColors: Record<string, string> = {};
	for (const name of SCHEME_NAMES) {
		const hex = colors.scheme(name);
		if (hex) schemeColors[name] = hex;
	}
	return {
		drawing: walk(drawing) as Drawing,
		schemeColors,
		unappliedTransforms: [...unapplied],
		alphaDropped,
	};
}
