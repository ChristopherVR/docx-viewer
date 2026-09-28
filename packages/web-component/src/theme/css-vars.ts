import type { EditorTheme } from './types';

/** Every token key; each becomes the kebab-case `--dve-*` custom property. */
export const THEME_KEYS = [
	'background',
	'foreground',
	'card',
	'cardForeground',
	'popover',
	'popoverForeground',
	'primary',
	'primaryForeground',
	'secondary',
	'secondaryForeground',
	'muted',
	'mutedForeground',
	'accent',
	'accentForeground',
	'destructive',
	'destructiveForeground',
	'border',
	'input',
	'ring',
	'radius',
] as const satisfies readonly (keyof EditorTheme)[];

const kebab = (key: string) => key.replace(/[A-Z]/g, (letter) => `-${letter.toLowerCase()}`);

/** Convert a (partial) theme to `--dve-*` custom properties. Unknown or empty values are skipped. */
export function themeToCssVars(theme: Partial<EditorTheme> | undefined): Record<string, string> {
	const vars: Record<string, string> = {};
	if (!theme) return vars;
	for (const key of THEME_KEYS) {
		const value = theme[key];
		if (typeof value === 'string' && value.trim()) vars[`--dve-${kebab(key)}`] = value.trim();
	}
	return vars;
}
