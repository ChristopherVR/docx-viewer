import { EDITOR_THEME_MODES, type EditorTheme, type EditorThemeMode } from './types';
import { themeToCssVars } from './css-vars';

/** Unknown values fall back to `auto`, which follows the OS color scheme. */
export function normalizeThemeMode(value: unknown): EditorThemeMode {
	return EDITOR_THEME_MODES.includes(value as EditorThemeMode)
		? (value as EditorThemeMode)
		: 'auto';
}

/**
 * Apply custom-property overrides inline on the host so they beat the stylesheet's presets in
 * both light and dark. Returns the applied names so the next call can clear stale ones.
 */
export function applyThemeColors(
	host: HTMLElement,
	previous: readonly string[],
	colors: Partial<EditorTheme> | undefined,
): string[] {
	for (const name of previous) host.style.removeProperty(name);
	const vars = themeToCssVars(colors);
	for (const [name, value] of Object.entries(vars)) host.style.setProperty(name, value);
	return Object.keys(vars);
}
