/**
 * Editor theme tokens.
 *
 * Provenance: the semantic key set mirrors the shared theme of pptx-viewer
 * (packages/shared/src/theme/types.ts, `ViewerThemeColors` plus `radius`), which follows the
 * shadcn/ui naming convention. Keys are camelCase here and become `--dve-*` custom properties
 * (kebab-case). Values accept any CSS color. Only editor chrome is themed: the document paper
 * stays white like Word's page.
 */
export interface EditorTheme {
	/** Application chrome background (title bar, window). */
	background: string;
	foreground: string;
	/** Ribbon, panel and status-bar surface. */
	card: string;
	cardForeground: string;
	/** Menus, dialogs and other floating surfaces. */
	popover: string;
	popoverForeground: string;
	/** Brand and primary action color. */
	primary: string;
	primaryForeground: string;
	secondary: string;
	secondaryForeground: string;
	/** Muted surface; also the area behind the document pages. */
	muted: string;
	mutedForeground: string;
	/** Hover and selected highlight. */
	accent: string;
	accentForeground: string;
	destructive: string;
	destructiveForeground: string;
	border: string;
	input: string;
	/** Focus ring. */
	ring: string;
	/** Base corner radius, e.g. `4px`. */
	radius: string;
}
export type EditorThemeMode = 'light' | 'dark' | 'auto';
export const EDITOR_THEME_MODES: readonly EditorThemeMode[] = ['light', 'dark', 'auto'];
