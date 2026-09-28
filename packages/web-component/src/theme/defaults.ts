import type { EditorTheme } from './types';

/**
 * Light preset: Word's neutral chrome with Word blue as primary. Neutrals follow the same
 * role split as pptx-viewer's light theme (background, card, muted, border).
 */
export const lightTheme: EditorTheme = {
	background: '#f3f2f1',
	foreground: '#1f1f1f',
	card: '#ffffff',
	cardForeground: '#1f1f1f',
	popover: '#ffffff',
	popoverForeground: '#1f1f1f',
	primary: '#185abd',
	primaryForeground: '#ffffff',
	secondary: '#eceae8',
	secondaryForeground: '#1f1f1f',
	muted: '#eceae8',
	mutedForeground: '#605e5c',
	accent: '#e6edf8',
	accentForeground: '#103f91',
	destructive: '#dc2626',
	destructiveForeground: '#ffffff',
	border: '#e1dfdd',
	input: '#c8c6c4',
	ring: '#185abd',
	radius: '4px',
};

/**
 * Dark preset: pptx-viewer's default dark gray scale (gray-950/900/800/700) with a lighter Word
 * blue as primary so it stays legible on dark surfaces.
 */
export const darkTheme: EditorTheme = {
	background: '#030712',
	foreground: '#f3f4f6',
	card: '#111827',
	cardForeground: '#f3f4f6',
	popover: '#111827',
	popoverForeground: '#f3f4f6',
	primary: '#4f8ef7',
	primaryForeground: '#ffffff',
	secondary: '#1f2937',
	secondaryForeground: '#f3f4f6',
	muted: '#1f2937',
	mutedForeground: '#9ca3af',
	accent: '#1e2f4f',
	accentForeground: '#f3f4f6',
	destructive: '#ef4444',
	destructiveForeground: '#ffffff',
	border: '#374151',
	input: '#374151',
	ring: '#4f8ef7',
	radius: '4px',
};
