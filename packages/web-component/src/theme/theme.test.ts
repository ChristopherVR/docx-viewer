// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { DocxEditorElement, registerDocxEditor } from '../index';
import { themeTokenText } from '../styles';
import { darkTheme, lightTheme } from './defaults';
import { THEME_KEYS, themeToCssVars } from './css-vars';

describe('theme tokens', () => {
	it('mirrors pptx-viewer semantics: 19 color tokens plus radius, set in both presets', () => {
		expect(THEME_KEYS).toHaveLength(20);
		for (const preset of [lightTheme, darkTheme])
			for (const key of THEME_KEYS) expect(preset[key], key).toBeTruthy();
	});

	it('keeps Word blue as the light primary and distinct surfaces in dark', () => {
		expect(lightTheme.primary).toBe('#185abd');
		expect(darkTheme.background).not.toBe(lightTheme.background);
		expect(darkTheme.card).not.toBe(lightTheme.card);
	});

	it('emits kebab-case --dve-* custom properties and skips empty values', () => {
		const vars = themeToCssVars({ cardForeground: ' #111 ', ring: '', radius: '8px' });
		expect(vars).toEqual({ '--dve-card-foreground': '#111', '--dve-radius': '8px' });
		expect(themeToCssVars(undefined)).toEqual({});
		expect(Object.keys(themeToCssVars(lightTheme))).toHaveLength(THEME_KEYS.length);
	});

	it('generates the shadow-root token CSS from the presets for light, dark and auto', () => {
		expect(themeTokenText).toContain(`--dve-background:${lightTheme.background}`);
		expect(themeTokenText).toContain(
			`:host([theme="dark"]){--dve-background:${darkTheme.background}`,
		);
		expect(themeTokenText).toContain(
			'@media (prefers-color-scheme: dark){:host(:not([theme="light"]))',
		);
	});
});

describe('theme element API', () => {
	it('reflects theme to the attribute, defaults to auto and rejects unknown values', () => {
		registerDocxEditor();
		const editor = document.createElement('docx-editor') as DocxEditorElement;
		expect(editor.theme).toBe('auto');
		editor.theme = 'dark';
		expect(editor.getAttribute('theme')).toBe('dark');
		editor.setAttribute('theme', 'light');
		expect(editor.theme).toBe('light');
		editor.setAttribute('theme', 'sepia');
		expect(editor.theme).toBe('auto');
	});

	it('applies themeColors inline and clears stale overrides', () => {
		registerDocxEditor();
		const editor = document.createElement('docx-editor') as DocxEditorElement;
		editor.themeColors = { primary: '#ff0000', radius: '2px' };
		expect(editor.style.getPropertyValue('--dve-primary')).toBe('#ff0000');
		editor.themeColors = { ring: '#00ff00' };
		expect(editor.style.getPropertyValue('--dve-primary')).toBe('');
		expect(editor.style.getPropertyValue('--dve-ring')).toBe('#00ff00');
		editor.themeColors = undefined;
		expect(editor.style.getPropertyValue('--dve-ring')).toBe('');
	});
});
