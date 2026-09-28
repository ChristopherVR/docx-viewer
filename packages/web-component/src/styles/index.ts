import { darkTheme, lightTheme } from '../theme/defaults';
import { themeToCssVars } from '../theme/css-vars';
import type { EditorTheme } from '../theme/types';
import aliases from './aliases.css?inline';
import base from './base.css?inline';
import ribbon from './ribbon.css?inline';
import documentCss from './document.css?inline';
import print from './print.css?inline';
import titlebar from './chrome-titlebar.css?inline';
import backstage from './backstage.css?inline';
import statusDialogs from './status-dialogs.css?inline';
import chromeDocument from './chrome-document.css?inline';
import pageNavigator from './page-navigator.css?inline';

const declarations = (theme: EditorTheme, scheme: 'light' | 'dark') =>
	Object.entries(themeToCssVars(theme))
		.map(([name, value]) => `${name}:${value};`)
		.join('') + `color-scheme:${scheme};`;

/**
 * Token presets are generated from theme/defaults.ts so the TypeScript presets and the shadow-root
 * CSS cannot drift. `theme="auto"` (the default) follows the OS; explicit values always win.
 */
export const themeTokenText = `:host{${declarations(lightTheme, 'light')}}
@media (prefers-color-scheme: dark){:host(:not([theme="light"])){${declarations(darkTheme, 'dark')}}}
:host([theme="dark"]){${declarations(darkTheme, 'dark')}}
:host([theme="light"]){${declarations(lightTheme, 'light')}}`;

export const editorStyleText = [
	themeTokenText,
	aliases,
	base,
	ribbon,
	documentCss,
	print,
	titlebar,
	backstage,
	statusDialogs,
	chromeDocument,
	pageNavigator,
].join('\n');
