import { darkTheme, lightTheme } from '../theme/defaults';
import { themeToCssVars } from '../theme/css-vars';
import type { EditorTheme } from '../theme/types';
import aliases from './aliases.css?inline';
import base from './base.css?inline';
import ribbon from './ribbon.css?inline';
import ribbonButtons from './ribbon-buttons.css?inline';
import ribbonGallery from './ribbon-gallery.css?inline';
import ribbonWidgets from './ribbon-widgets.css?inline';
import ribbonStylesGroup from './ribbon-styles-group.css?inline';
import ribbonPopups from './ribbon-popups.css?inline';
import ribbonPolish from './ribbon-polish.css?inline';
import documentCss from './document.css?inline';
import print from './print.css?inline';
import titlebar from './chrome-titlebar.css?inline';
import backstage from './backstage.css?inline';
import statusDialogs from './status-dialogs.css?inline';
import chromeDocument from './chrome-document.css?inline';
import pageNavigator from './page-navigator.css?inline';
import keyboardMenu from './keyboard-menu.css?inline';
import smartArt from './smartart.css?inline';

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
	ribbonButtons,
	ribbonGallery,
	ribbonWidgets,
	ribbonStylesGroup,
	ribbonPopups,
	ribbonPolish,
	documentCss,
	print,
	titlebar,
	backstage,
	statusDialogs,
	chromeDocument,
	pageNavigator,
	keyboardMenu,
	smartArt,
].join('\n');
