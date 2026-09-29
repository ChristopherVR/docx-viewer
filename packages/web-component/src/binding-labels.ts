import { formatKeys, isMacPlatform } from './keyboard';
import type { LocalizationKey } from './localization';

/**
 * Labels for the editor's ProseMirror key bindings. The help dialog lists exactly the bindings that
 * have a label here, and ribbon tooltips append the same keys ("Bold (Ctrl+B)"), so a shortcut is
 * documented in one place. A label that is also a ribbon control's English name links the two.
 */
export const EDITOR_BINDING_LABELS: Record<string, LocalizationKey> = {
	'Mod-z': 'Undo',
	'Mod-y': 'Redo',
	'Mod-Shift-z': 'Redo',
	'Mod-b': 'Bold',
	'Mod-i': 'Italic',
	'Mod-u': 'Underline',
	'Mod-l': 'Align left',
	'Mod-e': 'Align center',
	'Mod-r': 'Align right',
	'Mod-j': 'Justify',
	'Mod-Shift-.': 'Grow font',
	'Mod-Shift-,': 'Shrink font',
	'Mod-=': 'Subscript',
	'Mod-Shift-=': 'Superscript',
	'Mod-Space': 'Clear formatting',
	'Shift-Enter': 'shortcut.lineBreak',
	Tab: 'Increase list level',
	'Shift-Tab': 'Decrease list level',
	'Mod-Enter': 'Insert page break',
};

/** Extra shortcuts that live outside the ProseMirror keymap, by ribbon control label. */
const EXTERNAL_HINTS: Record<string, string> = {
	'Find and replace': 'Mod+F',
	Replace: 'Mod+H',
	'Insert link': 'Mod+K',
	Print: 'Mod+P',
	'Select all': 'Mod+A',
	Copy: 'Mod+C',
	Cut: 'Mod+X',
	Paste: 'Mod+V',
};

/** The shortcut for a ribbon control's English label, formatted for the platform, if it has one. */
export function shortcutHint(label: string, mac = isMacPlatform()): string | undefined {
	const external = EXTERNAL_HINTS[label];
	if (external) return formatKeys(external, mac);
	const binding = Object.entries(EDITOR_BINDING_LABELS).find(([, name]) => name === label)?.[0];
	return binding ? formatKeys(binding.replaceAll('-', '+'), mac) : undefined;
}
