/** The keyboard shortcut help dialog: lists the shortcut registry plus the editor's own key bindings. */
import { EDITOR_BINDING_LABELS } from './binding-labels';
import { editorBindings } from './editor-commands';
import { formatKeys, type ShortcutRegistry } from './keyboard';
import { translate, type EditorLocale, type LocalizationKey } from './localization';

export interface ShortcutRow {
	keys: string;
	label: LocalizationKey;
}

/** Only bindings that exist in the editor's keymap are listed, grouped by action. */
export function editorShortcutRows(
	mac: boolean,
	bindings: Record<string, unknown> = editorBindings,
): ShortcutRow[] {
	const rows = new Map<LocalizationKey, string[]>();
	for (const [binding, label] of Object.entries(EDITOR_BINDING_LABELS)) {
		if (!(binding in bindings)) continue;
		const keys = formatKeys(binding.replaceAll('-', '+'), mac);
		rows.set(label, [...(rows.get(label) ?? []), keys]);
	}
	return [...rows].map(([label, keys]) => ({ label, keys: keys.join(' / ') }));
}

export function shortcutRows(registry: ShortcutRegistry, mac: boolean): ShortcutRow[] {
	const general = registry.shortcuts.map((shortcut) => ({
		label: shortcut.label,
		keys: registry.keysFor(shortcut.id, mac).join(' / '),
	}));
	return [...general, ...editorShortcutRows(mac)].filter((row) => row.keys);
}

export interface ShortcutHelp {
	element: HTMLElement;
	readonly isOpen: boolean;
	open(rows: ShortcutRow[], locale: EditorLocale): void;
	close(): void;
}

export function createShortcutHelp(onClose: () => void): ShortcutHelp {
	const element = document.createElement('section');
	element.className = 'dve-dialog dve-shortcut-help';
	element.setAttribute('role', 'dialog');
	element.setAttribute('aria-modal', 'true');
	element.hidden = true;
	const heading = document.createElement('h2');
	heading.id = 'dve-shortcut-help-title';
	element.setAttribute('aria-labelledby', heading.id);
	const note = document.createElement('p');
	note.className = 'dve-shortcut-note';
	const table = document.createElement('table');
	const closeButton = document.createElement('button');
	closeButton.type = 'button';
	closeButton.className = 'dve-dialog-primary';
	const actions = document.createElement('div');
	actions.className = 'dve-dialog-actions';
	actions.append(closeButton);
	element.append(heading, note, table, actions);

	const close = () => {
		if (element.hidden) return;
		element.hidden = true;
		onClose();
	};
	closeButton.addEventListener('click', close);
	element.addEventListener('keydown', (event) => {
		if (event.key === 'Escape') {
			event.preventDefault();
			close();
		} else if (event.key === 'Tab') {
			// The dialog is modal and has a single control: keep focus inside it.
			event.preventDefault();
			closeButton.focus();
		}
	});

	return {
		element,
		get isOpen() {
			return !element.hidden;
		},
		open(rows, locale) {
			heading.textContent = translate(locale, 'shortcut.dialogTitle');
			note.textContent = translate(locale, 'shortcut.dialogNote');
			closeButton.textContent = translate(locale, 'shortcut.close');
			table.replaceChildren();
			const head = table.createTHead().insertRow();
			for (const key of ['shortcut.columnAction', 'shortcut.columnKeys'] as const) {
				const cell = document.createElement('th');
				cell.scope = 'col';
				cell.textContent = translate(locale, key);
				head.append(cell);
			}
			const body = table.createTBody();
			for (const row of rows) {
				const tr = body.insertRow();
				tr.insertCell().textContent = translate(locale, row.label);
				const kbd = document.createElement('kbd');
				kbd.textContent = row.keys;
				tr.insertCell().append(kbd);
			}
			element.hidden = false;
			closeButton.focus();
		},
		close,
	};
}
