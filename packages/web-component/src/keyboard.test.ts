// @vitest-environment jsdom
import { describe, expect, it, vi } from 'vitest';
import {
	createShortcutRegistry,
	formatKeys,
	installKeyboard,
	matchesKeys,
	type KeyEventLike,
	type ShortcutActions,
} from './keyboard';
import { editorShortcutRows, shortcutRows } from './shortcut-help';
import { translate } from './localization';

const key = (name: string, modifiers: Partial<KeyEventLike> = {}): KeyEventLike => ({
	key: name,
	ctrlKey: false,
	metaKey: false,
	altKey: false,
	shiftKey: false,
	...modifiers,
});

function actions() {
	const spies = {
		find: vi.fn(),
		replace: vi.fn(),
		save: vi.fn(),
		print: vi.fn(),
		moveRegion: vi.fn(),
		focusRibbon: vi.fn(),
		contextMenu: vi.fn(),
		escape: vi.fn(() => true),
		help: vi.fn(),
	} satisfies ShortcutActions;
	return spies;
}

describe('shortcut key matching', () => {
	it('maps Mod to Ctrl on PC and Cmd on Mac, never both', () => {
		expect(matchesKeys(key('f', { ctrlKey: true }), 'Mod+F', false)).toBe(true);
		expect(matchesKeys(key('f', { metaKey: true }), 'Mod+F', false)).toBe(false);
		expect(matchesKeys(key('f', { metaKey: true }), 'Mod+F', true)).toBe(true);
		expect(matchesKeys(key('f', { ctrlKey: true }), 'Mod+F', true)).toBe(false);
		expect(matchesKeys(key('f', { ctrlKey: true, metaKey: true }), 'Mod+F', true)).toBe(false);
	});

	it('requires the exact modifier set', () => {
		expect(matchesKeys(key('F6'), 'F6', false)).toBe(true);
		expect(matchesKeys(key('F6', { shiftKey: true }), 'F6', false)).toBe(false);
		expect(matchesKeys(key('F6', { shiftKey: true }), 'Shift+F6', false)).toBe(true);
		expect(matchesKeys(key('f', { ctrlKey: true, altKey: true }), 'Mod+F', false)).toBe(false);
		expect(matchesKeys(key('F', { ctrlKey: true }), 'Mod+F', false)).toBe(true);
		expect(matchesKeys(key('/', { ctrlKey: true }), 'Mod+/', false)).toBe(true);
	});

	it('formats key legends per platform', () => {
		expect(formatKeys('Mod+Shift+Z', false)).toBe('Ctrl+Shift+Z');
		expect(formatKeys('Mod+F', true)).toBe('⌘F');
		expect(formatKeys('Escape', false)).toBe('Esc');
	});
});

describe('shortcut registry', () => {
	const pc = { readOnly: false, mac: false };

	it('dispatches every documented shortcut to its command', () => {
		const spies = actions();
		const registry = createShortcutRegistry(spies);
		const run = (event: KeyEventLike) => registry.dispatch(event, pc);
		expect(run(key('f', { ctrlKey: true }))).toBe('ran');
		expect(run(key('h', { ctrlKey: true }))).toBe('ran');
		expect(run(key('s', { ctrlKey: true }))).toBe('ran');
		expect(run(key('p', { ctrlKey: true }))).toBe('ran');
		run(key('F6'));
		run(key('F6', { shiftKey: true }));
		run(key('F10'));
		run(key('F10', { shiftKey: true }));
		run(key('ContextMenu'));
		run(key('Escape'));
		run(key('F1'));
		run(key('/', { ctrlKey: true }));
		expect(spies.find).toHaveBeenCalledOnce();
		expect(spies.replace).toHaveBeenCalledOnce();
		expect(spies.save).toHaveBeenCalledOnce();
		expect(spies.print).toHaveBeenCalledOnce();
		expect(spies.moveRegion.mock.calls).toEqual([[1], [-1]]);
		expect(spies.focusRibbon).toHaveBeenCalledOnce();
		expect(spies.contextMenu).toHaveBeenCalledTimes(2);
		expect(spies.escape).toHaveBeenCalledOnce();
		expect(spies.help).toHaveBeenCalledTimes(2);
	});

	it('uses Cmd on Mac and ignores Ctrl there', () => {
		const spies = actions();
		const registry = createShortcutRegistry(spies);
		const mac = { readOnly: false, mac: true };
		expect(registry.dispatch(key('s', { ctrlKey: true }), mac)).toBe('ignored');
		expect(registry.dispatch(key('s', { metaKey: true }), mac)).toBe('ran');
		expect(spies.save).toHaveBeenCalledOnce();
	});

	it('blocks editing shortcuts when read-only but keeps navigation ones', () => {
		const spies = actions();
		const registry = createShortcutRegistry(spies);
		const readOnly = { readOnly: true, mac: false };
		expect(registry.dispatch(key('h', { ctrlKey: true }), readOnly)).toBe('blocked');
		expect(spies.replace).not.toHaveBeenCalled();
		expect(registry.dispatch(key('f', { ctrlKey: true }), readOnly)).toBe('ran');
		expect(registry.dispatch(key('s', { ctrlKey: true }), readOnly)).toBe('ran');
		expect(spies.find).toHaveBeenCalledOnce();
	});

	it('reports a command that declines as ignored so the key keeps its default', () => {
		const spies = actions();
		spies.escape.mockReturnValue(false as never);
		const registry = createShortcutRegistry(spies);
		expect(registry.dispatch(key('Escape'), pc)).toBe('ignored');
	});

	it('treats a lone Alt tap as focusing the ribbon on PC only', () => {
		const spies = actions();
		const registry = createShortcutRegistry(spies);
		expect(registry.dispatchAltTap({ readOnly: false, mac: true })).toBe('ignored');
		expect(registry.dispatchAltTap(pc)).toBe('ran');
		expect(spies.focusRibbon).toHaveBeenCalledOnce();
		expect(registry.dispatch(key('Alt', { altKey: true }), pc)).toBe('ignored');
	});

	it('gives every shortcut a translated label in both locales', () => {
		const registry = createShortcutRegistry(actions());
		for (const shortcut of registry.shortcuts) {
			expect(translate('en', shortcut.label)).not.toBe(shortcut.label);
			expect(translate('fr', shortcut.label)).not.toBe(translate('en', shortcut.label));
		}
	});
});

describe('keyboard listener', () => {
	it('prevents the default of handled keys but leaves keys ProseMirror already consumed', () => {
		const host = document.createElement('div');
		const inner = host.appendChild(document.createElement('div'));
		const spies = actions();
		installKeyboard(host, createShortcutRegistry(spies), () => ({ readOnly: false, mac: false }));
		const press = (init: KeyboardEventInit) => {
			const event = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ...init });
			inner.dispatchEvent(event);
			return event;
		};
		expect(press({ key: 'f', ctrlKey: true }).defaultPrevented).toBe(true);
		expect(spies.find).toHaveBeenCalledOnce();
		inner.addEventListener('keydown', (event) => event.preventDefault());
		press({ key: 's', ctrlKey: true });
		expect(spies.save).not.toHaveBeenCalled();
	});

	it('runs the ribbon shortcut on Alt keyup only when no other key intervened', () => {
		const host = document.createElement('div');
		const spies = actions();
		installKeyboard(host, createShortcutRegistry(spies), () => ({ readOnly: false, mac: false }));
		const send = (type: string, init: KeyboardEventInit) =>
			host.dispatchEvent(new KeyboardEvent(type, { bubbles: true, cancelable: true, ...init }));
		send('keydown', { key: 'Alt', altKey: true });
		send('keydown', { key: 'ArrowLeft', altKey: true });
		send('keyup', { key: 'Alt' });
		expect(spies.focusRibbon).not.toHaveBeenCalled();
		send('keydown', { key: 'Alt', altKey: true });
		send('keyup', { key: 'Alt' });
		expect(spies.focusRibbon).toHaveBeenCalledOnce();
	});
});

describe('shortcut help listing', () => {
	it('lists the registry and only the editor bindings that exist', () => {
		const registry = createShortcutRegistry(actions());
		const rows = shortcutRows(registry, false);
		expect(rows.find((row) => row.label === 'shortcut.find')?.keys).toBe('Ctrl+F');
		expect(rows.find((row) => row.label === 'shortcut.help')?.keys).toBe('Ctrl+/ / F1');
		expect(rows.find((row) => row.label === 'Bold')?.keys).toBe('Ctrl+B');
		expect(rows.find((row) => row.label === 'Redo')?.keys).toBe('Ctrl+Y / Ctrl+Shift+Z');
		expect(editorShortcutRows(false, { 'Mod-b': () => true })).toEqual([
			{ label: 'Bold', keys: 'Ctrl+B' },
		]);
	});
});
