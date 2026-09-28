/**
 * Central keyboard shortcut registry. Each shortcut has a stable id, one or more key specs, a
 * localized label and a command. The registry only decides *which* shortcut an event triggers;
 * the commands themselves dispatch through the editor's existing command runners
 * (see keyboard-actions.ts). ProseMirror's own formatting keymap is not duplicated here.
 */
import type { LocalizationKey } from './localization';

/**
 * Key spec grammar: modifiers joined with `+` then a key name, e.g. `Mod+F`, `Shift+F6`, `Escape`.
 * `Mod` is Cmd on Apple platforms and Ctrl elsewhere. A bare `Alt` means "tap Alt on its own".
 */
export type KeySpec = string;

export interface KeyEventLike {
	key: string;
	ctrlKey: boolean;
	metaKey: boolean;
	altKey: boolean;
	shiftKey: boolean;
}

export interface ShortcutContext {
	readOnly: boolean;
	mac: boolean;
}

export type ShortcutId =
	| 'find'
	| 'replace'
	| 'save'
	| 'print'
	| 'next-region'
	| 'previous-region'
	| 'focus-ribbon'
	| 'context-menu'
	| 'escape'
	| 'help';

export interface Shortcut {
	id: ShortcutId;
	keys: readonly KeySpec[];
	label: LocalizationKey;
	/** Editing shortcuts are blocked while the editor is read-only. */
	editing: boolean;
	/** Return `false` to report the event as not handled (it is then left untouched). */
	run(): boolean | void;
}

/** What each shortcut does; supplied by the editor so this module stays free of editor state. */
export interface ShortcutActions {
	find(): void;
	replace(): void;
	save(): void;
	print(): void;
	moveRegion(direction: 1 | -1): void;
	focusRibbon(): void;
	contextMenu(): void;
	escape(): boolean;
	help(): void;
}

interface ParsedKey {
	mod: boolean;
	shift: boolean;
	alt: boolean;
	key: string;
	tap: boolean;
}

const parsed = new Map<KeySpec, ParsedKey>();

function parse(spec: KeySpec): ParsedKey {
	let result = parsed.get(spec);
	if (result) return result;
	const parts = spec.split('+');
	const key = (parts.pop() ?? '').toLowerCase();
	const modifiers = new Set(parts.map((part) => part.toLowerCase()));
	result = {
		mod: modifiers.has('mod'),
		shift: modifiers.has('shift'),
		alt: modifiers.has('alt'),
		key,
		tap: key === 'alt' && !modifiers.size,
	};
	parsed.set(spec, result);
	return result;
}

/** Whether the platform uses Cmd (not Ctrl) as the primary modifier. */
export function isMacPlatform(): boolean {
	if (typeof navigator === 'undefined') return false;
	const platform =
		(navigator as Navigator & { userAgentData?: { platform?: string } }).userAgentData?.platform ??
		navigator.platform ??
		'';
	return /mac|iphone|ipad|ipod/i.test(platform);
}

/** Exact modifier match: Ctrl on a Mac, or Cmd elsewhere, never satisfies `Mod`. */
export function matchesKeys(event: KeyEventLike, spec: KeySpec, mac: boolean): boolean {
	const wanted = parse(spec);
	if (wanted.tap) return false;
	const mod = mac ? event.metaKey : event.ctrlKey;
	const foreign = mac ? event.ctrlKey : event.metaKey;
	return (
		mod === wanted.mod &&
		!foreign &&
		event.shiftKey === wanted.shift &&
		event.altKey === wanted.alt &&
		event.key.toLowerCase() === wanted.key
	);
}

const DISPLAY: Record<string, string> = { escape: 'Esc', contextmenu: 'Menu', ' ': 'Space' };

/** Human-readable key text, e.g. `Ctrl+F` or `⌘F`. Key names are not localized (they are legends). */
export function formatKeys(spec: KeySpec, mac: boolean): string {
	const parts = spec.split('+').map((part) => {
		const lower = part.toLowerCase();
		if (lower === 'mod') return mac ? '⌘' : 'Ctrl';
		if (lower === 'shift') return mac ? '⇧' : 'Shift';
		if (lower === 'alt') return mac ? '⌥' : 'Alt';
		return DISPLAY[lower] ?? (part.length === 1 ? part.toUpperCase() : part);
	});
	return parts.join(mac ? '' : '+');
}

export type DispatchResult = 'ran' | 'blocked' | 'ignored';

export interface ShortcutRegistry {
	readonly shortcuts: readonly Shortcut[];
	/** `blocked` means the shortcut matched but is disabled (read-only); the caller may swallow it. */
	dispatch(event: KeyEventLike, context: ShortcutContext): DispatchResult;
	/** A lone Alt press-and-release (non-Apple platforms), which focuses the ribbon. */
	dispatchAltTap(context: ShortcutContext): DispatchResult;
	/** Key text for the shortcut's first spec, for menus and help. */
	keysFor(id: ShortcutId, mac: boolean): string[];
}

export function createShortcutRegistry(actions: ShortcutActions): ShortcutRegistry {
	const shortcuts: Shortcut[] = [
		{ id: 'find', keys: ['Mod+F'], label: 'shortcut.find', editing: false, run: actions.find },
		{
			id: 'replace',
			keys: ['Mod+H'],
			label: 'shortcut.replace',
			editing: true,
			run: actions.replace,
		},
		{ id: 'save', keys: ['Mod+S'], label: 'shortcut.save', editing: false, run: actions.save },
		{ id: 'print', keys: ['Mod+P'], label: 'shortcut.print', editing: false, run: actions.print },
		{
			id: 'next-region',
			keys: ['F6'],
			label: 'shortcut.nextRegion',
			editing: false,
			run: () => actions.moveRegion(1),
		},
		{
			id: 'previous-region',
			keys: ['Shift+F6'],
			label: 'shortcut.previousRegion',
			editing: false,
			run: () => actions.moveRegion(-1),
		},
		{
			id: 'focus-ribbon',
			keys: ['F10', 'Alt'],
			label: 'shortcut.focusRibbon',
			editing: false,
			run: actions.focusRibbon,
		},
		{
			id: 'context-menu',
			keys: ['Shift+F10', 'ContextMenu'],
			label: 'shortcut.contextMenu',
			editing: false,
			run: actions.contextMenu,
		},
		{
			id: 'escape',
			keys: ['Escape'],
			label: 'shortcut.escape',
			editing: false,
			run: actions.escape,
		},
		{
			id: 'help',
			keys: ['Mod+/', 'F1'],
			label: 'shortcut.help',
			editing: false,
			run: actions.help,
		},
	];
	const attempt = (shortcut: Shortcut, context: ShortcutContext): DispatchResult => {
		if (shortcut.editing && context.readOnly) return 'blocked';
		return shortcut.run() === false ? 'ignored' : 'ran';
	};
	const shown = (shortcut: Shortcut, mac: boolean) =>
		shortcut.keys.filter((spec) => !(mac && parse(spec).tap));
	return {
		shortcuts,
		dispatch(event, context) {
			for (const shortcut of shortcuts)
				if (shortcut.keys.some((spec) => matchesKeys(event, spec, context.mac)))
					return attempt(shortcut, context);
			return 'ignored';
		},
		dispatchAltTap(context) {
			if (context.mac) return 'ignored';
			const shortcut = shortcuts.find((item) => item.keys.some((spec) => parse(spec).tap));
			return shortcut ? attempt(shortcut, context) : 'ignored';
		},
		keysFor(id, mac) {
			const shortcut = shortcuts.find((item) => item.id === id);
			return shortcut ? shown(shortcut, mac).map((spec) => formatKeys(spec, mac)) : [];
		},
	};
}

/**
 * Listens for keydown on `host` (events from the shadow tree bubble to it) and runs the registry.
 * Events a deeper handler already consumed (`defaultPrevented`, e.g. ProseMirror's own keymap) are
 * left alone. Returns a disposer.
 */
export function installKeyboard(
	host: HTMLElement,
	registry: ShortcutRegistry,
	context: () => ShortcutContext,
): () => void {
	let altPending = false;
	const onKeyDown = (event: KeyboardEvent) => {
		altPending = event.key === 'Alt' && !event.ctrlKey && !event.metaKey && !event.shiftKey;
		if (event.defaultPrevented || event.isComposing) return;
		const result = registry.dispatch(event, context());
		if (result !== 'ignored') event.preventDefault();
	};
	const onKeyUp = (event: KeyboardEvent) => {
		const tapped = altPending && event.key === 'Alt';
		altPending = false;
		if (tapped && registry.dispatchAltTap(context()) !== 'ignored') event.preventDefault();
	};
	host.addEventListener('keydown', onKeyDown);
	host.addEventListener('keyup', onKeyUp);
	return () => {
		host.removeEventListener('keydown', onKeyDown);
		host.removeEventListener('keyup', onKeyUp);
	};
}
