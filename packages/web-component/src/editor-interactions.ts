/**
 * Wires the keyboard shortcut registry, the shortcut help dialog and the context menu to the
 * editor. Everything dispatches through the existing controllers (`routeRibbonAction`, the chrome's
 * file commands, the search panel); this module holds only the focus and layer bookkeeping.
 */
import type { EditorCore } from './editor-core';
import { emit } from './events';
import { focusView } from './focus-view';
import { createContextMenu, type ContextMenu } from './context-menu';
import {
	createShortcutRegistry,
	installKeyboard,
	isMacPlatform,
	type ShortcutActions,
} from './keyboard';
import { routeRibbonAction } from './ribbon-router';
import { createShortcutHelp, shortcutRows } from './shortcut-help';

type Region = 'ribbon' | 'document' | 'status';
const REGIONS: Region[] = ['ribbon', 'document', 'status'];

/** Something Escape can close: a dialog, panel or menu. */
interface Layer {
	element: HTMLElement | undefined;
	isOpen(): boolean;
	close(): void;
	/** The layer already handles its own Escape key while focus is inside it. */
	selfHandled: boolean;
}

export function attachEditorInteractions(core: EditorCore, frame: HTMLElement): void {
	const { element, shell } = core;
	const activeElement = () => element.shadowRoot?.activeElement ?? null;
	// A read-only ProseMirror surface cannot take focus, so the canvas stands in for it.
	shell.canvas!.tabIndex = -1;
	const toDocument = () => {
		const view = core.targetView();
		if (view?.editable) focusView(view);
		else shell.canvas?.focus();
	};

	const help = createShortcutHelp(toDocument);
	frame.append(help.element);
	const menu: ContextMenu = createContextMenu({
		container: frame,
		target: shell.canvas!,
		view: () => core.targetView(),
		focusDocument: toDocument,
		readOnly: () => core.readOnly,
		locale: () => core.locale,
		run: (action) => routeRibbonAction(core, action),
		warn: (message) => emit(element, 'document-warning', message),
	});

	const regionElement = (region: Region): HTMLElement | undefined =>
		region === 'ribbon'
			? shell.toolbar
			: region === 'status'
				? shell.chrome?.statusBar.element
				: undefined;
	const currentRegion = (): Region => {
		const active = activeElement();
		return (
			REGIONS.find((region) => active && regionElement(region)?.contains(active)) ?? 'document'
		);
	};
	const focusRegion = (region: Region) => {
		if (region === 'document') return toDocument();
		const root = regionElement(region);
		const target =
			region === 'ribbon'
				? root?.querySelector<HTMLElement>('[role="tab"][tabindex="0"], [role="tab"]')
				: root?.querySelector<HTMLElement>('button:not([hidden])');
		(target ?? root)?.focus();
	};

	const layers: Layer[] = [
		{ element: undefined, isOpen: () => menu.isOpen, close: () => menu.close(), selfHandled: true },
		{ element: help.element, isOpen: () => help.isOpen, close: help.close, selfHandled: true },
		{
			element: core.inserts.linkDialog.element,
			isOpen: () => core.inserts.linkDialog.isOpen,
			close: () => core.inserts.linkDialog.close(),
			selfHandled: true,
		},
		{
			element: core.inserts.pictureDialog.element,
			isOpen: () => core.inserts.pictureDialog.isOpen,
			close: () => core.inserts.pictureDialog.close(),
			selfHandled: true,
		},
		{
			element: shell.chrome?.backstage.element,
			isOpen: () => Boolean(shell.chrome?.backstage.isOpen),
			close: () => shell.chrome?.closeBackstage(),
			selfHandled: true,
		},
		{
			element: shell.searchPanel?.element,
			isOpen: () => Boolean(shell.searchPanel?.isOpen),
			close: () => shell.searchPanel?.close(),
			selfHandled: true,
		},
		{
			element: shell.review?.commentsPanel.element,
			isOpen: () => Boolean(shell.review?.commentsOpen),
			close: () => {
				shell.review?.handleComments('toggle');
				toDocument();
			},
			selfHandled: false,
		},
	];
	// Snapshot in the capture phase, before a layer's own Escape handler closes it.
	let escapeOwned = false;
	element.addEventListener(
		'keydown',
		(event) => {
			if (event.key !== 'Escape') return;
			const path = event.composedPath();
			escapeOwned = layers.some(
				(layer) =>
					layer.selfHandled &&
					layer.isOpen() &&
					(layer.element ? path.includes(layer.element) : path.includes(menu.element as Node)),
			);
		},
		true,
	);

	const actions: ShortcutActions = {
		find: () => shell.searchPanel?.open('find'),
		replace: () => shell.searchPanel?.open('replace'),
		save: () => void shell.chrome?.run('save'),
		print: () => void shell.chrome?.run('print'),
		moveRegion: (direction) => {
			const next = (REGIONS.indexOf(currentRegion()) + direction + REGIONS.length) % REGIONS.length;
			const region = REGIONS[next];
			if (region) focusRegion(region);
		},
		focusRibbon: () => focusRegion('ribbon'),
		contextMenu: () => {
			if (currentRegion() === 'document') menu.openAtCaret();
		},
		help: () => {
			menu.close(false);
			help.open(shortcutRows(registry, mac()), core.locale);
		},
		escape: () => {
			if (escapeOwned) return false;
			const open = layers.find((layer) => layer.isOpen());
			if (open) open.close();
			else if (currentRegion() === 'document') return false;
			toDocument();
			return true;
		},
	};
	const mac = () => isMacPlatform();
	const registry = createShortcutRegistry(actions);
	installKeyboard(element, registry, () => ({ readOnly: core.readOnly, mac: mac() }));
}
