/**
 * Context menu for the document surface. Opens on right click, Shift+F10 or the ContextMenu key.
 * What it offers comes from context-menu-items.ts; this module owns the DOM, keyboard navigation,
 * dismissal and viewport clamping.
 */
import type { EditorView } from 'prosemirror-view';
import { clipboardReadAllowed, copyOrCut, pasteClipboard } from './context-menu-actions';
import {
	clampToViewport,
	computeMenuItems,
	menuStateFromView,
	type MenuItem,
} from './context-menu-items';
import { removeLink } from './link-commands';
import { translate, type EditorLocale } from './localization';
import type { RibbonAction } from './ribbon-action';

export interface ContextMenuHost {
	/** Element the menu is appended to (inside the shadow root, so theme tokens apply). */
	container: HTMLElement;
	/** Element whose `contextmenu` events open the menu. */
	target: HTMLElement;
	view(): EditorView | undefined;
	/** Returns focus to the document surface (which may be non-editable). */
	focusDocument(): void;
	readOnly(): boolean;
	locale(): EditorLocale;
	run(action: RibbonAction): void;
	warn(message: string): void;
}

export interface ContextMenu {
	readonly isOpen: boolean;
	readonly element: HTMLElement | undefined;
	open(x: number, y: number): void;
	/** Opens at the caret, for Shift+F10 and the ContextMenu key. */
	openAtCaret(): void;
	close(restoreFocus?: boolean): void;
}

/** Browsers may also fire `contextmenu` for the key that we already handled; ignore that echo. */
const KEY_ECHO_MS = 500;

export function createContextMenu(host: ContextMenuHost): ContextMenu {
	let menu: HTMLElement | undefined;
	let disposeDismiss: (() => void) | undefined;
	let keyOpenedAt = -Infinity;

	const enabled = () =>
		menu ? [...menu.querySelectorAll<HTMLElement>('[role="menuitem"]:not([aria-disabled])')] : [];

	const close = (restoreFocus = true) => {
		if (!menu) return;
		menu.remove();
		menu = undefined;
		disposeDismiss?.();
		disposeDismiss = undefined;
		if (restoreFocus) host.focusDocument();
	};

	const execute = async (item: MenuItem, view: EditorView) => {
		const { command } = item;
		if (command.kind === 'ribbon') host.run(command.action);
		else if (command.kind === 'removeLink') {
			removeLink(view);
			view.focus();
		} else if (command.op === 'paste') {
			if (!(await pasteClipboard(view)))
				host.warn(translate(host.locale(), 'menu.clipboardDenied'));
		} else if (!(await copyOrCut(view, command.op)))
			host.warn(translate(host.locale(), 'menu.clipboardDenied'));
	};

	const activate = (item: MenuItem, button: HTMLElement) => {
		const view = host.view();
		if (!view || button.hasAttribute('aria-disabled')) return;
		close();
		void execute(item, view);
	};

	const move = (step: number | 'first' | 'last') => {
		const list = enabled();
		if (!list.length) return;
		const root = menu!.getRootNode() as Document | ShadowRoot;
		const current = list.indexOf(root.activeElement as HTMLElement);
		const last = list.length - 1;
		const next =
			step === 'first' || (current < 0 && step === 1)
				? 0
				: step === 'last' || current < 0
					? last
					: (current + step + list.length) % list.length;
		list[next]?.focus();
	};

	const onKeyDown = (event: KeyboardEvent) => {
		const handled = (action: () => void) => {
			event.preventDefault();
			event.stopPropagation();
			action();
		};
		if (event.key === 'ArrowDown') handled(() => move(1));
		else if (event.key === 'ArrowUp') handled(() => move(-1));
		else if (event.key === 'Home') handled(() => move('first'));
		else if (event.key === 'End') handled(() => move('last'));
		else if (event.key === 'Escape' || event.key === 'Tab') handled(() => close());
		else if (event.key === 'Enter' || event.key === ' ')
			handled(() => (event.target as HTMLElement).click());
	};

	const render = (items: MenuItem[], locale: EditorLocale) => {
		const element = document.createElement('div');
		element.className = 'dve-context-menu';
		element.setAttribute('role', 'menu');
		element.setAttribute('aria-label', translate(locale, 'menu.context'));
		for (const item of items) {
			if (item.separatorBefore) {
				const separator = document.createElement('div');
				separator.setAttribute('role', 'separator');
				element.append(separator);
			}
			const button = document.createElement('button');
			button.type = 'button';
			button.tabIndex = -1;
			button.dataset.item = item.id;
			button.setAttribute('role', 'menuitem');
			button.textContent = translate(locale, item.label);
			if (item.disabled) button.setAttribute('aria-disabled', 'true');
			if (item.disabled && item.reason) button.title = translate(locale, item.reason);
			button.addEventListener('click', () => activate(item, button));
			element.append(button);
		}
		element.addEventListener('keydown', onKeyDown);
		// Pointer movement follows hover like a native menu, without stealing it from the keyboard.
		element.addEventListener('mouseover', (event) => {
			const target = (event.target as Element).closest<HTMLElement>('[role="menuitem"]');
			if (target && !target.hasAttribute('aria-disabled')) target.focus();
		});
		return element;
	};

	const watchDismissal = () => {
		const onPointerDown = (event: Event) => {
			if (!event.composedPath().includes(menu!)) close(false);
		};
		const dismiss = () => close(false);
		document.addEventListener('pointerdown', onPointerDown, true);
		window.addEventListener('scroll', dismiss, true);
		// `scroll` is not composed, so the shadow-tree canvas needs its own listener.
		host.target.addEventListener('scroll', dismiss);
		window.addEventListener('resize', dismiss);
		window.addEventListener('blur', dismiss);
		return () => {
			document.removeEventListener('pointerdown', onPointerDown, true);
			window.removeEventListener('scroll', dismiss, true);
			host.target.removeEventListener('scroll', dismiss);
			window.removeEventListener('resize', dismiss);
			window.removeEventListener('blur', dismiss);
		};
	};

	const open = (x: number, y: number) => {
		const view = host.view();
		if (!view) return;
		close(false);
		const locale = host.locale();
		const items = computeMenuItems(menuStateFromView(view, host.readOnly()));
		const element = render(items, locale);
		menu = element;
		element.style.visibility = 'hidden';
		host.container.append(element);
		const box = element.getBoundingClientRect();
		const { left, top } = clampToViewport(
			x,
			y,
			{ width: box.width, height: box.height },
			{ width: window.innerWidth, height: window.innerHeight },
		);
		element.style.left = `${left}px`;
		element.style.top = `${top}px`;
		element.style.visibility = '';
		disposeDismiss = watchDismissal();
		move('first');
		const paste = element.querySelector<HTMLElement>('[data-item="paste"]');
		if (paste && !paste.hasAttribute('aria-disabled'))
			void clipboardReadAllowed().then((allowed) => {
				if (allowed || menu !== element) return;
				paste.setAttribute('aria-disabled', 'true');
				paste.title = translate(locale, 'menu.clipboardDenied');
			});
	};

	host.target.addEventListener('contextmenu', (event) => {
		if (!(event.target as Element).closest?.('.ProseMirror')) return;
		event.preventDefault();
		if (performance.now() - keyOpenedAt < KEY_ECHO_MS) return;
		const { clientX, clientY } = event;
		// Let the browser and ProseMirror move the caret to the click before reading the selection.
		setTimeout(() => open(clientX, clientY), 0);
	});

	return {
		get isOpen() {
			return Boolean(menu);
		},
		get element() {
			return menu;
		},
		open,
		openAtCaret() {
			const view = host.view();
			if (!view) return;
			keyOpenedAt = performance.now();
			const caret = view.coordsAtPos(view.state.selection.head);
			open(caret.left, caret.bottom);
		},
		close,
	};
}
