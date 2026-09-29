import { localeOf, translate } from './localization';
import { swatchColor } from './ribbon-colors';

let close: (() => void) | undefined;

/**
 * Shows `pop` under `anchor` and wires dismissal: Escape, a pointer press elsewhere, or `close()`.
 * Fixed positioning keeps it clear of the ribbon's overflow clipping. Returns false (and closes) when
 * the same anchor was already open, so the caret works as a toggle.
 */
function mountPopover(anchor: HTMLElement, pop: HTMLElement, expanded: HTMLElement): boolean {
	const wasOpen = expanded.getAttribute('aria-expanded') === 'true';
	close?.();
	if (wasOpen) return false;
	const box = anchor.getBoundingClientRect();
	pop.style.left = `${Math.max(4, box.left)}px`;
	pop.style.top = `${box.bottom + 2}px`;
	const root = anchor.getRootNode();
	(root instanceof ShadowRoot ? root : document.body).append(pop);
	expanded.setAttribute('aria-expanded', 'true');
	const outside = (event: Event) => {
		if (!event.composedPath().includes(pop) && !event.composedPath().includes(anchor)) close?.();
	};
	const escape = (event: KeyboardEvent) => {
		if (event.key !== 'Escape') return;
		close?.();
		anchor.querySelector('input')?.focus();
	};
	document.addEventListener('pointerdown', outside, true);
	document.addEventListener('keydown', escape, true);
	close = () => {
		document.removeEventListener('pointerdown', outside, true);
		document.removeEventListener('keydown', escape, true);
		pop.remove();
		expanded.removeAttribute('aria-expanded');
		close = undefined;
	};
	return true;
}

/** Opens a swatch grid under the caret button `anchor`. */
export function openSwatchPopover(
	anchor: HTMLElement,
	choices: Array<[string, string]>,
	choose: (value: string) => void,
): void {
	const locale = localeOf(anchor);
	const pop = document.createElement('div');
	pop.className = 'ribbon-popover';
	pop.setAttribute('role', 'menu');
	for (const [value, name] of choices) {
		const item = document.createElement('button');
		item.type = 'button';
		item.setAttribute('role', 'menuitem');
		const label = translate(locale, name as never);
		item.setAttribute('aria-label', label);
		item.title = label;
		item.className = value === 'none' ? 'swatch swatch-none' : 'swatch';
		item.style.setProperty('--swatch', swatchColor(value));
		item.addEventListener('mousedown', (event) => event.preventDefault());
		item.addEventListener('click', () => {
			close?.();
			choose(value);
		});
		pop.append(item);
	}
	const box = anchor.getBoundingClientRect();
	mountPopover(anchor, pop, anchor);
	pop.style.left = `${Math.max(4, box.left - 60)}px`;
}

/**
 * Opens a scrolling list of `items` under a combo box, with `current` marked and scrolled into view.
 * Up/Down move through the entries, Enter chooses, Escape closes. `preview` sets each entry in its
 * own font face.
 */
export function openListPopover(
	anchor: HTMLElement,
	items: readonly string[],
	current: string,
	preview: boolean,
	choose: (value: string) => void,
): void {
	const pop = document.createElement('div');
	pop.className = 'ribbon-popover ribbon-list';
	pop.setAttribute('role', 'listbox');
	const buttons = items.map((value) => {
		const item = document.createElement('button');
		item.type = 'button';
		item.setAttribute('role', 'option');
		item.textContent = value;
		item.setAttribute('aria-selected', String(value === current));
		if (preview) item.style.fontFamily = `"${value}", sans-serif`;
		item.addEventListener('mousedown', (event) => event.preventDefault());
		item.addEventListener('click', () => {
			close?.();
			choose(value);
		});
		item.addEventListener('keydown', (event) => {
			const at = buttons.indexOf(item);
			const next =
				event.key === 'ArrowDown' ? at + 1 : event.key === 'ArrowUp' ? at - 1 : Number.NaN;
			if (Number.isNaN(next)) return;
			event.preventDefault();
			buttons[(next + buttons.length) % buttons.length]?.focus();
		});
		pop.append(item);
		return item;
	});
	if (!mountPopover(anchor, pop, anchor.querySelector('input') ?? anchor)) return;
	const selected = buttons.find((item) => item.getAttribute('aria-selected') === 'true');
	(selected ?? buttons[0])?.scrollIntoView?.({ block: 'center' });
	(selected ?? buttons[0])?.focus({ preventScroll: true });
}
