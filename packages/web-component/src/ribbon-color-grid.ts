import { localeOf, translate } from './localization';
import { closeRibbonPopover, mountPopover } from './ribbon-popover';

/** Word's Office theme colours, in Word's column order, with their names. */
const THEME: Array<[string, string]> = [
	['#ffffff', 'White, Background 1'],
	['#000000', 'Black, Text 1'],
	['#e7e6e6', 'Light Gray, Background 2'],
	['#44546a', 'Blue-Gray, Text 2'],
	['#4472c4', 'Blue, Accent 1'],
	['#ed7d31', 'Orange, Accent 2'],
	['#a5a5a5', 'Gray, Accent 3'],
	['#ffc000', 'Gold, Accent 4'],
	['#5b9bd5', 'Blue, Accent 5'],
	['#70ad47', 'Green, Accent 6'],
];
const STANDARD: Array<[string, string]> = [
	['#c00000', 'Dark Red'],
	['#ff0000', 'Red'],
	['#ffc000', 'Orange'],
	['#ffff00', 'Yellow'],
	['#92d050', 'Light Green'],
	['#00b050', 'Green'],
	['#00b0f0', 'Light Blue'],
	['#0070c0', 'Blue'],
	['#002060', 'Dark Blue'],
	['#7030a0', 'Purple'],
];

/** Mixes `hex` toward `target` (0 to 255 per channel) by `amount` (0 to 1). */
function mix(hex: string, target: number, amount: number): string {
	const channel = (at: number) => {
		const value = parseInt(hex.slice(at, at + 2), 16);
		return Math.round(value + (target - value) * amount)
			.toString(16)
			.padStart(2, '0');
	};
	return `#${channel(1)}${channel(3)}${channel(5)}`;
}

/** The five variants under a theme colour, as Word's grid shows them (lighter, then darker). */
export function themeShades(hex: string): string[] {
	if (hex === '#ffffff') return [0.05, 0.15, 0.25, 0.35, 0.5].map((n) => mix(hex, 0, n));
	if (hex === '#000000') return [0.5, 0.35, 0.25, 0.15, 0.05].map((n) => mix(hex, 255, n));
	return [
		mix(hex, 255, 0.8),
		mix(hex, 255, 0.6),
		mix(hex, 255, 0.4),
		mix(hex, 0, 0.25),
		mix(hex, 0, 0.5),
	];
}

export interface ColorGridOptions {
	/** Label of the first row's "clear" command, or omitted when the property cannot be cleared. */
	noneLabel?: string;
}

/**
 * Word's colour picker: a no-colour row, a Theme Colors grid (ten columns, each a colour over five
 * variants), Standard Colors and More Colors (the browser's colour dialog).
 */
export function openColorGridPopover(
	anchor: HTMLElement,
	choose: (value: string) => void,
	options: ColorGridOptions = {},
): void {
	const locale = localeOf(anchor);
	const say = (text: string) => translate(locale, text as never);
	const pop = document.createElement('div');
	pop.className = 'ribbon-popover color-grid';
	pop.setAttribute('role', 'menu');
	const pick = (value: string) => {
		closeRibbonPopover();
		choose(value);
	};
	const swatch = (value: string, name: string) => {
		const item = document.createElement('button');
		item.type = 'button';
		item.setAttribute('role', 'menuitem');
		item.className = 'swatch';
		item.style.setProperty('--swatch', value);
		item.setAttribute('aria-label', say(name));
		item.title = say(name);
		item.dataset.color = value;
		item.addEventListener('mousedown', (event) => event.preventDefault());
		item.addEventListener('click', () => pick(value));
		return item;
	};
	const title = (text: string) => {
		const el = document.createElement('div');
		el.className = 'color-grid-title';
		el.textContent = say(text);
		return el;
	};
	const grid = (...items: HTMLElement[]) => {
		const el = document.createElement('div');
		el.className = 'color-grid-cells';
		el.append(...items);
		return el;
	};
	if (options.noneLabel) {
		const none = document.createElement('button');
		none.type = 'button';
		none.setAttribute('role', 'menuitem');
		none.className = 'color-grid-command';
		none.textContent = say(options.noneLabel);
		none.addEventListener('mousedown', (event) => event.preventDefault());
		none.addEventListener('click', () => pick('none'));
		pop.append(none);
	}
	const shades = THEME.map(([hex]) => themeShades(hex));
	pop.append(
		title('Theme Colors'),
		grid(
			...THEME.map(([hex, name]) => swatch(hex, name)),
			...[0, 1, 2, 3, 4].flatMap((row) =>
				THEME.map(([, name], column) => swatch(shades[column]![row]!, name)),
			),
		),
		title('Standard Colors'),
		grid(...STANDARD.map(([hex, name]) => swatch(hex, name))),
	);
	const more = document.createElement('button');
	more.type = 'button';
	more.setAttribute('role', 'menuitem');
	more.className = 'color-grid-command';
	more.textContent = say('More Colors…');
	const input = Object.assign(document.createElement('input'), { type: 'color' });
	input.className = 'color-grid-input';
	input.tabIndex = -1;
	input.addEventListener('change', () => pick(input.value));
	more.addEventListener('mousedown', (event) => event.preventDefault());
	more.addEventListener('click', () => input.click());
	pop.append(more, input);
	if (!mountPopover(anchor, pop, anchor)) return;
	const box = anchor.getBoundingClientRect();
	pop.style.left = `${Math.max(4, box.left - 120)}px`;
}
