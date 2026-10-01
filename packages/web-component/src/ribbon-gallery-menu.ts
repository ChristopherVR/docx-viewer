import { PAGE_SIZES } from './page-size';
import { closeRibbonPopover, mountPopover } from './ribbon-popover';
import type { RibbonAction } from './ribbon-action';
import type { RibbonIcon } from './ribbon-icons';
import { menuSelect } from './ribbon-parts';
import { emit } from './events';
import { localeOf, translate } from './localization';

/** What a gallery entry looks like: a page thumbnail and a line of detail under the name. */
export interface GalleryLook {
	thumb: SVGSVGElement;
	detail?: string;
}
export type GalleryKind =
	| 'margins'
	| 'size'
	| 'orientation'
	| 'columns'
	| 'verticalAlign'
	| 'borders'
	| 'pageNumber'
	| 'lineSpacing';

/** A command listed under a gallery's entries, such as Line Spacing Options. */
export interface GalleryCommand {
	label: string;
	action: RibbonAction;
}

const NS = 'http://www.w3.org/2000/svg';
const inches = (twips: number) => `${+(twips / 1440).toFixed(2)}"`;

function shape(name: string, attrs: Record<string, string | number>): SVGElement {
	const el = document.createElementNS(NS, name);
	for (const [key, value] of Object.entries(attrs)) el.setAttribute(key, String(value));
	return el;
}

/** A page outline `w` by `h`, plus whatever `body` draws inside it. */
function page(w: number, h: number, ...body: SVGElement[]): SVGSVGElement {
	const svg = document.createElementNS(NS, 'svg');
	svg.setAttribute('viewBox', `0 0 ${w + 2} ${h + 2}`);
	svg.setAttribute('width', String((w + 2) * 1.4));
	svg.setAttribute('height', String((h + 2) * 1.4));
	svg.setAttribute('aria-hidden', 'true');
	svg.classList.add('gallery-thumb');
	svg.append(shape('rect', { x: 1, y: 1, width: w, height: h, class: 'thumb-page' }), ...body);
	return svg;
}

/** `count` text lines spread over the box `x`, `y`, `w` by `h`. */
const lines = (x: number, y: number, w: number, h: number, count = 5) =>
	Array.from({ length: count }, (_, i) =>
		shape('line', {
			x1: x,
			x2: x + Math.max(w, 1),
			y1: y + (h * (i + 0.5)) / count,
			y2: y + (h * (i + 0.5)) / count,
			class: 'thumb-line',
		}),
	);

/** Word's Margins presets: top and bottom, then left and right, in inches. */
const MARGIN_INCHES: Record<string, [top: number, side: number]> = {
	normal: [1, 1],
	narrow: [0.5, 0.5],
	moderate: [1, 0.75],
	wide: [1, 2],
};

/** A small square with the edges `value` draws in solid ink and the others dotted. */
function borderThumb(value: string): SVGSVGElement {
	const svg = document.createElementNS(NS, 'svg');
	svg.setAttribute('viewBox', '0 0 20 20');
	svg.setAttribute('width', '22');
	svg.setAttribute('height', '22');
	svg.setAttribute('aria-hidden', 'true');
	svg.classList.add('gallery-thumb');
	const on: Record<string, string[]> = {
		bottom: ['b'],
		top: ['t'],
		left: ['l'],
		right: ['r'],
		all: ['t', 'b', 'l', 'r', 'h', 'v'],
		outside: ['t', 'b', 'l', 'r'],
		insideH: ['h'],
		insideV: ['v'],
		horizontal: ['h'],
		none: [],
	};
	const edges: Record<string, [number, number, number, number]> = {
		t: [3, 3, 17, 3],
		b: [3, 17, 17, 17],
		l: [3, 3, 3, 17],
		r: [17, 3, 17, 17],
		h: [3, 10, 17, 10],
		v: [10, 3, 10, 17],
	};
	const active = new Set(on[value] ?? []);
	for (const [edge, [x1, y1, x2, y2]] of Object.entries(edges))
		svg.append(
			shape('line', {
				x1,
				y1,
				x2,
				y2,
				class: active.has(edge) ? 'thumb-edge-on' : 'thumb-edge-off',
			}),
		);
	return svg;
}

/** Three lines spaced like the chosen line spacing, over a small arrow marking the gap. */
function spacingThumb(value: string): SVGSVGElement {
	const twips = Number(/^auto:(\d+)$/.exec(value)?.[1] ?? 240);
	const gap = Math.max(3, Math.min(9, (twips / 240) * 3.5));
	const svg = document.createElementNS(NS, 'svg');
	svg.setAttribute('viewBox', '0 0 24 24');
	svg.setAttribute('width', '24');
	svg.setAttribute('height', '24');
	svg.setAttribute('aria-hidden', 'true');
	svg.classList.add('gallery-thumb');
	const top = 12 - gap;
	for (const y of [top, 12, 12 + gap])
		svg.append(shape('line', { x1: 8, x2: 21, y1: y, y2: y, class: 'thumb-line thumb-line-dark' }));
	svg.append(shape('path', { d: 'M4 4v16M2 6l2-2 2 2M2 18l2 2 2-2', class: 'thumb-arrow' }));
	return svg;
}

/** The thumbnail and detail for `value` of the gallery `kind`; unknown values get a plain page. */
export function galleryLook(
	kind: GalleryKind,
	value: string,
	say: (text: string) => string = (text) => text,
): GalleryLook {
	if (kind === 'margins') {
		const [top, side] = MARGIN_INCHES[value] ?? [1, 1];
		const k = 4;
		return {
			thumb: page(30, 40, ...lines(1 + side * k, 1 + top * k, 30 - side * 2 * k, 40 - top * 2 * k)),
			detail: `${say('Top')}: ${top}"  ${say('Bottom')}: ${top}"  ${say('Left')}: ${side}"  ${say('Right')}: ${side}"`,
		};
	}
	if (kind === 'size') {
		const [w, h] = PAGE_SIZES[value] ?? [12240, 15840];
		const scale = 44 / 24480;
		const width = Math.round(w * scale);
		const height = Math.round(h * scale);
		return {
			thumb: page(width, height, ...lines(4, 5, width - 6, height - 9, 4)),
			detail: `${inches(w)} x ${inches(h)}`,
		};
	}
	if (kind === 'orientation') {
		return {
			thumb:
				value === 'landscape'
					? page(44, 32, ...lines(5, 5, 34, 22, 4))
					: page(32, 44, ...lines(5, 5, 22, 34, 6)),
		};
	}
	if (kind === 'columns') {
		if (value === 'left' || value === 'right') {
			const first = value === 'left' ? 6 : 13;
			return {
				thumb: page(
					30,
					40,
					...lines(4, 5, first, 30, 6),
					...lines(4 + first + 3, 5, 19 - first, 30, 6),
				),
			};
		}
		const count = Number(value) || 1;
		const gap = 3;
		const width = (30 - 8 - gap * (count - 1)) / count;
		const columns = Array.from({ length: count }, (_, i) =>
			lines(4 + i * (width + gap), 5, width, 30, 6),
		);
		return { thumb: page(30, 40, ...columns.flat()) };
	}
	if (kind === 'pageNumber') {
		const [position, align, format] = value.split(':');
		const number = shape('text', {
			x: align === 'left' ? 5 : align === 'right' ? 27 : 16,
			y: position === 'top' ? 6 : 39,
			'font-size': 3,
			'text-anchor': align === 'left' ? 'start' : align === 'right' ? 'end' : 'middle',
			fill: 'currentColor',
		});
		number.textContent = format === 'pageOfTotal' ? '1 / 2' : '1';
		return { thumb: page(30, 40, ...lines(5, 10, 22, 22, 6), number) };
	}
	if (kind === 'borders') return { thumb: borderThumb(value) };
	if (kind === 'lineSpacing') return { thumb: spacingThumb(value) };
	const top = { top: 5, center: 15, bottom: 25, both: 5 }[value as 'top'] ?? 5;
	const both = value === 'both';
	return { thumb: page(30, 40, ...lines(5, top, 20, both ? 30 : 10, both ? 7 : 3)) };
}

/**
 * A dropdown command like `menuSelect`, but choosing opens a gallery of page thumbnails (Word's
 * Margins, Size, Orientation and Columns) instead of the browser's plain list. The real `<select>`
 * stays in the DOM so the editor keeps reading and writing `select.value`, and the keyboard still
 * changes it directly.
 */
export function menuGallery(
	label: string,
	icon: RibbonIcon,
	kind: GalleryKind,
	values: Array<[string, string]>,
	action: (value: string) => RibbonAction,
	options: { compact?: boolean; momentary?: boolean; commands?: GalleryCommand[] } = {},
): HTMLElement {
	const wrap = menuSelect(label, icon, values, action, options);
	wrap.classList.add('ribbon-menu-gallery');
	const control = wrap.querySelector('select')!;
	const open = () => openGallery(wrap, control, kind, options.commands ?? []);
	wrap.addEventListener('click', (event) => {
		if (control.disabled) return;
		event.preventDefault();
		open();
	});
	control.addEventListener('keydown', (event) => {
		if (event.key === 'Enter' || event.key === ' ' || (event.altKey && event.key === 'ArrowDown')) {
			event.preventDefault();
			open();
		}
	});
	return wrap;
}

function openGallery(
	anchor: HTMLElement,
	control: HTMLSelectElement,
	kind: GalleryKind,
	commands: GalleryCommand[],
): void {
	const pop = document.createElement('div');
	pop.className = 'ribbon-popover ribbon-gallery-menu';
	pop.setAttribute('role', 'menu');
	pop.setAttribute('aria-label', control.getAttribute('aria-label') ?? '');
	const locale = localeOf(anchor);
	const say = (text: string) => translate(locale, text as never);
	for (const option of [...control.options]) {
		const look = galleryLook(kind, option.value, say);
		const item = document.createElement('button');
		item.type = 'button';
		item.setAttribute('role', 'menuitemradio');
		item.setAttribute('aria-checked', String(option.value === control.value));
		item.dataset.value = option.value;
		const name = document.createElement('span');
		name.className = 'gallery-name';
		name.textContent = option.textContent;
		const words = document.createElement('span');
		words.className = 'gallery-words';
		words.append(name);
		if (look.detail) {
			const detail = document.createElement('span');
			detail.className = 'gallery-detail';
			detail.textContent = look.detail;
			words.append(detail);
		}
		item.append(look.thumb, words);
		item.addEventListener('mousedown', (event) => event.preventDefault());
		item.addEventListener('click', () => {
			closeRibbonPopover();
			control.value = option.value;
			control.dispatchEvent(new Event('change', { bubbles: true }));
		});
		pop.append(item);
	}
	for (const command of commands) {
		const item = document.createElement('button');
		item.type = 'button';
		item.setAttribute('role', 'menuitem');
		item.className = 'gallery-command';
		item.textContent = say(command.label);
		item.addEventListener('mousedown', (event) => event.preventDefault());
		item.addEventListener('click', () => {
			closeRibbonPopover();
			emit(anchor, 'ribbon-action', command.action);
		});
		pop.append(item);
	}
	mountPopover(anchor, pop, anchor);
	pop.querySelector<HTMLElement>('[aria-checked="true"]')?.focus({ preventScroll: true });
}
