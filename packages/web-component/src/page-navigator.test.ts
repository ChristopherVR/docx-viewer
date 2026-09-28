// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { PageNavigator, THUMBNAIL_WIDTH } from './page-navigator';
import { PageTracker } from './page-sync';
import { at } from './test-support';

function sheets(count: number): HTMLElement[] {
	return Array.from({ length: count }, (_, index) => {
		const sheet = document.createElement('div');
		sheet.className = 'dve-print-page';
		sheet.id = `sheet-${index}`;
		sheet.style.width = '600px';
		sheet.style.height = '800px';
		sheet.textContent = `Page ${index + 1} text`;
		return sheet;
	});
}

function setup(count = 3) {
	let pages = sheets(count);
	let version = 1;
	const goTo = vi.fn();
	const close = vi.fn();
	const navigator = new PageNavigator({ pages: () => pages, version: () => version, goTo, close });
	document.body.append(navigator.element);
	const state = (current: number, total = pages.length, active = true) => ({
		active,
		status: active ? { current, total } : null,
		locale: 'en' as const,
	});
	return {
		navigator,
		goTo,
		close,
		state,
		relayout(next: number) {
			pages = sheets(next);
			version++;
		},
	};
}

const options = (navigator: PageNavigator) => [
	...navigator.element.querySelectorAll<HTMLElement>('[role="option"]'),
];

describe('PageNavigator', () => {
	afterEach(() => document.body.replaceChildren());

	it('is hidden and empty until opened, then lists one option per page', () => {
		const { navigator, state } = setup(3);
		navigator.sync(state(1));
		expect(navigator.element.hidden).toBe(true);
		expect(navigator.pageCount).toBe(0);
		navigator.setOpen(true);
		expect(navigator.element.hidden).toBe(false);
		expect(navigator.pageCount).toBe(3);
		expect(options(navigator).map((option) => option.dataset.page)).toEqual(['1', '2', '3']);
		expect(navigator.element.getAttribute('aria-label')).toBe('Page thumbnails');
		expect(navigator.element.querySelector('[role="listbox"]')).not.toBeNull();
	});

	it('tracks the current page with aria-selected and a single tab stop', () => {
		const { navigator, state } = setup(3);
		navigator.setOpen(true);
		navigator.sync(state(2));
		expect(navigator.currentPage).toBe(2);
		expect(options(navigator).map((o) => o.getAttribute('aria-selected'))).toEqual([
			'false',
			'true',
			'false',
		]);
		expect(options(navigator).map((o) => o.tabIndex)).toEqual([-1, 0, -1]);
		navigator.sync(state(3));
		expect(navigator.currentPage).toBe(3);
		expect(at(options(navigator), 2).getAttribute('aria-label')).toBe('Page 3 of 3');
	});

	it('draws scaled clones without duplicate ids', () => {
		const { navigator, state } = setup(2);
		navigator.setOpen(true);
		navigator.sync(state(1));
		const clone = at(options(navigator), 0).querySelector<HTMLElement>('.dve-print-page')!;
		expect(clone.textContent).toBe('Page 1 text');
		expect(clone.hasAttribute('id')).toBe(false);
		expect(clone.style.transform).toBe(`scale(${THUMBNAIL_WIDTH / 600})`);
		expect(
			at(options(navigator), 0).querySelector<HTMLElement>('.dve-page-thumb-sheet')!.style.height,
		).toBe('160px');
	});

	it('scrolls to a page on click and on Enter, and moves focus with the arrow keys', () => {
		const { navigator, goTo, state } = setup(3);
		navigator.setOpen(true);
		navigator.sync(state(1));
		at(options(navigator), 2).click();
		expect(goTo).toHaveBeenLastCalledWith(3);
		const second = at(options(navigator), 1);
		second.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
		expect(goTo).toHaveBeenLastCalledWith(2);
		const first = at(options(navigator), 0);
		first.focus();
		first.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
		expect(at(options(navigator), 1).tabIndex).toBe(0);
		expect(at(options(navigator), 0).tabIndex).toBe(-1);
		at(options(navigator), 1).dispatchEvent(
			new KeyboardEvent('keydown', { key: 'End', bubbles: true }),
		);
		expect(at(options(navigator), 2).tabIndex).toBe(0);
		at(options(navigator), 2).dispatchEvent(
			new KeyboardEvent('keydown', { key: 'Home', bubbles: true }),
		);
		expect(at(options(navigator), 0).tabIndex).toBe(0);
		expect(goTo).toHaveBeenCalledTimes(2);
	});

	it('rebuilds after a relayout with a different page count', () => {
		const { navigator, state, relayout } = setup(2);
		navigator.setOpen(true);
		navigator.sync(state(1));
		relayout(5);
		navigator.sync(state(4, 5));
		expect(navigator.pageCount).toBe(5);
		expect(navigator.currentPage).toBe(4);
	});

	it('explains that thumbnails need Print Layout instead of listing pages', () => {
		const { navigator, state } = setup(3);
		navigator.setOpen(true);
		navigator.sync(state(1, 3, false));
		const message = navigator.element.querySelector<HTMLElement>('.dve-pages-message')!;
		expect(message.hidden).toBe(false);
		expect(message.textContent).toContain('need Print Layout');
		expect(navigator.element.querySelector<HTMLElement>('[role="listbox"]')!.hidden).toBe(true);
		expect(navigator.pageCount).toBe(0);
	});

	it('labels itself in French and keeps the approximation tooltip', () => {
		const { navigator, state } = setup(2);
		navigator.setOpen(true);
		navigator.sync({ ...state(1), locale: 'fr' });
		expect(navigator.element.getAttribute('aria-label')).toBe('Miniatures des pages');
		expect(at(options(navigator), 1).getAttribute('aria-label')).toBe('Page 2 sur 2');
		expect(navigator.element.title).toContain('approximation');
	});

	it('closes through the header button', () => {
		const { navigator, close } = setup(1);
		navigator.element.querySelector<HTMLButtonElement>('.dve-pages-close')!.click();
		expect(close).toHaveBeenCalledTimes(1);
	});
});

describe('PageTracker', () => {
	it('emits page-change only when the page or count changes', () => {
		const target = new EventTarget();
		const seen: unknown[] = [];
		target.addEventListener('page-change', (event) => seen.push((event as CustomEvent).detail));
		const tracker = new PageTracker(target);
		tracker.report(null);
		tracker.report({ current: 1, total: 3 });
		tracker.report({ current: 1, total: 3 });
		tracker.report({ current: 2, total: 3 });
		tracker.report({ current: 2, total: 4 });
		expect(seen).toEqual([
			{ page: 1, pageCount: 3 },
			{ page: 2, pageCount: 3 },
			{ page: 2, pageCount: 4 },
		]);
	});
});
