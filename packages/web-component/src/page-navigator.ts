import { formatPageStatus, translate, type EditorLocale } from './localization';

/** What the navigator needs from the Print Layout view and the element. */
export interface PageNavigatorHost {
	/** Laid-out page sheets in order; empty until a layout exists. */
	pages(): HTMLElement[];
	/** Changes whenever the layout was rebuilt. */
	version(): number;
	goTo(page: number): void;
	close(): void;
}

export interface PageNavigatorState {
	/** Print Layout is the active view. */
	active: boolean;
	/** Current page and total, or null while no layout exists. */
	status: { current: number; total: number } | null;
	locale: EditorLocale;
}

interface Thumb {
	option: HTMLElement;
	sheet: HTMLElement;
	source: HTMLElement;
	scale: number;
}

export const THUMBNAIL_WIDTH = 120;
/** Without IntersectionObserver (tests, very old browsers) only the first pages are drawn. */
const EAGER_LIMIT = 20;
const DEFAULT_PAGE = { width: 816, height: 1056 };

/**
 * Collapsible left rail of scaled page thumbnails. Thumbnails are clones of the Print Layout
 * sheets, drawn only while near the rail's viewport, so large documents stay cheap. Pagination
 * comes from the approximate layout engine, so the rail says so rather than implying Word parity.
 */
export class PageNavigator {
	readonly element: HTMLElement;
	private readonly heading = document.createElement('span');
	private readonly closeButton = document.createElement('button');
	private readonly message = document.createElement('p');
	private readonly list = document.createElement('div');
	private thumbs: Thumb[] = [];
	private observer?: IntersectionObserver;
	private open = false;
	private built = -1;
	private marked = 0;
	private labels = '';
	private state: PageNavigatorState = { active: false, status: null, locale: 'en' };

	constructor(private readonly host: PageNavigatorHost) {
		this.element = document.createElement('aside');
		this.element.className = 'dve-pages-rail';
		this.element.hidden = true;
		const head = document.createElement('div');
		head.className = 'dve-pages-head';
		this.closeButton.type = 'button';
		this.closeButton.className = 'dve-pages-close';
		this.closeButton.textContent = '×';
		this.closeButton.addEventListener('click', () => host.close());
		head.append(this.heading, this.closeButton);
		this.message.className = 'dve-pages-message';
		this.message.setAttribute('role', 'status');
		this.list.className = 'dve-pages-list';
		this.list.setAttribute('role', 'listbox');
		this.list.setAttribute('aria-orientation', 'vertical');
		this.list.addEventListener('click', (event) => {
			const index = this.indexOf(event.target);
			if (index >= 0) host.goTo(index + 1);
		});
		this.list.addEventListener('keydown', (event) => this.onKey(event));
		this.element.append(head, this.message, this.list);
		this.render();
	}

	get isOpen(): boolean {
		return this.open;
	}

	get pageCount(): number {
		return this.thumbs.length;
	}

	/** The 1-based page currently marked, or 0 when no thumbnails exist. */
	get currentPage(): number {
		return (
			this.thumbs.findIndex((thumb) => thumb.option.getAttribute('aria-selected') === 'true') + 1
		);
	}

	setOpen(open: boolean): void {
		this.open = open;
		this.element.hidden = !open;
		this.render();
	}

	/** Called after every layout, scroll and locale change; cheap when nothing moved. */
	sync(state: PageNavigatorState): void {
		this.state = state;
		this.render();
	}

	private render(): void {
		const { active, status, locale } = this.state;
		const ready = active && status !== null && status.total > 0;
		this.relabelChrome(locale, active, ready);
		if (!this.open || !ready) {
			if (this.thumbs.length) this.clear();
			return;
		}
		const version = this.host.version();
		if (version !== this.built || this.thumbs.length !== status.total) this.rebuild(version);
		const labels = `${locale}:${status.total}`;
		if (labels !== this.labels) this.relabelThumbs(locale, status.total);
		this.labels = labels;
		this.mark(status.current);
	}

	private relabelChrome(locale: EditorLocale, active: boolean, ready: boolean): void {
		const title = translate(locale, 'Page thumbnails');
		this.element.setAttribute('aria-label', title);
		this.list.setAttribute('aria-label', title);
		this.heading.textContent = translate(locale, 'Pages');
		const close = translate(locale, 'Close page thumbnails');
		this.closeButton.setAttribute('aria-label', close);
		this.closeButton.title = close;
		this.message.textContent = ready
			? ''
			: translate(locale, active ? 'nav.noPages' : 'nav.needsPrintLayout');
		this.message.hidden = ready;
		this.list.hidden = !ready;
		this.element.title = ready ? translate(locale, 'nav.approximate') : '';
	}

	private relabelThumbs(locale: EditorLocale, total: number): void {
		this.thumbs.forEach(({ option }, index) =>
			option.setAttribute('aria-label', formatPageStatus(locale, index + 1, total)),
		);
	}

	private clear(): void {
		this.observer?.disconnect();
		this.thumbs = [];
		this.built = -1;
		this.marked = 0;
		this.labels = '';
		this.list.replaceChildren();
	}

	private rebuild(version: number): void {
		this.clear();
		this.built = version;
		this.thumbs = this.host.pages().map((source, index) => this.createThumb(source, index));
		this.list.replaceChildren(...this.thumbs.map((thumb) => thumb.option));
		if (typeof IntersectionObserver === 'undefined') {
			this.thumbs.slice(0, EAGER_LIMIT).forEach((thumb) => this.draw(thumb));
			return;
		}
		this.observer = new IntersectionObserver(
			(entries) => {
				for (const entry of entries) {
					const thumb = this.thumbs.find((item) => item.option === entry.target);
					if (!thumb) continue;
					if (entry.isIntersecting) this.draw(thumb);
					else thumb.sheet.replaceChildren();
				}
			},
			{ root: this.list, rootMargin: '300px 0px' },
		);
		for (const thumb of this.thumbs) this.observer.observe(thumb.option);
	}

	private createThumb(source: HTMLElement, index: number): Thumb {
		const width = parseFloat(source.style.width) || DEFAULT_PAGE.width;
		const height = parseFloat(source.style.height) || DEFAULT_PAGE.height;
		const scale = THUMBNAIL_WIDTH / width;
		const option = document.createElement('div');
		option.className = 'dve-page-thumb';
		option.setAttribute('role', 'option');
		option.dataset.page = String(index + 1);
		option.tabIndex = -1;
		const sheet = document.createElement('div');
		sheet.className = 'dve-page-thumb-sheet';
		sheet.setAttribute('aria-hidden', 'true');
		sheet.style.width = `${THUMBNAIL_WIDTH}px`;
		sheet.style.height = `${Math.round(height * scale)}px`;
		const number = document.createElement('span');
		number.className = 'dve-page-thumb-number';
		number.setAttribute('aria-hidden', 'true');
		number.textContent = String(index + 1);
		option.append(sheet, number);
		return { option, sheet, source, scale };
	}

	private draw({ sheet, source, scale }: Thumb): void {
		if (sheet.firstChild) return;
		const clone = source.cloneNode(true) as HTMLElement;
		clone.removeAttribute('id');
		for (const node of clone.querySelectorAll('[id]')) node.removeAttribute('id');
		clone.inert = true;
		Object.assign(clone.style, { transformOrigin: '0 0', transform: `scale(${scale})` });
		sheet.replaceChildren(clone);
	}

	private hasFocus(): boolean {
		const root = this.element.getRootNode();
		const active =
			root instanceof ShadowRoot || root instanceof Document ? root.activeElement : null;
		return active !== null && this.list.contains(active);
	}

	/** Highlights `page` and, unless focus is inside the rail, makes it the tab stop and scrolls to it. */
	private mark(page: number): void {
		if (page === this.marked) return;
		this.marked = page;
		this.thumbs.forEach(({ option }, index) =>
			option.setAttribute('aria-selected', String(index + 1 === page)),
		);
		if (this.hasFocus()) return;
		this.setStop(page - 1);
		const option = this.thumbs[page - 1]?.option;
		if (option) this.reveal(option);
	}

	private setStop(index: number): void {
		this.thumbs.forEach(({ option }, i) => {
			option.tabIndex = i === index ? 0 : -1;
		});
	}

	/** Scrolls only the rail's own list, never an ancestor (`scrollIntoView` would move the page). */
	private reveal(option: HTMLElement): void {
		const { list } = this;
		if (option.offsetTop < list.scrollTop) list.scrollTop = option.offsetTop - 8;
		else if (option.offsetTop + option.offsetHeight > list.scrollTop + list.clientHeight)
			list.scrollTop = option.offsetTop + option.offsetHeight - list.clientHeight + 8;
	}

	private indexOf(target: EventTarget | null): number {
		const option = target instanceof Element ? target.closest('[role="option"]') : null;
		return this.thumbs.findIndex((thumb) => thumb.option === option);
	}

	private onKey(event: KeyboardEvent): void {
		const index = this.indexOf(event.target);
		if (index < 0) return;
		if (event.key === 'Enter' || event.key === ' ') {
			event.preventDefault();
			this.host.goTo(index + 1);
			return;
		}
		const last = this.thumbs.length - 1;
		const next =
			event.key === 'ArrowDown'
				? Math.min(last, index + 1)
				: event.key === 'ArrowUp'
					? Math.max(0, index - 1)
					: event.key === 'Home'
						? 0
						: event.key === 'End'
							? last
							: -1;
		if (next < 0) return;
		event.preventDefault();
		this.setStop(next);
		const thumb = this.thumbs[next];
		if (!thumb) return;
		thumb.option.focus({ preventScroll: true });
		this.reveal(thumb.option);
	}
}
