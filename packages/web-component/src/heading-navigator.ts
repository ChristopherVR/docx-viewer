import { tocEntries, type DocumentModel } from 'docx-core';
import type { EditorView } from 'prosemirror-view';
import { translate, type EditorLocale } from './localization';

export interface HeadingNavigatorHost {
	model(): DocumentModel;
	view(): EditorView | undefined;
	goTo(blockId: string): void;
	close(): void;
}

/** Body headings in document order; shared TOC style resolution handles inherited heading styles. */
export class HeadingNavigator {
	readonly element = document.createElement('aside');
	private readonly title = document.createElement('strong');
	private readonly closeButton = document.createElement('button');
	private readonly tree = document.createElement('div');
	private readonly empty = document.createElement('p');
	private readonly collapsed = new Set<string>();
	private signature = '';
	private locale: EditorLocale = 'en';
	private entries: ReturnType<typeof tocEntries> = [];
	private selected = -1;

	constructor(private readonly host: HeadingNavigatorHost) {
		this.element.className = 'dve-heading-rail';
		this.element.hidden = true;
		const style = document.createElement('style');
		style.textContent = `.dve-heading-rail{flex:0 0 230px;min-width:0;max-width:35vw;display:flex;flex-direction:column;background:var(--surface,#fff);border-right:1px solid var(--line,#ddd);font:13px/1.4 'Segoe UI',sans-serif;color:var(--ink,#222)}
		.dve-heading-rail[hidden]{display:none}.dve-heading-head{display:flex;justify-content:space-between;align-items:center;padding:12px}.dve-heading-head button{border:0;background:none;color:inherit;font-size:20px;cursor:pointer}.dve-heading-tree{overflow:auto;padding:4px 8px;flex:1}.dve-heading-item{display:flex;align-items:center;gap:4px;min-height:32px;border-radius:3px;padding:3px 5px;cursor:pointer;outline-offset:-2px}.dve-heading-item[hidden]{display:none}.dve-heading-item:hover{background:var(--hover,#f2f4f7)}.dve-heading-item[aria-selected=true]{background:var(--blue-soft,#e5f0fc);color:var(--blue,#1769aa)}.dve-heading-toggle{border:0;background:none;color:inherit;flex:0 0 16px;padding:0;cursor:pointer}.dve-heading-text{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.dve-heading-empty{padding:0 12px;color:var(--muted,#666)}`;
		const head = document.createElement('div');
		head.className = 'dve-heading-head';
		this.closeButton.type = 'button';
		this.closeButton.textContent = '×';
		this.closeButton.addEventListener('click', () => host.close());
		head.append(this.title, this.closeButton);
		this.tree.className = 'dve-heading-tree';
		this.tree.setAttribute('role', 'tree');
		this.empty.className = 'dve-heading-empty';
		this.empty.setAttribute('role', 'status');
		this.tree.addEventListener('keydown', (event) => this.onKey(event));
		this.element.append(style, head, this.empty, this.tree);
	}

	get isOpen(): boolean {
		return !this.element.hidden;
	}
	setOpen(open: boolean): void {
		this.element.hidden = !open;
		if (open) this.sync(this.locale);
	}

	sync(locale: EditorLocale): void {
		this.locale = locale;
		if (!this.isOpen) return;
		this.element.setAttribute('aria-label', translate(locale, 'Navigation pane'));
		this.title.textContent = translate(locale, 'Headings');
		this.tree.setAttribute('aria-label', translate(locale, 'Headings'));
		this.closeButton.setAttribute('aria-label', translate(locale, 'Close navigation pane'));
		this.entries = tocEntries(this.host.model(), ' TOC \\o "1-9" ');
		this.empty.textContent = translate(
			locale,
			'Apply a heading style to see document headings here.',
		);
		this.empty.hidden = this.entries.length > 0;
		const signature = JSON.stringify([locale, this.entries]);
		if (signature !== this.signature) {
			this.signature = signature;
			this.rebuild();
		}
		this.mark();
	}

	private rebuild(): void {
		const focused = this.tree.querySelector<HTMLElement>('[role=treeitem]:focus')?.dataset.id;
		this.tree.replaceChildren(
			...this.entries.map((entry, index) => {
				const item = document.createElement('div');
				item.className = 'dve-heading-item';
				item.dataset.id = entry.blockId;
				item.dataset.index = String(index);
				item.setAttribute('role', 'treeitem');
				item.setAttribute('aria-level', String(entry.level));
				item.style.paddingLeft = `${5 + (entry.level - 1) * 14}px`;
				item.tabIndex = -1;
				const toggle = document.createElement('button');
				toggle.type = 'button';
				toggle.className = 'dve-heading-toggle';
				toggle.tabIndex = -1;
				const hasChildren = (this.entries[index + 1]?.level ?? 0) > entry.level;
				if (hasChildren) {
					item.setAttribute('aria-expanded', String(!this.collapsed.has(entry.blockId)));
					toggle.addEventListener('click', (event) => {
						event.stopPropagation();
						this.toggle(index);
					});
				} else {
					toggle.disabled = true;
					toggle.setAttribute('aria-hidden', 'true');
				}
				const text = document.createElement('span');
				text.className = 'dve-heading-text';
				text.textContent = entry.text;
				item.title = entry.text;
				item.append(toggle, text);
				item.addEventListener('click', () => this.host.goTo(entry.blockId));
				return item;
			}),
		);
		this.updateVisibility();
		if (focused)
			this.items()
				.find((item) => item.dataset.id === focused)
				?.focus();
	}

	private items(): HTMLElement[] {
		return [...this.tree.querySelectorAll<HTMLElement>('[role=treeitem]')];
	}
	private toggle(index: number): void {
		const entry = this.entries[index];
		if (!entry || !this.items()[index]?.hasAttribute('aria-expanded')) return;
		if (this.collapsed.has(entry.blockId)) this.collapsed.delete(entry.blockId);
		else this.collapsed.add(entry.blockId);
		this.updateVisibility();
		this.mark();
	}
	private updateVisibility(): void {
		let hiddenBelow = 10;
		this.items().forEach((item, index) => {
			const entry = this.entries[index]!;
			if (entry.level <= hiddenBelow) hiddenBelow = 10;
			item.hidden = entry.level > hiddenBelow;
			const collapsed = this.collapsed.has(entry.blockId);
			if (!item.hidden && collapsed) hiddenBelow = entry.level;
			if (item.hasAttribute('aria-expanded')) {
				item.setAttribute('aria-expanded', String(!collapsed));
				const toggle = item.querySelector('button')!;
				toggle.textContent = collapsed ? '▸' : '▾';
				toggle.setAttribute(
					'aria-label',
					translate(this.locale, collapsed ? 'Expand heading' : 'Collapse heading'),
				);
			}
		});
	}
	private mark(): void {
		const view = this.host.view();
		let currentId = '';
		const ids = new Set(this.entries.map((entry) => entry.blockId));
		view?.state.doc.forEach((node, pos) => {
			if (pos <= view.state.selection.from && ids.has(String(node.attrs.id)))
				currentId = String(node.attrs.id);
		});
		this.selected = this.entries.findIndex((entry) => entry.blockId === currentId);
		const items = this.items();
		let stop = items[this.selected];
		if (!stop || stop.hidden) stop = items.find((item) => !item.hidden);
		items.forEach((item, index) => {
			item.setAttribute('aria-selected', String(index === this.selected));
			item.tabIndex = item === stop ? 0 : -1;
		});
	}
	private onKey(event: KeyboardEvent): void {
		const item =
			event.target instanceof Element ? event.target.closest<HTMLElement>('[role=treeitem]') : null;
		if (!item) return;
		const index = Number(item.dataset.index);
		if (event.key === 'Enter' || event.key === ' ') {
			event.preventDefault();
			this.host.goTo(this.entries[index]!.blockId);
			return;
		}
		if (event.key === 'ArrowLeft' || event.key === 'ArrowRight') {
			const expanded = item.getAttribute('aria-expanded');
			if (
				(expanded === 'true' && event.key === 'ArrowLeft') ||
				(expanded === 'false' && event.key === 'ArrowRight')
			)
				this.toggle(index);
			else if (event.key === 'ArrowLeft')
				this.items()
					.slice(0, index)
					.reverse()
					.find(
						(parent) =>
							!parent.hidden &&
							Number(parent.getAttribute('aria-level')) < Number(item.getAttribute('aria-level')),
					)
					?.focus();
			else if (expanded === 'true') this.items()[index + 1]?.focus();
			event.preventDefault();
			return;
		}
		const visible = this.items().filter((entry) => !entry.hidden);
		const at = visible.indexOf(item);
		const next =
			event.key === 'ArrowDown'
				? Math.min(visible.length - 1, at + 1)
				: event.key === 'ArrowUp'
					? Math.max(0, at - 1)
					: event.key === 'Home'
						? 0
						: event.key === 'End'
							? visible.length - 1
							: -1;
		if (next >= 0) {
			event.preventDefault();
			visible[next]?.focus();
		}
		if (event.key === 'Escape') {
			event.preventDefault();
			this.host.close();
		}
	}
}
