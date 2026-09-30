/** Word's own KeyTips for the commands this ribbon shares with it, by the control's English label. */
const WORD_KEYS: Record<string, string> = {
	Paste: 'V',
	Cut: 'X',
	Copy: 'C',
	'Format painter': 'FP',
	'Font family': 'FF',
	'Font size': 'FS',
	'Grow font': 'FG',
	'Shrink font': 'FK',
	'Change case': 'AA',
	'Clear formatting': 'E',
	Bold: '1',
	Italic: '2',
	Underline: '3',
	Strikethrough: '4',
	Subscript: '5',
	Superscript: '6',
	'Text highlight': 'I',
	'Font color': 'FC',
	'Bulleted list': 'U',
	'Numbered list': 'N',
	'Decrease indent': 'AO',
	'Increase indent': 'AI',
	Sort: 'SO',
	'Show paragraph marks': '8',
	'Align left': 'AL',
	'Align center': 'AC',
	'Align right': 'AR',
	Justify: 'AJ',
	'Line spacing': 'K',
	Shading: 'H',
	Borders: 'B',
	'Find and replace': 'FD',
	Replace: 'R',
	'Select all': 'SL',
	Table: 'T',
	Pictures: 'P',
	Link: 'K',
	Bookmark: 'K',
	Header: 'H',
	Footer: 'O',
	'Page number': 'NU',
	Margins: 'M',
	'Page size': 'SZ',
	Orientation: 'O',
	Columns: 'J',
	'Table of contents': 'T',
	'Update table': 'U',
	Footnote: 'F',
	Endnote: 'E',
	Spelling: 'C',
	'Word count': 'W',
	'Track changes': 'G',
	Zoom: 'Q',
};

export interface KeyTipTarget {
	key: string;
	element: HTMLElement;
	/** What pressing the tip does: open the control, or click it. */
	activate: () => void;
}

const nameOf = (el: HTMLElement) =>
	el.dataset.localearialabel ?? el.getAttribute('aria-label') ?? '';

/** The control a badge belongs on, and how to operate it. */
function targetOf(el: HTMLElement): KeyTipTarget | null {
	const wrapper = el.closest<HTMLElement>('.ribbon-menu, .ribbon-combo') ?? el;
	if (!wrapper.getClientRects().length) return null;
	const label = nameOf(el);
	if (!label) return null;
	const activate = () => {
		if (el instanceof HTMLSelectElement && !wrapper.classList.contains('ribbon-menu-gallery')) {
			el.focus();
			try {
				el.showPicker?.();
			} catch {
				// Some browsers only allow showPicker from a pointer gesture; focus is enough then.
			}
		} else if (el instanceof HTMLInputElement) el.focus();
		else wrapper.click();
	};
	return { key: WORD_KEYS[label] ?? '', element: wrapper, activate };
}

/**
 * KeyTips for the visible controls of `panel`: Word's own where this ribbon has the same command,
 * otherwise two letters from the label. No tip is the start of another, so typing one is never
 * ambiguous.
 */
export function collectKeyTips(panel: HTMLElement): KeyTipTarget[] {
	const seen = new Set<HTMLElement>();
	const targets: KeyTipTarget[] = [];
	const controls = panel.querySelectorAll<HTMLElement>(
		'button[aria-label]:not(.ribbon-launcher):not([data-split-caret]):not(.ribbon-combo-caret):not(.style-tile), select[aria-label], .ribbon-combo input[aria-label], button.ribbon-overflow-button',
	);
	for (const control of controls) {
		const target = targetOf(control);
		if (!target || seen.has(target.element)) continue;
		seen.add(target.element);
		targets.push(target);
	}
	// Explicit keys may collide with each other (Word's differ per tab); the later one is re-derived.
	const used: string[] = [];
	for (const target of targets) {
		if (target.key && !used.some((u) => u.startsWith(target.key) || target.key.startsWith(u))) {
			used.push(target.key);
		} else target.key = '';
	}
	for (const target of targets.filter((t) => !t.key)) {
		const label = nameOf(target.element.querySelector('[aria-label]') ?? target.element)
			.toUpperCase()
			.replace(/[^A-Z0-9]/g, '');
		const source = label || 'ZZ';
		let key = '';
		for (let a = 0; a < source.length && !key; a++)
			for (let b = a + 1; b < source.length && !key; b++) {
				const candidate = source[a]! + source[b]!;
				if (!used.some((u) => u.startsWith(candidate) || candidate.startsWith(u))) key = candidate;
			}
		for (let n = 10; !key; n++) if (!used.includes(String(n))) key = String(n);
		used.push(key);
		target.key = key;
	}
	return targets;
}

/**
 * Shows badges over the controls of `panel` and runs the one whose letters are typed. Escape (or a
 * key with no match) ends it; `onDone` runs after any end. Returns a function that stops it.
 */
export function showCommandTips(root: HTMLElement, panel: HTMLElement, onDone: () => void) {
	const targets = collectKeyTips(panel);
	const badges = targets.map((target) => {
		const badge = document.createElement('span');
		badge.className = 'dve-keytip dve-keytip-command';
		badge.textContent = target.key;
		badge.setAttribute('aria-hidden', 'true');
		const box = target.element.getBoundingClientRect();
		badge.style.position = 'fixed';
		badge.style.left = `${box.left + box.width / 2}px`;
		badge.style.top = `${box.bottom - 10}px`;
		root.append(badge);
		return { badge, target };
	});
	let typed = '';
	const stop = () => {
		for (const { badge } of badges) badge.remove();
		root.removeEventListener('keydown', onKey, true);
		root.removeEventListener('pointerdown', stop, true);
		onDone();
	};
	function onKey(event: KeyboardEvent) {
		if (event.key === 'Alt' || event.key === 'Shift') return;
		event.preventDefault();
		event.stopPropagation();
		if (event.key === 'Escape' || event.ctrlKey || event.metaKey) return stop();
		typed += event.key.toUpperCase();
		const matches = badges.filter(({ target }) => target.key.startsWith(typed));
		const exact = matches.find(({ target }) => target.key === typed);
		if (!matches.length) return stop();
		for (const { badge, target } of badges) badge.hidden = !target.key.startsWith(typed);
		if (exact) {
			stop();
			exact.target.activate();
		}
	}
	root.addEventListener('keydown', onKey, true);
	root.addEventListener('pointerdown', stop, true);
	return stop;
}
