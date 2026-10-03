import type { PageContext } from './backstage-pages';

export interface BackstageOptions {
	locale: string;
	theme: string;
	author: string;
	spellCheck?: boolean;
	showRuler?: boolean;
	showMarks?: boolean;
	showThumbnails?: boolean;
}
export type BackstageOptionKey = keyof BackstageOptions;

/** Editor preferences share the same live controls as the ribbon. */
export function renderOptions({ handlers, t, content }: PageContext): void {
	const current = handlers.options();
	const heading = (text: string, tag: 'h2' | 'h3') => {
		const el = document.createElement(tag);
		el.textContent = t(text);
		return el;
	};
	const field = (text: string, control: HTMLElement) => {
		const label = document.createElement('label');
		label.className = 'dve-backstage-field';
		const caption = document.createElement('span');
		caption.textContent = t(text);
		control.setAttribute('aria-label', t(text));
		label.append(caption, control);
		return label;
	};
	const select = (key: 'locale' | 'theme', values: ReadonlyArray<readonly [string, string]>) => {
		const el = document.createElement('select');
		for (const [value, text] of values) el.append(new Option(t(text), value));
		el.value = current[key];
		el.addEventListener('change', () => handlers.setOption(key, el.value));
		return el;
	};
	const author = document.createElement('input');
	author.type = 'text';
	author.value = current.author;
	author.addEventListener('change', () => handlers.setOption('author', author.value));
	const language = select('locale', [
		['en', 'English'],
		['fr', 'Français'],
		['de', 'Deutsch'],
		['es', 'Español'],
		['zh-CN', '简体中文'],
	]);
	const theme = select('theme', [
		['auto', 'Automatic'],
		['light', 'Light'],
		['dark', 'Dark'],
	]);
	if (!theme.value) theme.value = 'auto';
	const controls: HTMLElement[] = [];
	for (const [key, label] of [
		['spellCheck', 'Spelling'],
		['showRuler', 'Ruler'],
		['showMarks', 'Show paragraph marks'],
		['showThumbnails', 'Page thumbnails'],
	] as const) {
		if (current[key] === undefined) continue;
		const input = document.createElement('input');
		input.type = 'checkbox';
		input.checked = current[key];
		input.addEventListener('change', () => handlers.setOption(key, String(input.checked)));
		const row = field(label, input);
		row.classList.add('dve-backstage-toggle');
		controls.push(row);
	}
	const note = document.createElement('p');
	note.className = 'dve-backstage-muted';
	note.textContent = t('Options apply to this editor and are not stored between sessions.');
	content.replaceChildren(
		heading('Options', 'h2'),
		heading('General', 'h3'),
		field('Display language', language),
		field('Theme', theme),
		field('Author name', author),
		...(controls.length ? [heading('View', 'h3'), ...controls] : []),
		note,
	);
}
