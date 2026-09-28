/** Shared demo/Pages theme preference. */
export function initTheme() {
	const get = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
	const themeKey = 'vitepress-theme-appearance';
	const savedTheme = localStorage.getItem(themeKey);
	const systemTheme = matchMedia('(prefers-color-scheme: dark)');
	const resolveTheme = (preference: string | null): 'light' | 'dark' =>
		preference === 'dark' || preference === 'light'
			? preference
			: systemTheme.matches
				? 'dark'
				: 'light';
	// Drive the editor's own `theme` property (chrome tokens) alongside the demo page theme. Editors
	// mount lazily, so also sync any docx-editor added later.
	const syncEditors = () => {
		const theme = document.documentElement.dataset.theme as 'light' | 'dark';
		for (const editor of document.querySelectorAll<HTMLElement & { theme: string }>('docx-editor'))
			if (editor.theme !== theme) editor.theme = theme;
	};
	new MutationObserver(syncEditors).observe(document.body, { childList: true, subtree: true });
	const applyTheme = (theme: 'light' | 'dark') => {
		document.documentElement.dataset.theme = theme;
		syncEditors();
		const toggle = get<HTMLButtonElement>('theme-toggle');
		const dark = theme === 'dark';
		toggle.textContent = dark ? 'Light' : 'Dark';
		toggle.setAttribute('aria-label', `Switch to ${dark ? 'light' : 'dark'} theme`);
		toggle.setAttribute('aria-pressed', String(dark));
	};
	applyTheme(resolveTheme(savedTheme));
	// VitePress changes this preference in the parent window. Update the embedded
	// editor in place so switching appearance never requires reloading user edits.
	window.addEventListener('storage', (event) => {
		if (event.key === themeKey || event.key === null) applyTheme(resolveTheme(event.newValue));
	});
	systemTheme.addEventListener('change', () => {
		applyTheme(resolveTheme(localStorage.getItem(themeKey)));
	});
	get<HTMLButtonElement>('theme-toggle').addEventListener('click', () => {
		const next = document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark';
		localStorage.setItem(themeKey, next);
		applyTheme(next);
	});
}
