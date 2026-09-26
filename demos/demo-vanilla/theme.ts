/** Shared demo/Pages theme preference. */
export function initTheme() {
	const get = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
	const themeKey = 'vitepress-theme-appearance';
	const savedTheme = localStorage.getItem(themeKey);
	const applyTheme = (theme: 'light' | 'dark') => {
		document.documentElement.dataset.theme = theme;
		const toggle = get<HTMLButtonElement>('theme-toggle');
		const dark = theme === 'dark';
		toggle.textContent = dark ? 'Light' : 'Dark';
		toggle.setAttribute('aria-label', `Switch to ${dark ? 'light' : 'dark'} theme`);
		toggle.setAttribute('aria-pressed', String(dark));
	};
	applyTheme(
		savedTheme === 'dark' || savedTheme === 'light'
			? savedTheme
			: matchMedia('(prefers-color-scheme: dark)').matches
				? 'dark'
				: 'light',
	);
	get<HTMLButtonElement>('theme-toggle').addEventListener('click', () => {
		const next = document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark';
		localStorage.setItem(themeKey, next);
		applyTheme(next);
	});
}
