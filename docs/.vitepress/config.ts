import { defineConfig } from 'vitepress';

export default defineConfig({
	title: 'docx-viewer',
	description:
		'A browser-based Word document editing foundation with one document model, one web-component editor, and adapters for React, Vue, Angular, Svelte, Solid, and vanilla JavaScript.',
	base: '/docx-viewer/',
	cleanUrls: true,
	lastUpdated: true,
	ignoreDeadLinks: true,
	markdown: { theme: { light: 'vitesse-dark', dark: 'vitesse-dark' } },
	head: [
		['meta', { name: 'theme-color', content: '#c2431f' }],
		['link', { rel: 'preconnect', href: 'https://fonts.googleapis.com' }],
		['link', { rel: 'preconnect', href: 'https://fonts.gstatic.com', crossorigin: '' }],
		[
			'link',
			{
				rel: 'stylesheet',
				href: 'https://fonts.googleapis.com/css2?family=Bricolage+Grotesque:opsz,wght@12..96,400..800&family=IBM+Plex+Mono:ital,wght@0,400;0,500;1,400&display=swap',
			},
		],
		['meta', { property: 'og:title', content: 'docx-viewer documentation' }],
		[
			'meta',
			{
				property: 'og:description',
				content: 'A shared Word editing foundation and framework adapters.',
			},
		],
	],
	themeConfig: {
		siteTitle: 'docx-viewer',
		darkModeSwitchLabel: 'Appearance',
		lightModeSwitchTitle: 'Switch to light theme',
		darkModeSwitchTitle: 'Switch to dark theme',
		nav: [
			{ text: 'Developer Guide', link: '/architecture', activeMatch: '/architecture' },
			{ text: 'User Guide', link: '/editing' },
			{
				text: 'Packages',
				items: [
					{ text: 'Core engine', link: '/architecture' },
					{ text: 'React', link: '/frameworks/react' },
					{ text: 'Vue 3', link: '/frameworks/vue' },
					{ text: 'Angular', link: '/frameworks/angular' },
					{ text: 'Vanilla JavaScript', link: '/frameworks/vanilla' },
					{ text: 'Svelte', link: '/frameworks/svelte' },
					{ text: 'Solid', link: '/frameworks/solid' },
				],
			},
			{ text: 'Releases', link: '/releasing' },
			{
				text: 'Resources',
				items: [
					{ text: 'Live demos', link: '/#live-demo' },
					{ text: 'Support roadmap', link: '/parity-roadmap' },
					{ text: 'Shared format code', link: '/reuse-audit' },
					{ text: 'Release policy', link: '/releasing' },
				],
			},
		],
		sidebar: [
			{
				text: 'Documentation',
				items: [
					{ text: 'Overview', link: '/' },
					{ text: 'Architecture', link: '/architecture' },
					{ text: 'Framework bindings', link: '/bindings' },
					{ text: 'Model units', link: '/model-units' },
					{ text: 'Editing text', link: '/editing' },
					{ text: 'Collaboration', link: '/collaboration' },
					{ text: 'Support roadmap', link: '/parity-roadmap' },
					{ text: 'Shared format code', link: '/reuse-audit' },
					{ text: 'Package releases', link: '/releasing' },
				],
			},
			{
				text: 'Framework guides',
				items: [
					{ text: 'React', link: '/frameworks/react' },
					{ text: 'Vue', link: '/frameworks/vue' },
					{ text: 'Angular', link: '/frameworks/angular' },
					{ text: 'Vanilla JS', link: '/frameworks/vanilla' },
					{ text: 'Svelte', link: '/frameworks/svelte' },
					{ text: 'Solid', link: '/frameworks/solid' },
				],
			},
		],
		socialLinks: [{ icon: 'github', link: 'https://github.com/ChristopherVR/docx-viewer' }],
		search: { provider: 'local' },
		footer: {
			message: 'Released under the Apache-2.0 License.',
			copyright: 'Copyright © 2026 ChristopherVR',
		},
	},
});
