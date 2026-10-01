import { defineConfig } from 'vite';
import { fileURLToPath } from 'node:url';
import { svelte } from '@sveltejs/vite-plugin-svelte';
import tailwindcss from '@tailwindcss/vite';
export default defineConfig({
	plugins: [tailwindcss(), svelte()],
	build: {
		target: 'es2022',
		rollupOptions: {
			input: {
				editor: fileURLToPath(new URL('./demos/demo-vanilla/index.html', import.meta.url)),
				collaboration: fileURLToPath(
					new URL('./demos/demo-vanilla/collaboration.html', import.meta.url),
				),
			},
		},
	},
	resolve: {
		alias: [
			{
				find: /^@christophervr\/docx-core\/embedded$/,
				replacement: fileURLToPath(new URL('./packages/core/src/embedded.ts', import.meta.url)),
			},
			...['core', 'legacy', 'web-component', 'document', 'layout'].map((name) => ({
				find: new RegExp(`^@christophervr/docx-${name}$`),
				replacement: fileURLToPath(new URL(`./packages/${name}/src/index.ts`, import.meta.url)),
			})),
		],
	},
	// Pre-bundle every framework the demo can mount on demand. Otherwise the first visit to a
	// framework makes Vite discover new dependencies, re-optimize and reload the page mid-session.
	optimizeDeps: {
		include: [
			'@angular/compiler',
			'@angular/core',
			'@angular/common',
			'@angular/platform-browser',
			'@xmldom/xmldom',
			'jszip',
			'prosemirror-commands',
			'prosemirror-history',
			'prosemirror-keymap',
			'prosemirror-model',
			'prosemirror-state',
			'prosemirror-view',
			'react',
			'react-dom/client',
			'rxjs',
			'solid-js',
			'solid-js/web',
			'svelte',
			'vue',
		],
	},
	server: { fs: { allow: [fileURLToPath(new URL('../', import.meta.url))] } },
});
