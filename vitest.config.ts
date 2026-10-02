import { defineConfig } from 'vitest/config';
import { fileURLToPath } from 'node:url';
export default defineConfig({
	resolve: {
		alias: [
			{
				find: /^solid-js\/web$/,
				replacement: fileURLToPath(
					new URL('./node_modules/solid-js/web/dist/web.js', import.meta.url),
				),
			},
			{
				find: /^solid-js$/,
				replacement: fileURLToPath(
					new URL('./node_modules/solid-js/dist/solid.js', import.meta.url),
				),
			},
			{
				find: /^docx-core\/embedded$/,
				replacement: fileURLToPath(new URL('./packages/core/src/embedded.ts', import.meta.url)),
			},
			{
				find: /^docx-bindings\/(react|vue|angular|solid|common)$/,
				replacement: fileURLToPath(new URL('./packages/bindings/src/$1', import.meta.url)),
			},
			...['core', 'web-component', 'bindings'].map((name) => ({
				find: new RegExp(`^docx-${name}$`),
				replacement: fileURLToPath(new URL(`./packages/${name}/src/index.ts`, import.meta.url)),
			})),
		],
	},
	test: {
		include: ['packages/**/*.test.ts', 'packages/**/*.test.tsx', 'scripts/**/*.test.ts'],
		environment: 'node',
	},
});
