import { defineConfig } from 'vitest/config';
import { fileURLToPath } from 'node:url';
export default defineConfig({
	resolve: {
		alias: [
			{
				find: /^@christophervr\/docx-core\/embedded$/,
				replacement: fileURLToPath(new URL('./packages/core/src/embedded.ts', import.meta.url)),
			},
			...['core', 'legacy', 'web-component', 'document'].map((name) => ({
				find: new RegExp(`^@christophervr/docx-${name}$`),
				replacement: fileURLToPath(new URL(`./packages/${name}/src/index.ts`, import.meta.url)),
			})),
		],
	},
	test: { include: ['packages/**/*.test.ts', 'packages/**/*.test.tsx'], environment: 'node' },
});
