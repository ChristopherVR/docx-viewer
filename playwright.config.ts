import { defineConfig } from '@playwright/test';
export default defineConfig({
	testDir: './tests',
	fullyParallel: false,
	workers: 1,
	use: { baseURL: 'http://127.0.0.1:4180', headless: true },
	webServer: {
		// Test the production build: the dev server issues hundreds of unbundled module requests per
		// page, which exhausts ephemeral ports on Windows (net::ERR_ADDRESS_IN_USE) over a full run.
		command:
			'bun run build && bun x vite preview demos/demo-vanilla --config vite.config.ts --host 127.0.0.1 --port 4180 --strictPort',
		url: 'http://127.0.0.1:4180',
		reuseExistingServer: !process.env.CI,
		timeout: 180_000,
	},
	reporter: 'list',
});
