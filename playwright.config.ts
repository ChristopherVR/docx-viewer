import { defineConfig } from '@playwright/test';
export default defineConfig({
	testDir: './tests',
	fullyParallel: false,
	workers: 1,
	use: { baseURL: 'http://127.0.0.1:4180', headless: true },
	webServer: {
		command: 'bun run demo --port 4180',
		url: 'http://127.0.0.1:4180',
		reuseExistingServer: !process.env.CI,
	},
	reporter: 'list',
});
