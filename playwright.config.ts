import { chromium, defineConfig } from '@playwright/test';
import { resolveChromiumExecutable } from './scripts/playwright-chromium';

const port = Number(process.env.PLAYWRIGHT_PORT ?? 4180);
const origin = `http://127.0.0.1:${port}`;
const executablePath = resolveChromiumExecutable(chromium.executablePath());

export default defineConfig({
	testDir: './tests',
	fullyParallel: false,
	workers: 1,
	use: {
		baseURL: origin,
		headless: true,
		...(executablePath ? { launchOptions: { executablePath } } : {}),
	},
	webServer: {
		// Test the production build: the dev server issues hundreds of unbundled module requests per
		// page, which exhausts ephemeral ports on Windows (net::ERR_ADDRESS_IN_USE) over a full run.
		command: `bun run build && bun x vite preview demos/demo-vanilla --config vite.config.ts --host 127.0.0.1 --port ${port} --strictPort`,
		url: origin,
		reuseExistingServer: !process.env.CI,
		timeout: 180_000,
	},
	reporter: 'list',
});
