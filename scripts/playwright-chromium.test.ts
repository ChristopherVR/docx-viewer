import { mkdirSync, mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { findInstalledChromium, resolveChromiumExecutable } from './playwright-chromium';

function fakeCache(): string {
	const root = mkdtempSync(join(tmpdir(), 'pw-'));
	for (const [dir, exe] of [
		['chromium_headless_shell-1100', 'headless_shell'],
		['chromium_headless_shell-1194', 'headless_shell'],
		['chromium-1200', 'chrome'],
	] as const) {
		mkdirSync(join(root, dir, 'chrome-linux'), { recursive: true });
		writeFileSync(join(root, dir, 'chrome-linux', exe), '');
	}
	return root;
}

describe('chromium resolution', () => {
	it('prefers the newest headless shell', () => {
		const root = fakeCache();
		expect(findInstalledChromium(root, 'linux')).toBe(
			join(root, 'chromium_headless_shell-1194', 'chrome-linux', 'headless_shell'),
		);
	});
	it('honours the explicit override and leaves an installed bundle alone', () => {
		const root = fakeCache();
		const bundled = join(root, 'chromium-1200', 'chrome-linux', 'chrome');
		expect(resolveChromiumExecutable(bundled, { PLAYWRIGHT_CHROMIUM_EXECUTABLE: '/x' })).toBe('/x');
		expect(resolveChromiumExecutable(bundled, { PLAYWRIGHT_BROWSERS_PATH: root })).toBeUndefined();
	});
	it('falls back only when the bundled executable is missing', () => {
		const root = fakeCache();
		expect(resolveChromiumExecutable('/nope', { PLAYWRIGHT_BROWSERS_PATH: root })).toContain(
			'1194',
		);
	});
});
