import { existsSync, readdirSync } from 'node:fs';
import { join } from 'node:path';

const HEADLESS_SHELL: Record<string, string[]> = {
	linux: ['chrome-linux', 'headless_shell'],
	darwin: ['chrome-mac', 'headless_shell'],
	win32: ['chrome-win', 'headless_shell.exe'],
};

const FULL_CHROMIUM: Record<string, string[]> = {
	linux: ['chrome-linux', 'chrome'],
	darwin: ['chrome-mac', 'Chromium.app', 'Contents', 'MacOS', 'Chromium'],
	win32: ['chrome-win', 'chrome.exe'],
};

/** Newest installed chromium_headless_shell-* / chromium-* executable under a browsers directory. */
export function findInstalledChromium(
	browsersPath: string,
	platform: string = process.platform,
): string | undefined {
	if (!existsSync(browsersPath)) return undefined;
	const names = readdirSync(browsersPath);
	const candidates = [
		{ prefix: 'chromium_headless_shell-', parts: HEADLESS_SHELL[platform] },
		{ prefix: 'chromium-', parts: FULL_CHROMIUM[platform] },
	];
	for (const { prefix, parts } of candidates) {
		if (!parts) continue;
		const found = names
			.filter((n) => n.startsWith(prefix) && /^\d+$/.test(n.slice(prefix.length)))
			.sort((a, b) => Number(b.slice(prefix.length)) - Number(a.slice(prefix.length)))
			.map((n) => join(browsersPath, n, ...parts))
			.find((p) => existsSync(p));
		if (found) return found;
	}
	return undefined;
}

/**
 * Explicit PLAYWRIGHT_CHROMIUM_EXECUTABLE wins. Otherwise the bundled build is used when installed;
 * only if it is missing do we fall back to the newest build under PLAYWRIGHT_BROWSERS_PATH.
 */
export function resolveChromiumExecutable(
	bundled: string,
	env: NodeJS.ProcessEnv = process.env,
): string | undefined {
	const override = env.PLAYWRIGHT_CHROMIUM_EXECUTABLE;
	if (override) return override;
	if (existsSync(bundled)) return undefined;
	const dir = env.PLAYWRIGHT_BROWSERS_PATH;
	return dir && dir !== '0' ? findInstalledChromium(dir) : undefined;
}
