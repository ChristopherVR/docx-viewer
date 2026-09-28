import { expect, test, type Page } from '@playwright/test';
import { mkdirSync } from 'node:fs';

const LOCALES = ['en', 'fr', 'de', 'es', 'zh-CN'] as const;
const SHOTS =
	process.env.LOCALE_SHOT_DIR ??
	'/tmp/claude-0/-home-user-docx-viewer/cec0184c-692a-5fed-937f-11818dd0a964/scratchpad/ui';
const TABS = ['home', 'insert'] as const;

/** Names every visible ribbon control in `tab` whose text is clipped or wider than its box. */
async function clippedControls(page: Page, tab: string): Promise<string[]> {
	return page.evaluate((tabId) => {
		const root = document.querySelector('docx-editor')!.shadowRoot!;
		const panel = root.querySelector<HTMLElement>(`#dve-panel-${tabId}`)!;
		const canvas = document.createElement('canvas').getContext('2d')!;
		const bad: string[] = [];
		const name = (el: HTMLElement) =>
			el.getAttribute('aria-label') ||
			el.dataset.localearialabel ||
			el.textContent?.trim() ||
			el.tagName;
		for (const el of panel.querySelectorAll<HTMLElement>('button, select, .ribbon-group')) {
			if (el.getClientRects().length === 0) continue;
			if (el instanceof HTMLSelectElement) {
				// A select's scrollWidth ignores clipped text, so measure the selected label itself.
				const style = getComputedStyle(el);
				canvas.font = `${style.fontWeight} ${style.fontSize} ${style.fontFamily}`;
				const text = el.selectedOptions[0]?.textContent ?? '';
				const room =
					el.clientWidth - parseFloat(style.paddingLeft) - parseFloat(style.paddingRight) - 20;
				if (canvas.measureText(text).width > room + 1) {
					// Too long for its box: it must truncate with an ellipsis and expose the full text.
					el.dispatchEvent(new Event('pointerover', { bubbles: true }));
					if (style.textOverflow !== 'ellipsis' || el.title !== text.trim())
						bad.push(`select ${name(el)}: "${text}" is clipped without ellipsis and tooltip`);
				}
			} else if (el.scrollWidth > el.clientWidth + 1) {
				bad.push(`${el.tagName.toLowerCase()} ${name(el)}: ${el.scrollWidth} > ${el.clientWidth}`);
			}
			if (el instanceof HTMLButtonElement && el.textContent?.trim() && !el.title && !el.ariaLabel)
				bad.push(`button without tooltip: ${el.textContent.trim()}`);
		}
		return bad;
	}, tab);
}

for (const locale of LOCALES)
	test(`ribbon labels fit their controls in ${locale}`, async ({ page }) => {
		mkdirSync(SHOTS, { recursive: true });
		await page.setViewportSize({ width: 1280, height: 900 });
		await page.goto(`/?locale=${locale}`);
		await page.locator('html[data-demo-ready="true"]').waitFor({ state: 'attached' });
		await page.locator('#sample').click();
		const editor = page.locator('docx-editor');
		await expect(editor).toHaveAttribute('locale', locale);
		await expect(editor.locator('.ProseMirror')).toContainText('Document title');
		for (const tab of TABS) {
			await editor.locator(`#dve-tab-${tab}`).click();
			expect(await clippedControls(page, tab), `${locale} ${tab}`).toEqual([]);
			if (tab === 'home') await page.screenshot({ path: `${SHOTS}/locale-${locale}.png` });
		}
		for (const width of [900, 760, 480]) {
			await page.setViewportSize({ width, height: 900 });
			for (const tab of TABS) {
				await editor.locator(`#dve-tab-${tab}`).click();
				expect(await clippedControls(page, tab), `${locale} ${tab} @${width}`).toEqual([]);
			}
		}
	});
