import { expect, test } from '@playwright/test';

test('all bindings keep the shared ribbon, canvas, and status within their host', async ({
	page,
}) => {
	const frameworks = ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'];
	for (const framework of frameworks) {
		await page.setViewportSize({ width: 1280, height: 900 });
		await page.goto(`/?framework=${framework}`);
		const editor = page.locator('docx-editor');
		await expect(editor).toBeVisible();
		await expect(editor.locator('.dve-status')).toBeVisible();

		const desktop = await page.evaluate(() => {
			const host = document.querySelector('docx-editor')!;
			const container = document.querySelector('#editor')!;
			const root = host.shadowRoot!;
			const panel = root.querySelector<HTMLElement>('.ribbon-panel[data-panel="Home"]')!;
			const groups = [...panel.querySelectorAll<HTMLElement>('.ribbon-group')];
			return {
				hostWidth: host.getBoundingClientRect().width,
				containerWidth: container.getBoundingClientRect().width,
				panelScrolls: panel.scrollWidth > panel.clientWidth,
				groupsOverlap: groups.some((group, index) => {
					const current = group.getBoundingClientRect();
					const next = groups[index + 1]?.getBoundingClientRect();
					return next ? current.right > next.left + 1 : false;
				}),
				statusAtBottom:
					root.querySelector('.dve-status')?.previousElementSibling ===
					root.querySelector('.dve-canvas'),
			};
		});
		expect(desktop.hostWidth).toBeLessThanOrEqual(desktop.containerWidth + 1);
		expect(desktop.panelScrolls).toBe(false);
		expect(desktop.groupsOverlap).toBe(false);
		expect(desktop.statusAtBottom).toBe(true);

		await page.setViewportSize({ width: 390, height: 844 });
		const embedded = await editor.evaluate((node) => {
			const host = node as HTMLElement;
			host.style.width = '48%';
			const root = host.shadowRoot!;
			const tabs = root.querySelector<HTMLElement>('.ribbon-tabs')!;
			const panel = root.querySelector<HTMLElement>('.ribbon-panel[data-panel="Home"]')!;
			const canvas = root.querySelector<HTMLElement>('.dve-canvas')!;
			const status = root.querySelector<HTMLElement>('.dve-status')!;
			return {
				hostWidth: host.getBoundingClientRect().width,
				pageOverflows: document.documentElement.scrollWidth > document.documentElement.clientWidth,
				tabsScroll: tabs.scrollWidth > tabs.clientWidth,
				panelScroll: panel.scrollWidth > panel.clientWidth,
				canvasOwnsVerticalScroll: getComputedStyle(canvas).overflowY === 'auto',
				statusInsideHost:
					status.getBoundingClientRect().bottom <= host.getBoundingClientRect().bottom + 1,
			};
		});
		expect(embedded.hostWidth).toBeLessThan(200);
		expect(embedded.pageOverflows).toBe(false);
		expect(embedded.tabsScroll).toBe(true);
		expect(embedded.panelScroll).toBe(true);
		expect(embedded.canvasOwnsVerticalScroll).toBe(true);
		expect(embedded.statusInsideHost).toBe(true);
	}
});
