import { expect, test, type Page } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { fileInput } from './helpers';

const FIXTURE = 'tests/support/smartart.docx';
const DIAGRAM_PARTS = ['data', 'layout', 'quickStyle', 'colors', 'drawing'].flatMap((stem) =>
	[1, 2].map((index) => `word/diagrams/${stem}${index}.xml`),
);

async function openFixture(page: Page) {
	await page.goto('/?framework=vanilla');
	await (
		await fileInput(page)
	).setInputFiles({
		name: 'smartart.docx',
		mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		buffer: await readFile(FIXTURE),
	});
	const editor = page.locator('docx-editor');
	await expect(editor.locator('.ProseMirror')).toContainText('End of document.');
	return editor;
}

const saved = (page: Page) =>
	page.locator('docx-editor').evaluate(async (element) => {
		const bytes = await (element as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes();
		return Array.from(bytes);
	});

test('SmartArt displays its cached drawing, inline and floating, read-only and accessibly', async ({
	page,
}) => {
	await page.setViewportSize({ width: 1500, height: 900 });
	const editor = await openFixture(page);
	const diagrams = editor.locator('.ProseMirror .dve-smartart');
	await expect(diagrams).toHaveCount(2);
	const [inline, floating] = [diagrams.nth(0), diagrams.nth(1)];
	await expect(inline).toHaveAttribute('data-placement', 'inline');
	await expect(floating).toHaveAttribute('data-placement', 'anchor');

	// The shared renderer draws the saved shapes, with their text, into shadow-root SVG.
	for (const [diagram, shapes] of [
		[inline, 9],
		[floating, 13],
	] as const) {
		const summary = await diagram.locator('office-ui-smartart').evaluate((element) => {
			const svg = element.shadowRoot!.querySelector('svg')!;
			const box = svg.getBoundingClientRect();
			return {
				role: svg.getAttribute('role'),
				groups: svg.querySelectorAll(':scope > g').length,
				texts: [...svg.querySelectorAll('text')].map((text) => text.textContent),
				fills: [...svg.querySelectorAll(':scope > g > :first-child')].map((shape) =>
					shape.getAttribute('fill'),
				),
				width: box.width,
				height: box.height,
			};
		});
		expect(summary.role).toBe('img');
		expect(summary.groups).toBe(shapes);
		expect(summary.texts.filter(Boolean).length).toBeGreaterThan(0);
		// Theme colours are resolved to real colours, not the unresolved grey fallback.
		expect(summary.fills.some((fill) => fill && !['none', '#9ca3af'].includes(fill))).toBe(true);
		expect(summary.width).toBeGreaterThan(50);
		expect(summary.height).toBeGreaterThan(30);
	}
	// The drawing is scaled from the diagram frame (EMU) to the run's size.
	const frame = await inline.evaluate((element) => {
		const box = element.getBoundingClientRect();
		const svg = element.querySelector('office-ui-smartart')!.shadowRoot!.querySelector('svg')!;
		return { width: box.width, view: svg.getAttribute('viewBox') };
	});
	expect(frame.view).toBe(`0 0 ${8255000 / 9525} ${5080000 / 9525}`);
	expect(frame.width).toBeGreaterThan(300);

	// Accessible name, role and keyboard focus.
	for (const diagram of [inline, floating]) {
		await expect(diagram).toHaveAttribute('role', 'group');
		const name = await diagram.getAttribute('aria-label');
		expect(name).toBeTruthy();
		await diagram.focus();
		await expect(diagram).toBeFocused();
	}
	await expect(inline).toHaveAttribute('aria-label', /Diagram 1|\S/);
	await expect(
		editor
			.locator('.ProseMirror')
			.getByRole('img', { name: (await inline.getAttribute('aria-label')) ?? '' }),
	).toHaveCount(1);

	// The notice and approximations are reachable, closed by default, and never claim fidelity.
	const info = inline.getByRole('button', { name: 'About this SmartArt rendering' });
	await inline.hover();
	await info.click();
	await expect(info).toHaveAttribute('aria-expanded', 'true');
	const notes = inline.getByRole('status');
	await expect(notes).toContainText('not recomputed and not Word-identical');
	await expect(notes).not.toContainText(/identical to Word|matches Word/i);

	// Selecting the diagram shows the selection outline.
	await inline.click({ position: { x: 5, y: 5 } });
	await expect(inline).toHaveClass(/dve-smartart-selected/);
});

test('saving an unedited SmartArt document returns the original bytes', async ({ page }) => {
	await openFixture(page);
	const original = new Uint8Array(await readFile(FIXTURE));
	expect(await saved(page)).toEqual(Array.from(original));
});

test('editing other text keeps the diagram parts and markup intact', async ({ page }) => {
	const editor = await openFixture(page);
	await editor.locator('.ProseMirror p', { hasText: 'End of document.' }).click();
	await page.keyboard.press('End');
	await page.keyboard.type(' Edited.');
	await expect(editor.locator('.ProseMirror')).toContainText('End of document. Edited.');
	await expect(editor.locator('.ProseMirror .dve-smartart')).toHaveCount(2);

	const before = await JSZip.loadAsync(await readFile(FIXTURE));
	const after = await JSZip.loadAsync(Uint8Array.from(await saved(page)));
	for (const part of DIAGRAM_PARTS) {
		expect(await after.file(part)!.async('string'), part).toBe(
			await before.file(part)!.async('string'),
		);
	}
	const document = await after.file('word/document.xml')!.async('string');
	expect(document).toContain('End of document. Edited.');
	expect((document.match(/<dgm:relIds/g) ?? []).length).toBe(2);
	expect(document).toContain('wp:anchor');
	expect(await after.file('word/_rels/document.xml.rels')!.async('string')).toBe(
		await before.file('word/_rels/document.xml.rels')!.async('string'),
	);
});
