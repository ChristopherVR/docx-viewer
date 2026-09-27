import { test, expect } from '@playwright/test';

for (const guest of ['vue', 'solid']) {
	test(`local coauthor demo relays accessible presence to ${guest}`, async ({ page }) => {
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		await page.goto(`/collaboration.html?framework=vanilla&guest=${guest}`);
		const peerA = page.locator('#peer-a docx-editor');
		const peerB = page.locator('#peer-b docx-editor');
		const aliceCursor = peerB.getByRole('img', { name: "Ada's cursor" });
		const graceCursor = peerA.getByRole('img', { name: "Grace's cursor" });
		await expect(aliceCursor).toBeVisible();
		await expect(graceCursor).toBeVisible();
		await page.getByLabel('Editor B cursor color').selectOption('#a21caf');
		await page.getByLabel('Name', { exact: true }).nth(1).fill('Grace Hopper');
		const updatedCursor = peerA.getByRole('img', { name: "Grace Hopper's cursor" });
		await expect(updatedCursor).toBeVisible();
		await expect(updatedCursor).toHaveCSS('border-left-color', 'rgb(162, 28, 175)');
		await expect(peerA.getByRole('img', { name: "Grace's cursor" })).toHaveCount(0);

		await peerB.evaluate((element) =>
			(element as unknown as { leavePresence: () => unknown }).leavePresence(),
		);
		await expect(peerA.getByRole('img', { name: "Grace Hopper's cursor" })).toHaveCount(0);
		await expect(page.getByRole('status')).toContainText('Synced');
		await expect(errors).toEqual([]);
	});
}
