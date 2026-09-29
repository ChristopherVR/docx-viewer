// @vitest-environment jsdom
import { createDocument } from '@christophervr/docx-core';
import { describe, expect, it, vi } from 'vitest';
import { syncStyleGallery } from './style-gallery';

const model = createDocument();
const styles = [
	{ id: '', name: 'Normal' },
	{ id: 'Heading1', name: 'heading 1' },
];

describe('style gallery', () => {
	it('renders a tile per style with a preview and marks the selected one', () => {
		const gallery = document.createElement('div');
		syncStyleGallery(gallery, styles, 'Heading1', model, () => {}, false);
		const tiles = [...gallery.querySelectorAll<HTMLButtonElement>('.style-tile')];
		expect(tiles.map((tile) => tile.getAttribute('aria-label'))).toEqual(['Normal', 'heading 1']);
		expect(tiles.map((tile) => tile.getAttribute('aria-pressed'))).toEqual(['false', 'true']);
		const heading = tiles[1]!.querySelector<HTMLElement>('.style-sample')!;
		expect(parseFloat(heading.style.fontSize)).toBeGreaterThan(
			parseFloat(tiles[0]!.querySelector<HTMLElement>('.style-sample')!.style.fontSize || '11'),
		);
	});

	it('keeps tiles when only the selection changes and reports the chosen id', () => {
		const gallery = document.createElement('div');
		const choose = vi.fn();
		syncStyleGallery(gallery, styles, '', model, choose, false);
		const first = gallery.querySelector('.style-tile');
		syncStyleGallery(gallery, styles, 'Heading1', model, choose, false);
		expect(gallery.querySelector('.style-tile')).toBe(first);
		gallery.querySelector<HTMLButtonElement>('[data-style-id="Heading1"]')!.click();
		expect(choose).toHaveBeenCalledWith('Heading1');
	});

	it('disables tiles for a read-only document', () => {
		const gallery = document.createElement('div');
		syncStyleGallery(gallery, styles, '', model, () => {}, true);
		expect(
			[...gallery.querySelectorAll<HTMLButtonElement>('.style-tile')].every((t) => t.disabled),
		).toBe(true);
	});
});
