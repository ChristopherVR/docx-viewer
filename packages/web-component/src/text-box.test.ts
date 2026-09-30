// @vitest-environment jsdom
import { DOMSerializer } from 'prosemirror-model';
import { describe, expect, it } from 'vitest';
import { schema } from './schema';

describe('text box placeholder', () => {
	it('shows the box text read-only inside the placeholder', () => {
		const node = schema.nodes.image!.create({
			unsupported: 'Text box',
			widthPx: 200,
			heightPx: 100,
			textBoxText: JSON.stringify(['Hello box', 'Second']),
		});
		const dom = DOMSerializer.fromSchema(schema).serializeNode(node) as HTMLElement;
		expect(dom.className).toBe('dve-image-placeholder');
		expect([...dom.querySelectorAll('.dve-textbox-line')].map((l) => l.textContent)).toEqual([
			'Hello box',
			'Second',
		]);
		expect(dom.firstChild!.textContent).toBe('Text box');
	});
});
