// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { createDocument, type DocumentModel } from '@christophervr/docx-core';
import { buildFooterElement, buildHeaderElement } from './header-footer-view';
import { buildNotesElement } from './notes-view';

function withSection(
	model: DocumentModel,
	overrides: Partial<DocumentModel['sections'] extends (infer T)[] | undefined ? T : never>,
) {
	model.sections = [
		{
			endsAtBlockId: model.blocks.at(-1)!.id,
			type: 'nextPage',
			pageWidthTwips: 12240,
			pageHeightTwips: 15840,
			orientation: 'portrait',
			marginTopTwips: 1440,
			marginRightTwips: 1440,
			marginBottomTwips: 1440,
			marginLeftTwips: 1440,
			columns: { count: 1, equalWidth: true },
			...overrides,
		} as never,
	];
	return model;
}

describe('header/footer read-only view', () => {
	it('renders the default header paragraph text and marks it read-only', () => {
		const model = withSection(createDocument(), {
			headers: {
				default: {
					blocks: [{ type: 'paragraph', id: 'h0', runs: [{ text: 'Header text' }] }],
				},
			},
		});
		const el = buildHeaderElement(model, 'en');
		expect(el).not.toBeNull();
		expect(el!.getAttribute('contenteditable')).toBe('false');
		expect(el!.textContent).toContain('Header text');
	});

	it('renders first/even/default footer slots in order', () => {
		const model = withSection(createDocument(), {
			footers: {
				first: { blocks: [{ type: 'paragraph', id: 'f0', runs: [{ text: 'First footer' }] }] },
				default: { blocks: [{ type: 'paragraph', id: 'f1', runs: [{ text: 'Default footer' }] }] },
			},
		});
		const el = buildFooterElement(model, 'en');
		const text = el!.textContent ?? '';
		expect(text.indexOf('First footer')).toBeLessThan(text.indexOf('Default footer'));
	});

	it('returns null when the section has no headers or footers', () => {
		const model = createDocument();
		expect(buildHeaderElement(model, 'en')).toBeNull();
		expect(buildFooterElement(model, 'en')).toBeNull();
	});
});

describe('footnote/endnote read-only view', () => {
	it('renders footnotes before endnotes, each numbered by reference order', () => {
		const model = createDocument();
		model.blocks = [
			{
				type: 'paragraph',
				id: 'p1',
				runs: [{ text: 'Body' }, { text: '', noteReference: { kind: 'footnote', id: '2' } }],
			},
		];
		model.footnotes = [
			{ id: '2', blocks: [{ type: 'paragraph', id: 'fn2-p0', runs: [{ text: 'Note body' }] }] },
		];
		model.endnotes = [
			{ id: '5', blocks: [{ type: 'paragraph', id: 'en5-p0', runs: [{ text: 'End body' }] }] },
		];
		const el = buildNotesElement(model, 'en');
		expect(el).not.toBeNull();
		const text = el!.textContent ?? '';
		expect(text).toContain('Note body');
		expect(text).toContain('End body');
		expect(text.indexOf('Note body')).toBeLessThan(text.indexOf('End body'));
	});

	it('returns null when there are no notes', () => {
		expect(buildNotesElement(createDocument(), 'en')).toBeNull();
	});
});
