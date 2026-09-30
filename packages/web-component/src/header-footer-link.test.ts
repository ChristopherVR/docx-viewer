import { createDocument } from '@christophervr/docx-core';
import { expect, it } from 'vitest';
import { withBlankHeaderFooter } from './header-footer-commands';
import { effectiveHeaderFooter, withHeaderFooterLink } from './header-footer-link';

it('links each story independently and makes an independent content copy on unlink', () => {
	const model = withBlankHeaderFooter(createDocument(), 'headers', () => 'h');
	const section = model.sections![0]!;
	model.sections = [section, { ...section, endsAtBlockId: 'p2', headers: {} }];
	expect(effectiveHeaderFooter(model, 1, 'headers', 'default')).toBe(section.headers!.default);
	const unlinked = withHeaderFooterLink(model, 1, 'headers', 'default', false);
	expect(unlinked.sections![1]!.headers!.default).toMatchObject({
		partName: 'word/header2.xml',
		sourcePartName: 'word/header1.xml',
	});
	expect(unlinked.sections![1]!.headers!.default!.blocks).not.toBe(
		section.headers!.default!.blocks,
	);
	const linked = withHeaderFooterLink(unlinked, 1, 'headers', 'default', true);
	expect(linked.sections![1]!.headers!.default).toBeUndefined();
	expect(withHeaderFooterLink(model, 0, 'headers', 'default', false)).toBe(model);
});
