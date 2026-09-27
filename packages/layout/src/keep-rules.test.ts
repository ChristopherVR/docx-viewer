import { describe, expect, it } from 'vitest';
import { adjustForWidowOrphan, suppressesSpacing, widowControlEnabled } from './keep-rules.js';

describe('widowControlEnabled', () => {
	it('defaults on when unspecified', () => {
		expect(widowControlEnabled(undefined)).toBe(true);
		expect(widowControlEnabled(true)).toBe(true);
		expect(widowControlEnabled(false)).toBe(false);
	});
});

describe('adjustForWidowOrphan', () => {
	it('leaves placement untouched when the whole paragraph fits or control is off', () => {
		expect(adjustForWidowOrphan(5, 5, true)).toBe(5);
		expect(adjustForWidowOrphan(5, 1, false)).toBe(1);
	});

	it('pushes a single stranded first line (orphan) entirely to the next page', () => {
		expect(adjustForWidowOrphan(5, 1, true)).toBe(0);
	});

	it('pulls one extra line back so a single last line never starts alone (widow)', () => {
		expect(adjustForWidowOrphan(5, 4, true)).toBe(3);
	});

	it('leaves a 2+/2+ split untouched', () => {
		expect(adjustForWidowOrphan(5, 2, true)).toBe(2);
		expect(adjustForWidowOrphan(5, 3, true)).toBe(3);
	});
});

describe('suppressesSpacing', () => {
	it('requires a shared style id and at least one contextualSpacing flag', () => {
		expect(suppressesSpacing(undefined, { styleId: 'Body', contextualSpacing: true })).toBe(false);
		expect(
			suppressesSpacing({ styleId: 'Body' }, { styleId: 'Body', contextualSpacing: true }),
		).toBe(true);
		expect(
			suppressesSpacing({ styleId: 'Body', contextualSpacing: true }, { styleId: 'Body' }),
		).toBe(true);
		expect(
			suppressesSpacing({ styleId: 'Body' }, { styleId: 'Heading', contextualSpacing: true }),
		).toBe(false);
		expect(suppressesSpacing({ styleId: 'Body' }, { styleId: 'Body' })).toBe(false);
	});
});
