import { describe, expect, it } from 'vitest';
import * as common from './common';
import * as react from './react';
import * as vue from './vue';
import * as solid from './solid';

describe('surface shared by every published framework package', () => {
	it('re-exports the editor configuration and document loading helpers', () => {
		for (const name of [
			'loadDocument',
			'detectDocumentFormat',
			'registerDocxEditor',
			'normalizeEditorLocale',
			'normalizeRibbonActions',
		] as const)
			expect(common[name], name).toBeTypeOf('function');
		expect(common.RIBBON_ACTION_IDS.length).toBeGreaterThan(0);
	});

	it('exports each framework component under the documented name', () => {
		expect(react.WordEditor).toBeDefined();
		expect(vue.WordEditor).toBeDefined();
		expect(solid.WordEditor).toBeTypeOf('function');
	});
});
