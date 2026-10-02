// @vitest-environment jsdom
import { createDocument, halfPoints, signedTwips } from 'docx-core';
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { describe, expect, it } from 'vitest';
import { modelToDoc } from './model-adapter';
import { runFormattingCss, runStylesPlugin } from './run-styles';
import { renderBlocks } from './header-footer-view';

describe('advanced editing styles', () => {
	it('keeps superscript/subscript smaller when resolving a font size and renders inherited script', () => {
		const span = document.createElement('span');
		span.style.cssText = runFormattingCss(
			{ fontSize: 12, verticalAlign: 'superscript' },
			undefined,
			{
				verticalAlign: 'superscript',
			},
		);
		expect(parseFloat(span.style.fontSize)).toBeCloseTo(7.8, 5);
		expect(runFormattingCss({ fontSize: 12, verticalAlign: 'subscript' })).toContain(
			'vertical-align:sub',
		);
		const composed = runFormattingCss({
			fontSize: 12,
			verticalAlign: 'superscript',
			positionHalfPoints: halfPoints(6),
		});
		expect(composed).toContain('vertical-align:super');
		expect(composed).toContain('top:-3pt');
	});
	it('resolves inherited story preview formatting without flattening source runs', () => {
		const model = createDocument();
		model.characterStyles = {
			docDefaults: {
				fontSize: 12,
				kerningHalfPoints: halfPoints(24),
				positionHalfPoints: halfPoints(6),
				characterSpacingTwips: signedTwips(40),
			},
			styles: {},
			warnings: [],
		};
		model.blocks = [
			{
				type: 'paragraph',
				id: 'preview',
				runs: [
					{ text: 'Inherited' },
					{
						text: 'Plain',
						positionHalfPoints: halfPoints(0),
						kerningHalfPoints: halfPoints(0),
						characterSpacingTwips: signedTwips(0),
					},
				],
			},
		];
		const source = JSON.stringify(model.blocks);
		const preview = document.createElement('div');
		preview.append(renderBlocks(model.blocks, model));
		const spans = [...preview.querySelectorAll<HTMLElement>('p > span')];
		expect(spans[0]!.style.verticalAlign).toBe('3pt');
		expect(spans[0]!.style.fontKerning).toBe('normal');
		expect(spans[1]!.style.fontKerning).toBe('none');
		expect(spans[1]!.style.verticalAlign).toBe('');
		expect(spans[1]!.querySelector<HTMLElement>('[data-run-props]')!.style.verticalAlign).toBe(
			'0pt',
		);
		expect(JSON.stringify(model.blocks)).toBe(source);
	});
	it('renders direct positions once, including in header/footer previews', () => {
		const model = createDocument();
		model.blocks = [
			{
				type: 'paragraph',
				id: 'position',
				runs: [{ text: 'Raised', positionHalfPoints: halfPoints(6) }],
			},
		];
		const view = new EditorView(document.createElement('div'), {
			state: EditorState.create({
				doc: modelToDoc(model),
				plugins: [runStylesPlugin(() => model)],
			}),
		});
		expect(view.dom.querySelectorAll('[style*="vertical-align"]')).toHaveLength(1);
		expect(view.dom.querySelector<HTMLElement>('[data-run-props]')!.style.verticalAlign).toBe(
			'3pt',
		);
		const preview = document.createElement('div');
		preview.append(renderBlocks(model.blocks));
		expect(preview.querySelector<HTMLElement>('[data-run-props]')!.style.verticalAlign).toBe('3pt');
		view.destroy();
	});

	it('renders direct kerning thresholds without needing a style catalog', () => {
		const model = createDocument();
		delete model.characterStyles;
		delete model.theme;
		delete model.tableStyles;
		model.blocks = [
			{
				type: 'paragraph',
				id: 'kerning',
				runs: [
					{ text: 'AV', fontSize: 12, kerningHalfPoints: halfPoints(24) },
					{ text: 'small', fontSize: 10, kerningHalfPoints: halfPoints(24) },
				],
			},
		];
		const view = new EditorView(document.createElement('div'), {
			state: EditorState.create({
				doc: modelToDoc(model),
				plugins: [runStylesPlugin(() => model)],
			}),
		});
		const styles = [...view.dom.querySelectorAll<HTMLElement>('[style*="font-kerning"]')].map(
			(element) => element.style.fontKerning,
		);
		expect(styles).toEqual(['normal', 'none']);
		view.destroy();
	});

	it('renders inherited position/spacing and accepts neutral direct overrides', () => {
		expect(
			runFormattingCss({
				positionHalfPoints: halfPoints(-6),
				characterSpacingTwips: signedTwips(40),
				kerningHalfPoints: halfPoints(24),
				fontSize: 12,
			}),
		).toContain('vertical-align:-3pt');
		expect(
			runFormattingCss(
				{
					positionHalfPoints: halfPoints(0),
					characterSpacingTwips: signedTwips(0),
					kerningHalfPoints: halfPoints(0),
				},
				undefined,
				{ positionHalfPoints: halfPoints(0) },
			),
		).toContain('letter-spacing:0pt');
		expect(
			runFormattingCss({ positionHalfPoints: halfPoints(6) }, undefined, {
				positionHalfPoints: halfPoints(6),
			}),
		).not.toContain('vertical-align');
	});
});
