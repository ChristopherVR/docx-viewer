// @vitest-environment jsdom
import { createDocument } from '@christophervr/docx-core';
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { describe, expect, it } from 'vitest';
import { createRibbon } from './ribbon';
import { createRuler, inchLabels, markerPositions, updateRuler, type RulerGeometry } from './ruler';
import { rulerGeometry, syncRuler } from './ruler-sync';
import { schema } from './schema';

const geometry = (over: Partial<RulerGeometry> = {}): RulerGeometry => ({
	pageWidth: 816,
	marginLeft: 96,
	marginRight: 96,
	indentLeft: 0,
	indentRight: 0,
	firstLine: 0,
	...over,
});

describe('ruler geometry', () => {
	it('places the markers from the margins and the paragraph indents', () => {
		expect(markerPositions(geometry())).toEqual({ firstLine: 96, left: 96, right: 720 });
		expect(markerPositions(geometry({ indentLeft: 48, indentRight: 24, firstLine: 24 }))).toEqual({
			firstLine: 168,
			left: 144,
			right: 696,
		});
		expect(markerPositions(geometry({ indentLeft: 48, firstLine: -48 })).firstLine).toBe(96);
	});

	it('labels whole inches from the left margin up to the right margin', () => {
		expect(inchLabels(geometry())).toEqual([
			{ at: 192, text: '1' },
			{ at: 288, text: '2' },
			{ at: 384, text: '3' },
			{ at: 480, text: '4' },
			{ at: 576, text: '5' },
			{ at: 672, text: '6' },
		]);
		expect(inchLabels(geometry({ marginLeft: 0, marginRight: 0, pageWidth: 90 }))).toEqual([]);
	});
});

describe('ruler element', () => {
	it('sizes itself, shades the margins, moves markers and follows the zoom', () => {
		const ruler = createRuler();
		updateRuler(ruler, geometry({ indentLeft: 48 }), 1.5);
		expect(ruler.style.width).toBe('816px');
		expect(ruler.style.getPropertyValue('--dve-zoom')).toBe('1.5');
		expect(ruler.querySelector<HTMLElement>('.dve-ruler-margin-left')!.style.width).toBe('96px');
		expect(ruler.querySelector<HTMLElement>('.dve-ruler-marker-left')!.style.left).toBe('144px');
		expect(ruler.querySelectorAll('.dve-ruler-labels span')).toHaveLength(6);
		updateRuler(ruler, geometry({ indentLeft: 48 }), 1.5);
		expect(ruler.querySelectorAll('.dve-ruler-labels span')).toHaveLength(6);
	});
});

describe('ruler sync', () => {
	function editor(attrs: Record<string, unknown>) {
		const doc = schema.node('doc', { pageWidth: 816, marginLeft: 96, marginRight: 96 }, [
			schema.nodes.paragraph!.create({ id: 'p', ...attrs }, schema.text('text')),
		]);
		return new EditorView(document.createElement('div'), {
			state: EditorState.create({ doc, schema }),
		});
	}

	it("reads the caret paragraph's indents, hanging indent included", () => {
		const view = editor({ indentLeftTwips: 720, hangingTwips: 360, indentRightTwips: 150 });
		expect(rulerGeometry(view, createDocument())).toEqual({
			pageWidth: 816,
			marginLeft: 96,
			marginRight: 96,
			indentLeft: 48,
			indentRight: 10,
			firstLine: -24,
		});
	});

	it('updates a ruler in the frame and does nothing when there is none', () => {
		const frame = document.createElement('div');
		const ribbon = createRibbon();
		frame.append(ribbon);
		const view = editor({ indentLeftTwips: 720 });
		syncRuler(ribbon, view, createDocument());
		const ruler = createRuler();
		frame.append(ruler);
		syncRuler(ribbon, view, createDocument());
		expect(ruler.querySelector<HTMLElement>('.dve-ruler-marker-left')!.style.left).toBe('144px');
	});

	it('has a View button that emits the toggle', () => {
		const ribbon = createRibbon();
		const seen: unknown[] = [];
		ribbon.addEventListener('ribbon-action', (event) => seen.push((event as CustomEvent).detail));
		ribbon.querySelector<HTMLButtonElement>('button[aria-label="Ruler"]')!.click();
		expect(seen).toEqual([{ type: 'ruler' }]);
	});
});
