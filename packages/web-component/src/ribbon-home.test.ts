// @vitest-environment jsdom
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { describe, expect, it, vi } from 'vitest';
import { steppedFontSize, stepFontSize } from './font-step';
import { applyRibbonVisibility } from './ribbon-visibility';
import { createRibbon } from './ribbon';
import { schema } from './schema';

const HOME_GROUPS = ['Clipboard', 'Font', 'Paragraph', 'Editing'];

describe('Home tab layout', () => {
	it('orders groups as Word does and drops the old Alignment group', () => {
		const home = createRibbon().querySelector('[data-panel="Home"]')!;
		const labels = [...home.querySelectorAll<HTMLElement>('.ribbon-group')].map(
			(group) => group.dataset.label,
		);
		expect(labels).toEqual(HOME_GROUPS);
	});

	it('moves paragraph spacing to the Layout tab', () => {
		const ribbon = createRibbon();
		expect(ribbon.querySelector('[data-panel="Home"] [aria-label="Spacing before"]')).toBeNull();
		expect(
			ribbon.querySelector('[data-panel="Layout"] [aria-label="Spacing before"]'),
		).not.toBeNull();
	});

	it('draws icons, not text glyphs, on formatting buttons', () => {
		const bold = createRibbon().querySelector('[aria-label="Bold"]')!;
		expect(bold.querySelector('svg')).not.toBeNull();
		expect(bold.textContent).toBe('');
		expect(bold.getAttribute('title')).toBe('Bold (Ctrl+B)');
	});
});

describe('split colour buttons', () => {
	it('open a swatch popover, apply the chosen colour and remember it on the button', () => {
		const host = document.createElement('div');
		document.body.append(host);
		const ribbon = createRibbon();
		host.append(ribbon);
		const actions: unknown[] = [];
		ribbon.addEventListener('ribbon-action', (event) =>
			actions.push((event as CustomEvent).detail),
		);
		const main = ribbon.querySelector<HTMLButtonElement>('[aria-label="Font color"]')!;
		main.click();
		expect(actions.at(-1)).toEqual({ type: 'font', key: 'color', value: '#c00000' });
		(main.nextElementSibling as HTMLButtonElement).click();
		const swatch = document.querySelector<HTMLButtonElement>(
			'.ribbon-popover [aria-label="Blue"]',
		)!;
		swatch.click();
		expect(actions.at(-1)).toEqual({ type: 'font', key: 'color', value: '#0070c0' });
		expect(main.dataset.value).toBe('#0070c0');
		expect(document.querySelector('.ribbon-popover')).toBeNull();
		main.click();
		expect(actions.at(-1)).toEqual({ type: 'font', key: 'color', value: '#0070c0' });
		host.remove();
	});

	it('closes the popover on Escape', () => {
		const host = document.createElement('div');
		document.body.append(host);
		const ribbon = createRibbon();
		host.append(ribbon);
		(ribbon.querySelector('[data-split-caret]') as HTMLButtonElement).click();
		expect(document.querySelector('.ribbon-popover')).not.toBeNull();
		document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
		expect(document.querySelector('.ribbon-popover')).toBeNull();
		host.remove();
	});

	it('hide their caret together with the colour action', () => {
		const ribbon = createRibbon();
		applyRibbonVisibility(ribbon, true, ['font-color']);
		const main = ribbon.querySelector('[aria-label="Font color"]')!;
		expect(main.hasAttribute('data-dve-hidden')).toBe(true);
		expect(main.nextElementSibling?.hasAttribute('data-dve-hidden')).toBe(true);
	});
});

describe('grow and shrink font', () => {
	it("steps along Word's size ladder", () => {
		expect(steppedFontSize(11, 'grow')).toBe(12);
		expect(steppedFontSize(11, 'shrink')).toBe(10.5);
		expect(steppedFontSize(72, 'grow')).toBe(74);
		expect(steppedFontSize(8, 'shrink')).toBe(6);
		expect(steppedFontSize(1, 'shrink')).toBe(1);
	});

	it('changes the size of the selected text only', () => {
		const doc = schema.node('doc', null, [
			schema.nodes.paragraph.create({ id: 'p' }, schema.text('grow me')),
		]);
		const view = new EditorView(document.createElement('div'), {
			state: EditorState.create({ doc, schema }),
		});
		view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 5)));
		stepFontSize(view, 'grow');
		const sizes: Array<[string, unknown]> = [];
		view.state.doc.forEach((p) =>
			p.forEach((text) =>
				sizes.push([
					text.text ?? '',
					text.marks.find((m) => m.type === schema.marks.font)?.attrs.size,
				]),
			),
		);
		expect(sizes).toEqual([
			['grow', 12],
			[' me', undefined],
		]);
		view.destroy();
	});
});

describe('clipboard buttons', () => {
	it('emit clipboard actions for the router', () => {
		const ribbon = createRibbon();
		const seen: unknown[] = [];
		ribbon.addEventListener('ribbon-action', (event) => seen.push((event as CustomEvent).detail));
		for (const name of ['Paste', 'Cut', 'Copy'])
			ribbon.querySelector<HTMLButtonElement>(`[aria-label="${name}"]`)!.click();
		expect(seen).toEqual([
			{ type: 'clipboard', key: 'paste' },
			{ type: 'clipboard', key: 'cut' },
			{ type: 'clipboard', key: 'copy' },
		]);
		vi.restoreAllMocks();
	});
});

describe('other tabs', () => {
	it('draw every button with an icon and keep text captions to real words', () => {
		const ribbon = createRibbon();
		for (const panel of ['Insert', 'Layout', 'References', 'Review', 'View', 'Table']) {
			const buttons = ribbon.querySelectorAll(`[data-panel="${panel}"] button[aria-label]`);
			expect(buttons.length).toBeGreaterThan(0);
			for (const button of buttons) {
				expect(
					button.querySelector('svg'),
					`${panel}: ${button.getAttribute('aria-label')}`,
				).not.toBeNull();
				expect(button.textContent).not.toMatch(/[▦↑↓←→−⇤⇥]/u);
			}
		}
	});

	it('lists the Language group after the review controls', () => {
		const labels = [
			...createRibbon().querySelectorAll('[data-panel="Review"] > .ribbon-group'),
		].map((group) => group.getAttribute('data-label'));
		expect(labels).toEqual(['Proofing', 'Tracking', 'Changes', 'Comments', 'Language']);
	});
});
