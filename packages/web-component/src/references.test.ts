// @vitest-environment jsdom
import { createDocument } from '@christophervr/docx-core';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import { insertCaption, renumberCaptions, seqInstruction, seqLabelOf } from './caption-commands';
import { insertCrossReference, referenceTargets } from './cross-reference-commands';
import { FormatDialogs } from './format-dialogs';
import { createRibbon } from './ribbon';
import { schema } from './schema';

beforeAll(() => {
	const rects = { length: 0, item: () => null, [Symbol.iterator]: function* () {} };
	Object.assign(Range.prototype, {
		getClientRects: () => rects,
		getBoundingClientRect: () => ({ top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0 }),
	});
});
afterEach(() => (document.body.innerHTML = ''));

const model = createDocument();

function editor(...paragraphs: Array<{ text: string; attrs?: Record<string, unknown> }>) {
	const doc = schema.node(
		'doc',
		null,
		paragraphs.map((p, index) =>
			schema.nodes.paragraph!.create({ id: `p${index}`, ...p.attrs }, schema.text(p.text)),
		),
	);
	const host = document.createElement('div');
	document.body.append(host);
	return new EditorView(host, { state: EditorState.create({ doc, schema }) });
}
const caretIn = (view: EditorView, index: number) => {
	let pos = 0;
	view.state.doc.forEach((_node, offset, i) => {
		if (i === index) pos = offset + 1;
	});
	view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, pos)));
};
const texts = (view: EditorView) => {
	const out: string[] = [];
	view.state.doc.forEach((node) => out.push(node.textContent));
	return out;
};
const caption = (view: EditorView, label = 'Figure', text = '') => {
	insertCaption(view, { label, labelText: label, text, position: 'below' });
};

describe('captions', () => {
	it('inserts a numbered caption below the current block with a SEQ field', () => {
		const view = editor({ text: 'Picture' }, { text: 'Next' });
		caretIn(view, 0);
		caption(view, 'Figure', 'A chart');
		expect(texts(view)).toEqual(['Picture', 'Figure 1: A chart', 'Next']);
		let found = '';
		view.state.doc.child(1).forEach((child) => {
			found = seqLabelOf(child) ?? found;
		});
		expect(found).toBe('Figure');
	});

	it('renumbers all captions of a label when one is inserted in the middle', () => {
		const view = editor({ text: 'a' }, { text: 'b' }, { text: 'c' });
		caretIn(view, 2);
		caption(view);
		caretIn(view, 0);
		caption(view);
		caretIn(view, 0);
		insertCaption(view, { label: 'Figure', labelText: 'Figure', text: '', position: 'above' });
		const numbers = texts(view).filter((t) => t.startsWith('Figure'));
		expect(numbers).toEqual(['Figure 1', 'Figure 2', 'Figure 3']);
	});

	it('numbers each label independently and uses the Caption style when there is one', () => {
		const view = editor({ text: 'a' });
		caption(view, 'Figure');
		caption(view, 'Table');
		expect(
			texts(view)
				.filter((t) => /^(Figure|Table)/.test(t))
				.sort(),
		).toEqual(['Figure 1', 'Table 1']);
		insertCaption(view, {
			label: 'Figure',
			labelText: 'Fig.',
			text: '',
			position: 'below',
			style: 'Caption',
		});
		let styled = view.state.doc.child(0);
		view.state.doc.forEach((node) => {
			if (node.textContent.startsWith('Fig.')) styled = node;
		});
		expect(styled.attrs.style).toBe('Caption');
		// Inserted before the earlier Figure caption, so it takes number 1 and that one becomes 2.
		expect(styled.textContent).toBe('Fig. 1');
	});

	it('refuses a read-only view, and renumber is a no-op when numbers already match', () => {
		const view = editor({ text: 'a' });
		const locked = new EditorView(document.createElement('div'), {
			state: view.state,
			editable: () => false,
		});
		expect(
			insertCaption(locked, { label: 'Figure', labelText: 'Figure', text: '', position: 'below' }),
		).toBe(false);
		const tr = renumberCaptions(view.state.tr, 'Figure');
		expect(tr.docChanged).toBe(false);
		expect(seqInstruction('Figure')).toBe(' SEQ Figure \\* ARABIC ');
	});
});

describe('cross-references', () => {
	function withTargets() {
		const view = editor(
			{ text: 'Introduction', attrs: { style: 'Heading1' } },
			{ text: 'Body', attrs: { bookmarks: ['Key', '_Toc1'] } },
			{ text: 'Closing', attrs: { style: 'Heading2' } },
			{ text: 'Cited here' },
		);
		caretIn(view, 0);
		caption(view, 'Figure', 'Result');
		caretIn(view, 3);
		return view;
	}

	it('lists headings, user bookmarks and numbered captions', () => {
		const view = withTargets();
		expect(referenceTargets(view.state.doc, model, 'heading').map((t) => t.label)).toEqual([
			'Introduction',
			'  Closing',
		]);
		expect(referenceTargets(view.state.doc, model, 'bookmark').map((t) => t.label)).toEqual([
			'Key',
		]);
		const [figure] = referenceTargets(view.state.doc, model, 'figure');
		expect(figure).toMatchObject({ numbered: 'Figure 1', text: 'Figure 1: Result' });
		expect(referenceTargets(view.state.doc, model, 'table')).toEqual([]);
	});

	it('inserts a REF field with the target text and adds a hidden bookmark to a heading', () => {
		const view = withTargets();
		caretIn(view, view.state.doc.childCount - 1);
		const [heading] = referenceTargets(view.state.doc, model, 'heading');
		expect(insertCrossReference(view, heading!, 'text', () => undefined)).toBe(true);
		const last = view.state.doc.child(view.state.doc.childCount - 1);
		expect(last.textContent).toContain('Introduction');
		let instr = '';
		last.forEach((child) => {
			const mark = child.marks.find((m) => m.type === schema.marks.field);
			if (mark) instr = String(mark.attrs.instr);
		});
		expect(instr).toMatch(/^ REF _Ref\d+ \\h $/);
		const name = /REF (\S+)/.exec(instr)![1]!;
		expect(view.state.doc.child(0).attrs.bookmarks).toContain(name);
	});

	it('reuses an existing bookmark, and can insert the label and number or the page', () => {
		const view = withTargets();
		caretIn(view, view.state.doc.childCount - 1);
		const [bookmark] = referenceTargets(view.state.doc, model, 'bookmark');
		insertCrossReference(view, bookmark!, 'page', (id) => (id === bookmark!.id ? '4' : undefined));
		let field = '';
		let instr = '';
		view.state.doc.lastChild!.forEach((child) => {
			const mark = child.marks.find((m) => m.type === schema.marks.field);
			if (mark) {
				field = child.text ?? '';
				instr = String(mark.attrs.instr);
			}
		});
		expect(instr).toBe(' PAGEREF Key \\h ');
		expect(field).toBe('4');
		const [figure] = referenceTargets(view.state.doc, model, 'figure');
		insertCrossReference(view, figure!, 'label', () => undefined);
		expect(view.state.doc.lastChild!.textContent).toContain('Figure 1');
		expect(view.state.doc.lastChild!.textContent).not.toContain('Result');
	});

	it('refuses a stale target and a read-only view', () => {
		const view = withTargets();
		expect(
			insertCrossReference(
				view,
				{ label: 'x', pos: 9999, id: 'x', text: 'x' },
				'text',
				() => undefined,
			),
		).toBe(false);
		const locked = new EditorView(document.createElement('div'), {
			state: view.state,
			editable: () => false,
		});
		expect(
			insertCrossReference(
				locked,
				referenceTargets(view.state.doc, model, 'heading')[0]!,
				'text',
				() => undefined,
			),
		).toBe(false);
	});
});

describe('caption and cross-reference dialogs', () => {
	const setup = () => {
		const view = editor({ text: 'Intro', attrs: { style: 'Heading1' } }, { text: 'Body' });
		const dialogs = new FormatDialogs({ view: () => view, model: () => model });
		document.body.append(...dialogs.elements);
		return { view, dialogs };
	};
	const field = (root: ParentNode, label: string) =>
		root.querySelector<HTMLInputElement & HTMLSelectElement>(`[aria-label="${label}"]`)!;
	const button = (root: ParentNode, text: string) =>
		[...root.querySelectorAll<HTMLButtonElement>('button')].find((b) => b.textContent === text)!;

	it('the caption dialog inserts a caption and defaults tables to above', () => {
		const { view, dialogs } = setup();
		dialogs.open('caption');
		const root = dialogs.elements[3]!;
		field(root, 'Caption').value = 'Sales';
		button(root, 'OK').click();
		expect(texts(view).some((t) => t === 'Figure 1: Sales')).toBe(true);
		dialogs.open('caption');
		const again = dialogs.elements[3]!;
		field(again, 'Label').value = 'Table';
		field(again, 'Label').dispatchEvent(new Event('change'));
		expect(field(again, 'Position').value).toBe('above');
	});

	it('the cross-reference dialog lists targets by type and inserts a reference', () => {
		const { view, dialogs } = setup();
		dialogs.open('crossReference');
		const root = dialogs.elements[4]!;
		expect([...field(root, 'For which item').options].map((o) => o.textContent)).toEqual(['Intro']);
		expect(button(root, 'Insert').disabled).toBe(false);
		const kind = field(root, 'Reference type');
		kind.value = 'figure';
		kind.dispatchEvent(new Event('change'));
		expect(button(root, 'Insert').disabled).toBe(true);
		expect([...field(root, 'Insert reference to').options].map((o) => o.value)).toEqual([
			'text',
			'label',
			'page',
		]);
		kind.value = 'heading';
		kind.dispatchEvent(new Event('change'));
		caretIn(view, 1);
		button(root, 'Insert').click();
		expect(dialogs.elements[4]!.hidden).toBe(false);
		expect(view.state.doc.child(1).textContent).toContain('Intro');
		button(root, 'Close').click();
		expect(dialogs.elements[4]!.hidden).toBe(true);
	});

	it('has ribbon buttons for both on the References tab', () => {
		const ribbon = createRibbon();
		const seen: unknown[] = [];
		ribbon.addEventListener('ribbon-action', (event) => seen.push((event as CustomEvent).detail));
		ribbon.querySelector<HTMLButtonElement>('button[aria-label="Insert caption"]')!.click();
		ribbon.querySelector<HTMLButtonElement>('button[aria-label="Cross-reference"]')!.click();
		expect(seen).toEqual([
			{ type: 'formatDialog', kind: 'caption' },
			{ type: 'formatDialog', kind: 'crossReference' },
		]);
	});
});
