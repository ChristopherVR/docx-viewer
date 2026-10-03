// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { EditorState, TextSelection, AllSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { schema } from './schema';
import { copyOrCut, pasteClipboard } from './context-menu-actions';

const views: EditorView[] = [];
const makeView = (doc = schema.nodes.doc.create(null, schema.nodes.paragraph.create())) => {
	vi.stubGlobal('ClipboardEvent', Event);
	const host = document.body.appendChild(document.createElement('div'));
	const view = new EditorView(host, {
		state: EditorState.create({ doc }),
		handleScrollToSelection: () => true,
	});
	views.push(view);
	return view;
};
const clipboard = (value: object) =>
	Object.defineProperty(navigator, 'clipboard', { configurable: true, value });
const htmlItem = (html: string) => ({
	types: ['text/html', 'text/plain'],
	getType: async (type: string) => ({
		text: async () => (type === 'text/html' ? html : 'flattened'),
	}),
});

afterEach(() => {
	views.splice(0).forEach((view) => view.destroy());
	document.body.replaceChildren();
	vi.restoreAllMocks();
	vi.unstubAllGlobals();
});

describe('formatted ribbon and context-menu clipboard', () => {
	it('prefers HTML and retains formatting, links, paragraphs and table structure', async () => {
		clipboard({
			read: async () => [
				htmlItem(
					'<p><strong>Bold</strong> <a data-docx-link="1" href="https://example.com">Link</a></p><table><tbody><tr><td><p>Cell A</p></td><td><p><em>Cell B</em></p></td></tr></tbody></table>',
				),
			],
			readText: vi.fn(),
		});
		const view = makeView();
		expect(await pasteClipboard(view)).toBe(true);
		expect(view.state.doc.child(0).firstChild!.marks[0]!.type.name).toBe('bold');
		expect(view.state.doc.child(0).lastChild!.marks[0]!.attrs.href).toBe('https://example.com');
		expect(view.state.doc.child(1).type.name).toBe('table');
		expect(view.state.doc.child(1).firstChild!.childCount).toBe(2);
		expect(
			view.state.doc.child(1).firstChild!.lastChild!.firstChild!.firstChild!.marks[0]!.type.name,
		).toBe('italic');
	});

	it('copies a ProseMirror HTML slice instead of flattening fallback copy', async () => {
		Object.defineProperty(document, 'execCommand', { configurable: true, value: () => false });
		vi.stubGlobal(
			'ClipboardItem',
			class {
				constructor(readonly data: Record<string, Blob>) {}
			},
		);
		const write = vi.fn(async (_items: ClipboardItem[]) => {});
		clipboard({ write });
		const view = makeView(
			schema.nodes.doc.create(null, [
				schema.nodes.paragraph.create(null, schema.text('bold', [schema.marks.bold.create()])),
				schema.nodes.paragraph.create(null, schema.text('second')),
			]),
		);
		view.dispatch(view.state.tr.setSelection(new AllSelection(view.state.doc)));
		expect(await copyOrCut(view, 'copy')).toBe(true);
		const item = write.mock.calls[0]![0] as unknown as { data: Record<string, Blob> }[];
		expect(item[0]!.data['text/html']!.type).toBe('text/html');
		expect(item[0]!.data['text/plain']!.type).toBe('text/plain');
		expect(view.state.doc.textContent).toBe('boldsecond');
	});

	it('never cuts selected content when clipboard write is denied', async () => {
		Object.defineProperty(document, 'execCommand', { configurable: true, value: () => false });
		clipboard({
			writeText: async () => {
				throw new Error('denied');
			},
		});
		const view = makeView(
			schema.nodes.doc.create(null, schema.nodes.paragraph.create(null, schema.text('keep'))),
		);
		view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 5)));
		expect(await copyOrCut(view, 'cut')).toBe(false);
		expect(view.state.doc.textContent).toBe('keep');
	});

	it('does not delete a changed selection after an asynchronous cut write', async () => {
		Object.defineProperty(document, 'execCommand', { configurable: true, value: () => false });
		let resolve!: () => void;
		clipboard({
			writeText: () =>
				new Promise<void>((done) => {
					resolve = done;
				}),
		});
		const view = makeView(
			schema.nodes.doc.create(
				null,
				schema.nodes.paragraph.create(null, schema.text('first second')),
			),
		);
		view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 6)));
		const result = copyOrCut(view, 'cut');
		view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 7, 13)));
		resolve();
		expect(await result).toBe(true);
		expect(view.state.doc.textContent).toBe('first second');
	});

	it('cuts only after a successful write and labels the transaction for tracked moves', async () => {
		Object.defineProperty(document, 'execCommand', { configurable: true, value: () => false });
		clipboard({ writeText: async () => {} });
		const view = makeView(
			schema.nodes.doc.create(
				null,
				schema.nodes.paragraph.create(null, schema.text('first second')),
			),
		);
		view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 7)));
		const dispatch = vi.spyOn(view, 'dispatch');
		expect(await copyOrCut(view, 'cut')).toBe(true);
		expect(view.state.doc.textContent).toBe('second');
		expect(dispatch.mock.calls[0]![0].getMeta('uiEvent')).toBe('cut');
	});

	it('supports plain text browsers and rejects paste in viewing mode', async () => {
		const readText = vi.fn(async () => 'plain\nsecond');
		clipboard({ readText });
		const view = makeView();
		expect(await pasteClipboard(view)).toBe(true);
		expect(view.state.doc.childCount).toBe(2);
		view.setProps({ editable: () => false });
		expect(await pasteClipboard(view)).toBe(false);
		expect(readText).toHaveBeenCalledTimes(1);
	});
});
