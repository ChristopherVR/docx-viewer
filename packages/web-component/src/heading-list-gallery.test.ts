// @vitest-environment jsdom
import { expect, it } from 'vitest';
import { createDocument } from 'docx-core';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { runListAction } from './list-commands';
import { modelToDoc } from './model-adapter';

for (const [kind, first, second] of [
	['headings', '%1', '%1.%2'],
	['article', 'Article %1.', 'Section %1.%2'],
	['roman', '%1.', '%2.'],
] as const) {
	it(`${kind}: creates a heading-linked definition and links the existing Heading styles`, () => {
		const model = createDocument();
		const view = new EditorView(document.createElement('div'), {
			state: EditorState.create({ doc: modelToDoc(model) }),
		});
		view.dispatch(view.state.tr.setSelection(TextSelection.atStart(view.state.doc)));
		runListAction(view, kind, model);
		const levels = model.numberingCatalog!.abstractNums['0']!.levels;
		expect([levels[0]!.lvlText, levels[1]!.lvlText]).toEqual([first, second]);
		const styles = model.paragraphStyles!.styles;
		expect(levels[0]!.paragraphStyleId).toBe('Heading1');
		expect(styles.Heading1!.numbering).toEqual({ numId: 1, level: 0 });
		expect(styles.Heading3!.numbering).toEqual({ numId: 1, level: 2 });
		// The default catalog has no Heading 4: that level is not linked to a missing style.
		expect(levels[3]!.paragraphStyleId).toBeUndefined();
		view.destroy();
	});
}
