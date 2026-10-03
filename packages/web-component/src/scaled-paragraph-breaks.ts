import type { Node as ProseMirrorNode } from 'prosemirror-model';
import { Decoration } from 'prosemirror-view';
import { tokenizeRun } from 'ooxml-core/docx/layout';

/** Atomic scaled spans must not create a line break at a formatting boundary.
 * Keep the paragraph's source whitespace and expose only the core's text breaks.
 * Empty wbr widgets have no text to leak into copying, editing or serialization.
 */
export function scaledParagraphBreaks(paragraph: ProseMirrorNode, start: number): Decoration[] {
	let text = '';
	paragraph.forEach((child) => {
		text += child.isText ? child.text! : '\t';
	});
	const decorations = [
		Decoration.node(start - 1, start + paragraph.content.size + 1, {
			style: 'white-space:pre',
		}),
	];
	for (const token of tokenizeRun(text, 0)) {
		if (token.kind !== 'word' || token.sourceStart === 0) continue;
		decorations.push(
			Decoration.widget(
				start + token.sourceStart,
				(view) => {
					const breakPoint = view.dom.ownerDocument.createElement('wbr');
					breakPoint.setAttribute('aria-hidden', 'true');
					return breakPoint;
				},
				{ side: -1, key: `scaled-break:${start + token.sourceStart}`, ignoreSelection: true },
			),
		);
	}
	return decorations;
}
