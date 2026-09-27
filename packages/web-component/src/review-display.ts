import { Plugin } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';

/**
 * All Markup shows insertions/deletions as authored (default CSS from schema.ts's marks).
 * No Markup (final) hides deleted text and shows insertions unmarked, matching the resulting document.
 * Original hides inserted text and shows deletions unmarked, matching the document before the changes.
 * Simple Markup is approximated as No Markup with a change indicator; per-line change bars are not
 * implemented, so it is visually identical to No Markup today (see docs/parity-roadmap.md).
 */
export type ReviewDisplayMode = 'all' | 'simple' | 'final' | 'original';

function hiddenMarkName(mode: ReviewDisplayMode): 'deletion' | 'insertion' | null {
	if (mode === 'final' || mode === 'simple') return 'deletion';
	if (mode === 'original') return 'insertion';
	return null;
}

/** Hides insertion or deletion runs via decorations rather than mutating the document. */
export function reviewDisplayPlugin(getMode: () => ReviewDisplayMode): Plugin {
	return new Plugin({
		props: {
			decorations(state) {
				const hidden = hiddenMarkName(getMode());
				if (!hidden) return null;
				const decorations: Decoration[] = [];
				state.doc.descendants((node, pos) => {
					if (!node.isText && node.type.name !== 'hardBreak') return;
					if (node.marks.some((mark) => mark.type.name === hidden))
						decorations.push(
							Decoration.inline(pos, pos + node.nodeSize, { class: 'dve-revision-hidden' }),
						);
				});
				return DecorationSet.create(state.doc, decorations);
			},
		},
	});
}
