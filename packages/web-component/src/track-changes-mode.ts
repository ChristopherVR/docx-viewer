import { Plugin, PluginKey } from 'prosemirror-state';
import { ReplaceStep, Transform } from 'prosemirror-transform';
import { Fragment, Slice, type Mark, type Node as ProseMirrorNode } from 'prosemirror-model';
import { schema } from './schema';

export const trackChangesPluginKey = new PluginKey<boolean>('dve-track-changes');
/** Transactions carrying this meta (remote collaboration steps) are never re-tracked. */
export const REMOTE_TRANSACTION_META = 'dve-remote';

let revisionSerial = 0;
function nextRevisionId(): string {
	return `dve-rev-${Date.now().toString(36)}-${++revisionSerial}`;
}

function withMark(fragment: Fragment, mark: Mark): Fragment {
	const children: ProseMirrorNode[] = [];
	fragment.forEach((node) => {
		if (node.isText || node.type.name === 'hardBreak')
			children.push(node.mark(mark.addToSet(node.marks)));
		else children.push(node.copy(withMark(node.content, mark)));
	});
	return Fragment.fromArray(children);
}

interface Segment {
	from: number;
	to: number;
	ownInsertion: boolean;
}
function segmentsOf(doc: ProseMirrorNode, from: number, to: number, author: string): Segment[] {
	const segments: Segment[] = [];
	doc.nodesBetween(from, to, (node, pos) => {
		if (!node.isText && node.type.name !== 'hardBreak') return true;
		const start = Math.max(pos, from);
		const end = Math.min(pos + node.nodeSize, to);
		if (end <= start) return true;
		const mark = node.marks.find((item) => item.type.name === 'insertion');
		const ownInsertion = Boolean(mark && mark.attrs.author === author);
		const last = segments.at(-1);
		if (last && last.ownInsertion === ownInsertion && last.to === start) last.to = end;
		else segments.push({ from: start, to: end, ownInsertion });
		return true;
	});
	return segments;
}

/**
 * Replays one `ReplaceStep` so that removed content the current author didn't just insert stays in
 * the document with a `deletion` mark instead of being removed, and newly inserted content gets an
 * `insertion` mark. Deleting one's own still-pending insertion removes it outright (Word's behavior).
 */
function applyTrackedReplace(transform: Transform, step: ReplaceStep, author: string): void {
	const { from, to, slice } = step;
	const segments = segmentsOf(transform.doc, from, to, author);
	const keptLength = segments
		.filter((segment) => !segment.ownInsertion)
		.reduce((sum, segment) => sum + (segment.to - segment.from), 0);
	const insertAt = from + keptLength;
	for (let i = segments.length - 1; i >= 0; i--) {
		const segment = segments[i];
		if (segment.ownInsertion) transform.delete(segment.from, segment.to);
		else
			transform.addMark(
				segment.from,
				segment.to,
				schema.marks.deletion.create({ author, id: nextRevisionId() }),
			);
	}
	if (slice.size) {
		const mark = schema.marks.insertion.create({ author, id: nextRevisionId() });
		transform.replace(
			insertAt,
			insertAt,
			new Slice(withMark(slice.content, mark), slice.openStart, slice.openEnd),
		);
	}
}

/**
 * When enabled, rewrites local editing transactions so insertions/deletions become tracked-change
 * marks (see applyTrackedReplace) instead of directly changing the visible text. Handles the common
 * case of a transaction made of `ReplaceStep`s only (typing, IME, backspace/delete, cut, paste);
 * anything else (structural table edits, attribute changes, remote collaboration steps) passes
 * through untouched.
 */
export function trackChangesPlugin(getAuthor: () => string, isEnabled: () => boolean): Plugin {
	return new Plugin({
		key: trackChangesPluginKey,
		appendTransaction(transactions, oldState, newState) {
			if (!isEnabled()) return null;
			if (
				transactions.some(
					(tr) => tr.getMeta(trackChangesPluginKey) || tr.getMeta(REMOTE_TRANSACTION_META),
				)
			)
				return null;
			const relevant = transactions.filter((tr) => tr.docChanged);
			if (!relevant.length) return null;
			const steps = relevant.flatMap((tr) => tr.steps);
			if (!steps.length || !steps.every((step) => step instanceof ReplaceStep)) return null;
			const author = getAuthor() || 'Author';
			const transform = new Transform(oldState.doc);
			for (const step of steps) applyTrackedReplace(transform, step as ReplaceStep, author);
			if (!transform.docChanged) return null;
			// newState already contains the untracked edit: undo it first so the tracked steps,
			// computed against oldState.doc, apply to the document they were derived from.
			const result = newState.tr;
			const applied = relevant.flatMap((tr) =>
				tr.steps.map((step, index) => ({ step, doc: tr.docs[index] })),
			);
			for (let index = applied.length - 1; index >= 0; index--)
				result.step(applied[index].step.invert(applied[index].doc));
			for (const step of transform.steps) result.step(step);
			result.setMeta(trackChangesPluginKey, true);
			return result;
		},
	});
}
