import { closeHistory } from 'prosemirror-history';
import { Fragment, type Mark, type Node as ProseMirrorNode } from 'prosemirror-model';
import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';

export interface SearchMatch {
	from: number;
	to: number;
}

export interface SearchStatus {
	query: string;
	count: number;
	active: number;
	readOnly: boolean;
}

type Segment = { text: string; from: number };

function escapedLiteral(value: string): string {
	return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function graphemeBoundaries(value: string): Set<number> {
	const boundaries = new Set<number>([0, value.length]);
	if (typeof Intl.Segmenter === 'function') {
		const segmenter = new Intl.Segmenter(undefined, { granularity: 'grapheme' });
		for (const part of segmenter.segment(value)) boundaries.add(part.index);
		return boundaries;
	}
	for (let index = 0; index < value.length;) {
		const point = value.codePointAt(index)!;
		index += point > 0xffff ? 2 : 1;
		while (index < value.length && /\p{Mark}/u.test(value[index])) index++;
		boundaries.add(index);
	}
	return boundaries;
}

function paragraphSegments(node: ProseMirrorNode, pos: number): Segment[] {
	const segments: Segment[] = [];
	let value = '';
	let from = -1;
	const flush = () => {
		if (value) segments.push({ text: value, from });
		value = '';
		from = -1;
	};
	node.forEach((child, offset) => {
		if (!child.isText) {
			flush();
			return;
		}
		if (!child.text) return;
		if (from < 0) from = pos + 1 + offset;
		value += child.text;
	});
	flush();
	return segments;
}

function collectMatches(
	doc: ProseMirrorNode,
	query: string,
	caseSensitive: boolean,
): SearchMatch[] {
	if (!query) return [];
	const expression = new RegExp(escapedLiteral(query), caseSensitive ? 'gu' : 'giu');
	const matches: SearchMatch[] = [];
	doc.descendants((node, pos) => {
		if (node.type.name !== 'paragraph') return;
		for (const segment of paragraphSegments(node, pos)) {
			const boundaries = graphemeBoundaries(segment.text);
			expression.lastIndex = 0;
			for (const match of segment.text.matchAll(expression)) {
				const start = match.index;
				const end = start + match[0].length;
				if (boundaries.has(start) && boundaries.has(end))
					matches.push({ from: segment.from + start, to: segment.from + end });
			}
		}
		return false;
	});
	return matches;
}

function marksAt(doc: ProseMirrorNode, pos: number): readonly Mark[] {
	const $pos = doc.resolve(pos);
	if ($pos.nodeAfter?.isText) return $pos.nodeAfter.marks;
	if ($pos.nodeBefore?.isText) return $pos.nodeBefore.marks;
	return $pos.marks();
}

export class SearchController {
	private query = '';
	private caseSensitive = false;
	private matches: SearchMatch[] = [];
	private activeIndex = -1;

	constructor(private readonly getView: () => EditorView | undefined) {}

	get status(): SearchStatus {
		return {
			query: this.query,
			count: this.matches.length,
			active: this.activeIndex < 0 ? 0 : this.activeIndex + 1,
			readOnly: !this.getView()?.editable,
		};
	}

	search(query: string, caseSensitive = this.caseSensitive): SearchStatus {
		this.query = query;
		this.caseSensitive = caseSensitive;
		this.matches = this.currentMatches();
		this.activeIndex = this.indexAtSelection();
		return this.status;
	}

	refresh(): SearchStatus {
		this.matches = this.currentMatches();
		this.activeIndex = this.indexAtSelection();
		return this.status;
	}

	findNext(): SearchMatch | null {
		const view = this.getView();
		this.refresh();
		if (!view || !this.matches.length) return null;
		let next = this.activeIndex + 1;
		if (this.activeIndex < 0) {
			const cursor = view.state.selection.from;
			next = this.matches.findIndex((match) => match.from >= cursor);
			if (next < 0) next = 0;
		}
		return this.select(next % this.matches.length);
	}

	findPrevious(): SearchMatch | null {
		const view = this.getView();
		this.refresh();
		if (!view || !this.matches.length) return null;
		let previous = this.activeIndex - 1;
		if (this.activeIndex < 0) {
			const cursor = view.state.selection.from;
			previous = -1;
			for (let index = this.matches.length - 1; index >= 0; index--) {
				if (this.matches[index].from < cursor) {
					previous = index;
					break;
				}
			}
			if (previous < 0) previous = this.matches.length - 1;
		}
		return this.select((previous + this.matches.length) % this.matches.length);
	}

	replaceOne(replacement: string): boolean {
		const view = this.getView();
		if (!view?.editable || !this.query) return false;
		this.refresh();
		if (this.activeIndex < 0 && !this.findNext()) return false;
		const match = this.matches[this.activeIndex];
		if (!match) return false;
		this.replace(view, [match], replacement);
		return true;
	}

	replaceAll(replacement: string): number {
		const view = this.getView();
		if (!view?.editable || !this.query) return 0;
		this.refresh();
		if (!this.matches.length) return 0;
		const targets = [...this.matches];
		this.replace(view, targets, replacement);
		return targets.length;
	}

	private currentMatches(): SearchMatch[] {
		const view = this.getView();
		return view ? collectMatches(view.state.doc, this.query, this.caseSensitive) : [];
	}

	private indexAtSelection(): number {
		const selection = this.getView()?.state.selection;
		if (!selection) return -1;
		return this.matches.findIndex(
			(match) => match.from === selection.from && match.to === selection.to,
		);
	}

	private select(index: number): SearchMatch {
		const view = this.getView()!;
		const match = this.matches[index];
		this.activeIndex = index;
		view.dispatch(
			view.state.tr
				.setSelection(TextSelection.create(view.state.doc, match.from, match.to))
				.scrollIntoView(),
		);
		return match;
	}

	private replace(view: EditorView, matches: SearchMatch[], replacement: string): void {
		const { state } = view;
		let transaction = state.tr;
		for (const match of [...matches].sort((left, right) => right.from - left.from)) {
			const marks = marksAt(state.doc, match.from);
			const content = replacementFragment(state.schema, replacement, marks);
			transaction = transaction.replaceWith(match.from, match.to, content);
		}
		const last = matches[0];
		const cursor =
			last.from +
			replacementFragment(state.schema, replacement, marksAt(state.doc, last.from)).size;
		transaction = transaction.setSelection(TextSelection.create(transaction.doc, cursor));
		view.dispatch(closeHistory(transaction.scrollIntoView()));
		this.refresh();
	}
}

function replacementFragment(
	schema: EditorView['state']['schema'],
	replacement: string,
	marks: readonly Mark[],
): Fragment {
	if (!replacement) return Fragment.empty;
	const lines = replacement.split(/\r\n|\r|\n/);
	const content: ProseMirrorNode[] = [];
	lines.forEach((line, index) => {
		if (line) content.push(schema.text(line, marks));
		if (index < lines.length - 1) content.push(schema.nodes.hardBreak.create(null, null, marks));
	});
	return Fragment.fromArray(content);
}

export function findTextMatches(
	doc: ProseMirrorNode,
	query: string,
	caseSensitive = false,
): SearchMatch[] {
	return collectMatches(doc, query, caseSensitive);
}
