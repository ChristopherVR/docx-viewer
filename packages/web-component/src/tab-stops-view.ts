import type { TabStop } from '@christophervr/docx-core';
import { placeTab, type LayoutTabStop } from '@christophervr/docx-layout';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import { Plugin, PluginKey, type EditorState } from 'prosemirror-state';
import { Decoration, DecorationSet, type EditorView } from 'prosemirror-view';

const key = new PluginKey<DecorationSet>('dve-tab-stops');
/** Measuring passes after one document change; tabs settle in one or two. */
const MAX_PASSES = 4;

interface TabBox {
	pos: number;
	widthPx: number;
	leader?: string;
}

const stopsOf = (paragraph: ProseMirrorNode): LayoutTabStop[] =>
	((paragraph.attrs.tabStops as TabStop[] | null) ?? []).map((stop) => ({
		posPx: stop.posTwips / 15,
		align: stop.align,
		...(stop.leader ? { leader: stop.leader } : {}),
	}));

/** The text margin (left edge of the paragraph's container content box) in client pixels. */
function textMargin(view: EditorView, paragraph: HTMLElement, scale: number): number {
	const container = paragraph.parentElement ?? view.dom;
	const style = getComputedStyle(container);
	const left = container.getBoundingClientRect().left;
	return left + (parseFloat(style.paddingLeft) + parseFloat(style.borderLeftWidth || '0')) * scale;
}

/** Measures every tab in the document and where Word's tab stops would end it, in CSS pixels. */
function measureTabs(view: EditorView): TabBox[] {
	const boxes: TabBox[] = [];
	const rect = view.dom.getBoundingClientRect();
	const scale = view.dom.offsetWidth ? rect.width / view.dom.offsetWidth : 1;
	view.state.doc.descendants((node, pos) => {
		if (!node.isTextblock) return true;
		if (!node.textContent.includes('\t')) return false;
		const dom = view.nodeDOM(pos);
		if (!(dom instanceof HTMLElement)) return false;
		const margin = textMargin(view, dom, scale);
		const stops = stopsOf(node);
		const tabs: number[] = [];
		node.forEach((child, offset) => {
			if (!child.isText) return;
			const text = child.text ?? '';
			for (let index = text.indexOf('\t'); index >= 0; index = text.indexOf('\t', index + 1))
				tabs.push(pos + 1 + offset + index);
		});
		const end = pos + node.nodeSize - 1;
		tabs.forEach((tab, index) => {
			const start = view.coordsAtPos(tab, 1);
			const after = view.coordsAtPos(tab + 1, 1);
			const segmentEnd = view.coordsAtPos(tabs[index + 1] ?? end, -1);
			// Text after the tab only counts when it stays on the tab's line.
			const followingPx =
				Math.abs(segmentEnd.top - after.top) < 2 ? (segmentEnd.left - after.left) / scale : 0;
			const placement = placeTab(
				(start.left - margin) / scale,
				stops,
				undefined,
				Math.max(0, followingPx),
				Math.max(0, followingPx),
			);
			boxes.push({
				pos: tab,
				widthPx: Math.round(placement.widthPx * 10) / 10,
				...(placement.leader ? { leader: placement.leader } : {}),
			});
		});
		return false;
	});
	return boxes;
}

function decorations(state: EditorState, boxes: TabBox[]): DecorationSet {
	return DecorationSet.create(
		state.doc,
		boxes.map((box) =>
			Decoration.inline(
				box.pos,
				box.pos + 1,
				{
					class: box.leader ? `dve-tab dve-tab-leader-${box.leader}` : 'dve-tab',
					style: `width:${box.widthPx}px`,
				},
				{ widthPx: box.widthPx, leader: box.leader },
			),
		),
	);
}

function sameBoxes(set: DecorationSet, boxes: TabBox[], state: EditorState): boolean {
	const current = set.find(0, state.doc.content.size);
	return (
		current.length === boxes.length &&
		current.every((decoration, index) => {
			const box = boxes[index];
			const spec = decoration.spec as { widthPx: number; leader?: string };
			return (
				decoration.from === box.pos && spec.widthPx === box.widthPx && spec.leader === box.leader
			);
		})
	);
}

/**
 * Sizes tabs on the editing surface at Word's tab stops (custom left/right/center/decimal stops
 * with leaders, then default half-inch stops) by measuring the rendered text, since CSS tab sizes
 * cannot follow per-paragraph stops. Display only: nothing enters the document or its history.
 */
export function tabStopsPlugin(): Plugin<DecorationSet> {
	return new Plugin<DecorationSet>({
		key,
		state: {
			init: () => DecorationSet.empty,
			apply: (tr, set) => tr.getMeta(key) ?? set.map(tr.mapping, tr.doc),
		},
		props: { decorations: (state) => key.getState(state) },
		view(view) {
			let frame = 0;
			let passes = 0;
			const measure = () => {
				frame = 0;
				if (!view.dom.isConnected || passes >= MAX_PASSES) return;
				let boxes: TabBox[];
				try {
					boxes = measureTabs(view);
				} catch {
					return; // No layout (e.g. a detached or headless DOM).
				}
				const set = key.getState(view.state) ?? DecorationSet.empty;
				if (sameBoxes(set, boxes, view.state)) return;
				passes++;
				view.dispatch(
					view.state.tr.setMeta(key, decorations(view.state, boxes)).setMeta('addToHistory', false),
				);
			};
			const schedule = () => {
				if (!frame && typeof requestAnimationFrame === 'function')
					frame = requestAnimationFrame(measure);
			};
			schedule();
			return {
				update(next, previous) {
					const changed = next.state.doc !== previous.doc;
					if (changed) passes = 0;
					if (changed || key.getState(next.state) !== key.getState(previous)) schedule();
				},
				destroy() {
					if (frame) cancelAnimationFrame(frame);
				},
			};
		},
	});
}
