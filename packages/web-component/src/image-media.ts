import { DOMSerializer, type Node as ProseMirrorNode } from 'prosemirror-model';
import type { EditorView, NodeView } from 'prosemirror-view';
import { schema } from './schema';
import { placementClass } from './inline-content-schema';
import { parseDiagram, smartArtNodeView } from './smartart-node-view';
import type { ThemeCatalog } from '@christophervr/docx-core';

/**
 * Resolves inline picture bytes (kept off the JSON model by docx-core) into object URLs. URLs are
 * cached per package part and revoked together when the owning editor releases the cache.
 */
export class ImageMediaCache {
	private readonly urls = new Map<string, string>();
	constructor(private readonly lookup: (partName: string) => Uint8Array | undefined) {}

	urlFor(partName: string, contentType: string): string | undefined {
		const cached = this.urls.get(partName);
		if (cached) return cached;
		const bytes = this.lookup(partName);
		if (!bytes || typeof URL.createObjectURL !== 'function') return undefined;
		const url = URL.createObjectURL(new Blob([bytes.slice()], { type: contentType }));
		this.urls.set(partName, url);
		return url;
	}

	release(): void {
		if (typeof URL.revokeObjectURL === 'function')
			for (const url of this.urls.values()) URL.revokeObjectURL(url);
		this.urls.clear();
	}
}

export interface ImageNodeViewOptions {
	/** Opens Format Picture for the picture at `pos` (double-click). */
	editPicture?(pos: number): void;
	/** The widest a picture may be resized to, in CSS pixels. */
	maxWidth?(): number;
	/** The document theme, for resolving SmartArt colours. */
	theme?(): ThemeCatalog | undefined;
}

const MIN_SIZE = 16;

/** Sets a picture's size in the document, keeping every other attribute. */
export function resizePicture(view: EditorView, pos: number, widthPx: number, heightPx: number) {
	const node = view.state.doc.nodeAt(pos);
	if (!node || node.type !== schema.nodes.image) return;
	view.dispatch(
		view.state.tr.setNodeMarkup(pos, undefined, {
			...node.attrs,
			widthPx: Math.max(MIN_SIZE, Math.round(widthPx)),
			heightPx: Math.max(MIN_SIZE, Math.round(heightPx)),
		}),
	);
}

function renderSpec(node: ProseMirrorNode): HTMLElement {
	// The placeholder spec carries child spans (a text box's lines), so render the whole spec.
	return DOMSerializer.renderSpec(document, schema.nodes.image.spec.toDOM!(node))
		.dom as HTMLElement;
}

/**
 * Renders pictures from package media and, for editable real pictures, adds Word's proportional
 * corner resize handle. The new size is applied as one step when the pointer is released.
 */
export function imageNodeView(cache: ImageMediaCache, options: ImageNodeViewOptions = {}) {
	return (node: ProseMirrorNode, view: EditorView, getPos: () => number | undefined): NodeView => {
		const diagram = parseDiagram(node.attrs.diagram);
		if (diagram) return smartArtNodeView(node, diagram, options.theme?.());
		const content = renderSpec(node);
		// A placeholder (unsupported drawing or text box) shows static child text that ProseMirror must leave alone.
		if (content.tagName !== 'IMG') return { dom: content, ignoreMutation: () => true };
		const src =
			(node.attrs.svgPartName && cache.urlFor(String(node.attrs.svgPartName), 'image/svg+xml')) ||
			cache.urlFor(String(node.attrs.partName), String(node.attrs.contentType));
		if (src) content.setAttribute('src', src);
		else content.classList.add('dve-image-missing');
		const dom = document.createElement('span');
		dom.className = ['dve-picture', placementClass(node.attrs.placement)].filter(Boolean).join(' ');
		dom.append(content);
		content.addEventListener('dblclick', () => {
			const pos = getPos();
			if (pos !== undefined && view.editable) options.editPicture?.(pos);
		});
		const handle = document.createElement('span');
		handle.className = 'dve-picture-handle';
		handle.setAttribute('aria-hidden', 'true');
		handle.addEventListener('pointerdown', (event) => {
			if (!view.editable) return;
			event.preventDefault();
			const startX = event.clientX;
			const startWidth = Number(node.attrs.widthPx);
			const ratio = Number(node.attrs.heightPx) / Math.max(1, startWidth);
			const maxWidth = options.maxWidth?.() ?? Number.POSITIVE_INFINITY;
			let width = startWidth;
			const move = (moveEvent: PointerEvent) => {
				width = Math.min(maxWidth, Math.max(MIN_SIZE, startWidth + moveEvent.clientX - startX));
				content.setAttribute('width', String(Math.round(width)));
				content.setAttribute('height', String(Math.round(width * ratio)));
			};
			const up = () => {
				window.removeEventListener('pointermove', move);
				window.removeEventListener('pointerup', up);
				const pos = getPos();
				if (pos !== undefined && Math.round(width) !== startWidth)
					resizePicture(view, pos, width, width * ratio);
			};
			window.addEventListener('pointermove', move);
			window.addEventListener('pointerup', up);
		});
		dom.append(handle);
		return {
			dom,
			selectNode: () => dom.classList.add('dve-picture-selected'),
			deselectNode: () => dom.classList.remove('dve-picture-selected'),
			stopEvent: (event) => event.target === handle,
			ignoreMutation: () => true,
		};
	};
}
