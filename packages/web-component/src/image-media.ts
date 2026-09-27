import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { EditorView, NodeView } from 'prosemirror-view';
import { schema } from './schema';

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
	const spec = schema.nodes.image.spec.toDOM!(node) as [
		string,
		Record<string, string>,
		...unknown[],
	];
	const [tag, attrs, ...content] = spec;
	const dom = document.createElement(tag);
	for (const [name, value] of Object.entries(attrs)) dom.setAttribute(name, value);
	if (typeof content[0] === 'string') dom.textContent = content[0];
	return dom;
}

/**
 * Renders pictures from package media and, for editable real pictures, adds Word's proportional
 * corner resize handle. The new size is applied as one step when the pointer is released.
 */
export function imageNodeView(cache: ImageMediaCache, options: ImageNodeViewOptions = {}) {
	return (node: ProseMirrorNode, view: EditorView, getPos: () => number | undefined): NodeView => {
		const content = renderSpec(node);
		if (content.tagName !== 'IMG') return { dom: content };
		const src = cache.urlFor(String(node.attrs.partName), String(node.attrs.contentType));
		if (src) content.setAttribute('src', src);
		else content.classList.add('dve-image-missing');
		const dom = document.createElement('span');
		dom.className = 'dve-picture';
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
