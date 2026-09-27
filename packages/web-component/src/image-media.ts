import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { NodeView } from 'prosemirror-view';
import { schema } from './schema';

/**
 * Resolves inline picture bytes (kept off the JSON model by docx-core) into object URLs. URLs are
 * cached per package part and revoked together when the owning editor releases the cache.
 */
export class ImageMediaCache {
	private readonly urls = new Map<string, string>();
	constructor(private readonly lookup: () => ReadonlyMap<string, Uint8Array> | undefined) {}

	urlFor(partName: string, contentType: string): string | undefined {
		const cached = this.urls.get(partName);
		if (cached) return cached;
		const bytes = this.lookup()?.get(partName);
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

/** Renders the schema's image DOM and fills in `src` from the package media, when available. */
export function imageNodeView(cache: ImageMediaCache) {
	return (node: ProseMirrorNode): NodeView => {
		const spec = schema.nodes.image.spec.toDOM!(node) as [
			string,
			Record<string, string>,
			...unknown[],
		];
		const [tag, attrs, ...content] = spec;
		const dom = document.createElement(tag);
		for (const [name, value] of Object.entries(attrs)) dom.setAttribute(name, value);
		if (typeof content[0] === 'string') dom.textContent = content[0];
		if (tag === 'img') {
			const src = cache.urlFor(String(node.attrs.partName), String(node.attrs.contentType));
			if (src) dom.setAttribute('src', src);
			else dom.classList.add('dve-image-missing');
		}
		return { dom };
	};
}
