import type { EditorView } from 'prosemirror-view';
import type { InlineImage, PendingMediaPart } from 'docx-core';
import { schema } from './schema';

/** Raster formats Word stores directly; SVG would need a PNG fallback part, which is not produced. */
export const PICTURE_TYPES: Record<string, string> = {
	'image/png': 'png',
	'image/jpeg': 'jpeg',
	'image/gif': 'gif',
	'image/bmp': 'bmp',
};

let pictureSerial = 0;

/** A package part name for a newly inserted picture that cannot collide with existing media. */
export function newPicturePartName(contentType: string): string {
	const extension = PICTURE_TYPES[contentType] ?? 'png';
	pictureSerial += 1;
	return `word/media/dve-picture-${Date.now().toString(36)}-${pictureSerial}.${extension}`;
}

/** Scales `width`×`height` down (never up) to fit `maxWidth`, keeping the aspect ratio. */
export function fitPicture(
	width: number,
	height: number,
	maxWidth: number,
): { widthPx: number; heightPx: number } {
	if (width <= 0 || height <= 0)
		return { widthPx: Math.max(1, maxWidth / 2), heightPx: Math.max(1, maxWidth / 2) };
	const scale = Math.min(1, maxWidth / width);
	return { widthPx: Math.round(width * scale), heightPx: Math.round(height * scale) };
}

/** Reads a Blob's bytes, falling back to FileReader where `Blob.arrayBuffer` is unavailable. */
export async function readBytes(file: Blob): Promise<Uint8Array> {
	if (typeof file.arrayBuffer === 'function') return new Uint8Array(await file.arrayBuffer());
	return new Promise((resolve, reject) => {
		const reader = new FileReader();
		reader.onload = () => resolve(new Uint8Array(reader.result as ArrayBuffer));
		reader.onerror = () => reject(reader.error ?? new Error('Could not read the picture.'));
		reader.readAsArrayBuffer(file);
	});
}

/** Reads a picture's natural pixel size in the browser. */
export async function pictureSize(file: Blob): Promise<{ width: number; height: number }> {
	if (typeof createImageBitmap === 'function') {
		const bitmap = await createImageBitmap(file).catch(() => {
			throw new Error('The picture could not be read; the file may be damaged.');
		});
		const size = { width: bitmap.width, height: bitmap.height };
		bitmap.close();
		return size;
	}
	return { width: 0, height: 0 };
}

/** Prepares an inline picture and the media bytes that must be written with it on save. */
/** Draws an SVG to a PNG (Word keeps a raster fallback beside every SVG picture). */
export async function rasterizeSvg(
	svg: Uint8Array,
): Promise<{ png: Uint8Array; width: number; height: number }> {
	const url = URL.createObjectURL(new Blob([svg.slice()], { type: 'image/svg+xml' }));
	try {
		const image = new Image();
		image.src = url;
		await image.decode().catch(() => {
			throw new Error('The SVG picture could not be read; the file may be damaged.');
		});
		const width = image.naturalWidth || 300;
		const height = image.naturalHeight || 150;
		const canvas = document.createElement('canvas');
		canvas.width = width;
		canvas.height = height;
		canvas.getContext('2d')!.drawImage(image, 0, 0, width, height);
		const blob = await new Promise<Blob | null>((resolve) => canvas.toBlob(resolve, 'image/png'));
		if (!blob) throw new Error('The SVG picture could not be converted to PNG.');
		return { png: await readBytes(blob), width, height };
	} finally {
		URL.revokeObjectURL(url);
	}
}

export interface StagedPicture {
	image: InlineImage;
	media: PendingMediaPart;
	/** For SVG pictures: the SVG original, stored alongside the PNG fallback in `media`. */
	svg?: PendingMediaPart;
}

/** Prepares an inline picture and the media bytes that must be written with it on save. */
export async function stagePicture(file: File, maxWidthPx: number): Promise<StagedPicture> {
	const contentType = file.type.toLowerCase();
	const altText = file.name.replace(/\.[^.]+$/, '');
	if (contentType === 'image/svg+xml') {
		const svg = await readBytes(file);
		const { png, width, height } = await rasterizeSvg(svg);
		const size = fitPicture(width, height, maxWidthPx);
		const partName = newPicturePartName('image/png');
		return {
			image: {
				relId: '',
				partName,
				svgPartName: partName.replace(/\.png$/, '.svg'),
				contentType: 'image/png',
				...size,
				...(altText ? { altText } : {}),
			},
			media: { bytes: png, contentType: 'image/png' },
			svg: { bytes: svg, contentType: 'image/svg+xml' },
		};
	}
	if (!PICTURE_TYPES[contentType])
		throw new Error(
			'Insert a PNG, JPEG, GIF, BMP or SVG picture; other image formats are not supported.',
		);
	const bytes = await readBytes(file);
	const size = await pictureSize(file);
	const { widthPx, heightPx } = fitPicture(size.width, size.height, maxWidthPx);
	return {
		image: {
			relId: '',
			partName: newPicturePartName(contentType),
			contentType,
			widthPx,
			heightPx,
			...(altText ? { altText } : {}),
		},
		media: { bytes, contentType },
	};
}

/** Inserts an inline picture at the selection, replacing any selected content. */
export function insertPicture(view: EditorView, image: InlineImage): boolean {
	const node = schema.nodes.image.create({
		relId: image.relId,
		partName: image.partName,
		contentType: image.contentType,
		widthPx: image.widthPx,
		heightPx: image.heightPx,
		altText: image.altText ?? null,
		title: image.title ?? null,
		anchored: false,
		unsupported: null,
		svgPartName: image.svgPartName ?? null,
	});
	view.dispatch(view.state.tr.replaceSelectionWith(node, false).scrollIntoView());
	return true;
}
