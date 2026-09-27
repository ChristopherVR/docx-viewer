import { DOMSerializer } from 'prosemirror-model';
import type {
	Block,
	DocumentModel,
	HeaderFooterContent,
	HeaderFooterSlots,
} from '@christophervr/docx-core';
import { schema } from './schema';
import { runToInlineNodes } from './run-adapter';
import { translateUiText } from './localization';

function blockNode(block: Block): ReturnType<typeof schema.node> {
	if (block.type === 'paragraph') {
		const children = block.runs.flatMap((run) => runToInlineNodes(run));
		return schema.node(
			'paragraph',
			{ id: block.id, align: block.align ?? null, direction: block.direction ?? null },
			children,
		);
	}
	const rows = block.rows.map((row) =>
		schema.node(
			'tableRow',
			null,
			row.map((cell) =>
				schema.node(
					'tableCell',
					null,
					cell.paragraphs.map((paragraph) => blockNode(paragraph)),
				),
			),
		),
	);
	return schema.node('table', { id: block.id, structureEditable: false }, rows);
}

export function renderBlocks(blocks: Block[]): DocumentFragment {
	const serializer = DOMSerializer.fromSchema(schema);
	const fragment = document.createDocumentFragment();
	for (const block of blocks) fragment.append(serializer.serializeNode(blockNode(block)));
	return fragment;
}

function slot(
	root: HTMLElement,
	labelText: string,
	content: HeaderFooterContent | undefined,
	name: 'default' | 'first' | 'even',
): void {
	if (!content || !content.blocks.length) return;
	const wrap = document.createElement('div');
	wrap.className = 'dve-header-footer-slot';
	wrap.dataset.slot = name;
	const label = document.createElement('div');
	label.className = 'dve-header-footer-label';
	label.textContent = labelText;
	label.dataset.localeAriaLabel = labelText;
	const body = document.createElement('div');
	body.className = 'dve-header-footer-body';
	body.append(renderBlocks(content.blocks));
	wrap.append(label, body);
	root.append(wrap);
}

function buildSlotsElement(
	slots: HeaderFooterSlots | undefined,
	kind: 'header' | 'footer',
	locale: string,
): HTMLElement | null {
	if (!slots || (!slots.default && !slots.first && !slots.even)) return null;
	const root = document.createElement('section');
	root.className = `dve-${kind}`;
	root.setAttribute('contenteditable', 'false');
	root.dataset.localeAriaLabel = kind === 'header' ? 'Header' : 'Footer';
	root.setAttribute('aria-label', translateUiText(root, kind === 'header' ? 'Header' : 'Footer'));
	slot(
		root,
		translateUiText(root, kind === 'header' ? 'First page header' : 'First page footer'),
		slots.first,
		'first',
	);
	slot(
		root,
		translateUiText(root, kind === 'header' ? 'Even page header' : 'Even page footer'),
		slots.even,
		'even',
	);
	slot(
		root,
		translateUiText(root, kind === 'header' ? 'Header' : 'Footer'),
		slots.default,
		'default',
	);
	root.dataset.editorLocale = locale;
	return root;
}

/** The first section drives the read-only header/footer preview on this continuous surface. */
export function buildHeaderElement(model: DocumentModel, locale: string): HTMLElement | null {
	return buildSlotsElement(model.sections?.[0]?.headers, 'header', locale);
}
export function buildFooterElement(model: DocumentModel, locale: string): HTMLElement | null {
	return buildSlotsElement(model.sections?.[0]?.footers, 'footer', locale);
}
