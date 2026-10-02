import { DOMSerializer } from 'prosemirror-model';
import type { Block, DocumentModel, HeaderFooterContent, HeaderFooterSlots } from 'docx-core';
import { schema } from './schema';
import { runToInlineNodes } from './run-adapter';
import { translateUiText } from './localization';
import { effectiveSlots } from './header-footer-link';
import { stylePreviewRuns } from './preview-run-styles';

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

export function renderBlocks(blocks: Block[], model?: DocumentModel): DocumentFragment {
	const serializer = DOMSerializer.fromSchema(schema);
	const fragment = document.createDocumentFragment();
	for (const block of blocks) {
		const element = serializer.serializeNode(blockNode(block)) as HTMLElement;
		if (model) stylePreviewRuns(element, block, model, serializer);
		fragment.append(element);
	}
	return fragment;
}

function slot(
	root: HTMLElement,
	labelText: string,
	content: HeaderFooterContent | undefined,
	name: 'default' | 'first' | 'even',
	model: DocumentModel,
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
	body.append(renderBlocks(content.blocks, model));
	wrap.append(label, body);
	root.append(wrap);
}

function buildSlotsElement(
	slots: HeaderFooterSlots | undefined,
	kind: 'header' | 'footer',
	locale: string,
	model: DocumentModel,
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
		model,
	);
	slot(
		root,
		translateUiText(root, kind === 'header' ? 'Even page header' : 'Even page footer'),
		slots.even,
		'even',
		model,
	);
	slot(
		root,
		translateUiText(root, kind === 'header' ? 'Header' : 'Footer'),
		slots.default,
		'default',
		model,
	);
	root.dataset.editorLocale = locale;
	return root;
}

/** The selected section drives the header/footer preview on this continuous surface. */
export function buildHeaderElement(
	model: DocumentModel,
	locale: string,
	sectionIndex = 0,
): HTMLElement | null {
	return buildSlotsElement(effectiveSlots(model, sectionIndex, 'headers'), 'header', locale, model);
}
export function buildFooterElement(
	model: DocumentModel,
	locale: string,
	sectionIndex = 0,
): HTMLElement | null {
	return buildSlotsElement(effectiveSlots(model, sectionIndex, 'footers'), 'footer', locale, model);
}
