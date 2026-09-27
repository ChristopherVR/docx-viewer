// Framework-neutral pagination/print-layout engine for @christophervr/docx-core documents.
export * from './units.js';
export * from './measure.js';
export * from './text-breaks.js';
export * from './input.js';
export * from './result.js';
export { layoutParagraph, type ParagraphLayoutResult } from './paragraph-layout.js';
export { adjustForWidowOrphan, suppressesSpacing, widowControlEnabled } from './keep-rules.js';
export { layoutRow, splitRowAtHeight, type RowLayout, type RowSplit } from './table-layout.js';
export { PageCursor } from './page-cursor.js';
export { placeParagraph, type ParagraphPlacement } from './flow-paragraph.js';
export { placeTable } from './flow-table.js';
export { layoutSections } from './page-flow.js';
export { adaptDocumentModel } from './adapter.js';
export { layoutDocument, layoutDocumentModel } from './layout.js';
export { positionFloats, FLOAT_WRAP_NOTE } from './floats.js';
