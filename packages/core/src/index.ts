// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
export * from './model.js';
export * from './highlight.js';
export { loadDocx } from './parse.js';
export { saveDocx } from './save.js';
