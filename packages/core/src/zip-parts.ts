// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Minimal string-level surgery for OPC package plumbing, used only to create parts (settings.xml,
// comments.xml...) that a source package may lack. The generic helpers are in `ooxml-opc`.
import { ensureRelationship } from '@christophervr/ooxml-core/opc';
import type JSZip from 'jszip';

export { ensureContentTypeOverride } from '@christophervr/ooxml-core/opc';

const DOCUMENT_RELS_PATH = 'word/_rels/document.xml.rels';

/** Returns the relationship id for `target`, creating the relationship (and rels part) if needed. */
export const ensureDocumentRelationship = (zip: JSZip, type: string, target: string) =>
	ensureRelationship(zip, DOCUMENT_RELS_PATH, type, target);
