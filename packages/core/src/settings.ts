// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import { onOffElement } from './simple-types.js';
import type JSZip from 'jszip';
import { buildXml, children, getW, makeW, parseXml, type XmlElement, WORD_NS } from './xml.js';
import { ensureContentTypeOverride, ensureDocumentRelationship } from './zip-parts.js';

const SETTINGS_PATH = 'word/settings.xml';
const SETTINGS_CONTENT_TYPE =
	'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml';
const SETTINGS_RELATIONSHIP_TYPE =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings';

/** The leading part of CT_Settings' element sequence, used to insert flags in schema order. */
const SETTINGS_ORDER = [
	'writeProtection',
	'view',
	'zoom',
	'removePersonalInformation',
	'removeDateAndTime',
	'doNotDisplayPageBoundaries',
	'displayBackgroundShape',
	'printPostScriptOverText',
	'printFractionalCharacterWidth',
	'printFormsData',
	'embedTrueTypeFonts',
	'embedSystemFonts',
	'saveSubsetFonts',
	'saveFormsData',
	'mirrorMargins',
	'alignBordersAndEdges',
	'bordersDoNotSurroundHeader',
	'bordersDoNotSurroundFooter',
	'gutterAtTop',
	'hideSpellingErrors',
	'hideGrammaticalErrors',
	'activeWritingStyle',
	'proofState',
	'formsDesign',
	'attachedTemplate',
	'linkStyles',
	'stylePaneFormatFilter',
	'stylePaneSortMethod',
	'documentType',
	'mailMerge',
	'revisionView',
	'trackRevisions',
	'doNotTrackMoves',
	'doNotTrackFormatting',
	'documentProtection',
	'autoFormatOverride',
	'styleLockTheme',
	'styleLockQFSet',
	'defaultTabStop',
	'autoHyphenation',
	'consecutiveHyphenLimit',
	'hyphenationZone',
	'doNotHyphenateCaps',
	'showEnvelope',
	'summaryLength',
	'clickAndTypeStyle',
	'defaultTableStyle',
	'evenAndOddHeaders',
	'bookFoldRevPrinting',
	'bookFoldPrinting',
	'bookFoldPrintingSheets',
	'drawingGridHorizontalSpacing',
	'drawingGridVerticalSpacing',
	'displayHorizontalDrawingGridEvery',
	'displayVerticalDrawingGridEvery',
	'doNotUseMarginsForDrawingGridOrigin',
	'drawingGridHorizontalOrigin',
	'drawingGridVerticalOrigin',
	'doNotShadeFormData',
	'noPunctuationKerning',
	'characterSpacingControl',
	'printTwoOnOne',
	'strictFirstAndLastChars',
	'noLineBreaksAfter',
	'noLineBreaksBefore',
	'savePreviewPicture',
	'doNotValidateAgainstSchema',
	'saveInvalidXml',
	'ignoreMixedContent',
	'alwaysShowPlaceholderText',
	'doNotDemarcateInvalidXml',
	'saveXmlDataOnly',
	'useXSLTWhenSaving',
	'saveThroughXslt',
	'showXMLTags',
	'alwaysMergeEmptyNamespace',
	'updateFields',
	'hdrShapeDefaults',
	'footnotePr',
	'endnotePr',
	'compat',
	'docVars',
	'rsids',
	'mathPr',
	'attachedSchema',
	'themeFontLang',
	'clrSchemeMapping',
	'doNotIncludeSubdocsInStats',
	'doNotAutoCompressPictures',
	'forceUpgrade',
	'captions',
	'readModeInkLockDown',
	'smartTagType',
	'schemaLibrary',
	'shapeDefaults',
	'doNotEmbedSmartTags',
	'decimalSymbol',
	'listSeparator',
];

/** Reads Word's Track Changes toggle (`w:trackRevisions`; `w:trackChanges` is accepted too). */
export function parseTrackChangesSetting(xml: string): boolean {
	const root = parseXml(xml).documentElement as XmlElement | null;
	if (!root) return false;
	return (
		onOffElement(children(root, 'trackRevisions')[0] ?? children(root, 'trackChanges')[0]) === true
	);
}

/**
 * Sets or clears an on/off element in settings.xml, inserting it in schema order. Creates the
 * settings part (with its relationship and content type) the first time it's needed.
 */
export async function applySettingsFlag(zip: JSZip, name: string, enabled: boolean): Promise<void> {
	const file = zip.file(SETTINGS_PATH);
	if (!file && !enabled) return;
	const xml =
		(await file?.async('string')) ??
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="${WORD_NS}"></w:settings>`;
	const doc = parseXml(xml);
	const root = doc.documentElement as XmlElement;
	for (const element of children(root, name)) root.removeChild(element);
	if (enabled) {
		const rank = SETTINGS_ORDER.indexOf(name);
		const after = Array.from(root.childNodes).find((node) => {
			const index = SETTINGS_ORDER.indexOf((node as XmlElement).localName ?? '');
			return node.nodeType === 1 && index > rank;
		});
		root.insertBefore(makeW(doc, name), after ?? null);
	}
	zip.file(SETTINGS_PATH, buildXml(doc));
	if (!file) {
		await ensureContentTypeOverride(zip, SETTINGS_PATH, SETTINGS_CONTENT_TYPE);
		await ensureDocumentRelationship(zip, SETTINGS_RELATIONSHIP_TYPE, 'settings.xml');
	}
}

/** Word's Review > Track Changes toggle, written as `w:trackRevisions`. */
export async function applyTrackChangesSetting(zip: JSZip, enabled: boolean): Promise<void> {
	const file = zip.file(SETTINGS_PATH);
	// Remove the non-standard element earlier versions of this editor wrote.
	if (file) {
		const doc = parseXml(await file.async('string'));
		const root = doc.documentElement as XmlElement;
		const legacy = children(root, 'trackChanges');
		if (legacy.length) {
			for (const element of legacy) root.removeChild(element);
			zip.file(SETTINGS_PATH, buildXml(doc));
		}
	}
	await applySettingsFlag(zip, 'trackRevisions', enabled);
}
