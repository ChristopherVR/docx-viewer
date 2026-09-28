import type { DocumentModel } from '@christophervr/docx-core';
import { history } from 'prosemirror-history';
import { keymap } from 'prosemirror-keymap';
import type { Plugin } from 'prosemirror-state';
import { editorKeymap } from './editor-commands';
import { fieldGuardPlugin } from './field-guard';
import { tabStopsPlugin } from './tab-stops-view';
import { noteNumberingPlugin } from './note-commands';
import { paragraphStylesPlugin } from './paragraph-styles';
import { reviewDisplayPlugin, type ReviewDisplayMode } from './review-display';
import { runStylesPlugin } from './run-styles';
import { sectionBreaksPlugin } from './section-commands';
import { trackChangesPlugin } from './track-changes-mode';

/** What the body editor's plugins read from the editor element. */
export interface BodyPluginHost {
	model(): DocumentModel;
	reviewAuthor(): string;
	reviewDisplayMode(): ReviewDisplayMode;
	insertNote(kind: 'footnote' | 'endnote'): void;
	showSearch(): void;
	/** Controller and collaboration plugins, in the order they apply. */
	extraPlugins: Plugin[];
	collaborationPlugins: Plugin[];
}

/** The main body editor's plugins: history, structure guards, style decorations, review and keys. */
export function bodyPlugins(host: BodyPluginHost): Plugin[] {
	return [
		history(),
		...host.extraPlugins,
		noteNumberingPlugin(),
		sectionBreaksPlugin(),
		fieldGuardPlugin(),
		tabStopsPlugin(),
		keymap({
			'Mod-Alt-f': () => (host.insertNote('footnote'), true),
			'Mod-Alt-d': () => (host.insertNote('endnote'), true),
		}),
		runStylesPlugin(() => host.model()),
		paragraphStylesPlugin(() => host.model()),
		trackChangesPlugin(
			() => host.reviewAuthor(),
			() => Boolean(host.model().trackChanges),
		),
		reviewDisplayPlugin(() => host.reviewDisplayMode()),
		editorKeymap(() => host.showSearch()),
		...host.collaborationPlugins,
	];
}
