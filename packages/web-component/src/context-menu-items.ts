/** Pure context-menu model: which items appear, and which are disabled, for a given editor state. */
import type { EditorView } from 'prosemirror-view';
import { linkAtSelection } from './link-commands';
import type { LocalizationKey } from './localization';
import type { RibbonAction } from './ribbon-action';
import { schema } from './schema';
import { canExecuteTableCommand, type TableCommand } from './table-commands';

export type MenuCommand =
	| { kind: 'clipboard'; op: 'cut' | 'copy' | 'paste' }
	| { kind: 'removeLink' }
	| { kind: 'ribbon'; action: RibbonAction };

export interface MenuItem {
	id: string;
	label: LocalizationKey;
	command: MenuCommand;
	disabled: boolean;
	/** Set when disabled because of the platform, so the UI can say why. */
	reason?: LocalizationKey;
	/** Starts a new group; rendered as a separator. */
	separatorBefore: boolean;
}

export interface MenuState {
	readOnly: boolean;
	hasSelection: boolean;
	onLink: boolean;
	inTable: boolean;
	/** The browser exposes an async clipboard read API. */
	canReadClipboard: boolean;
	canTable(command: TableCommand): boolean;
}

const TABLE_ITEMS: [TableCommand, LocalizationKey][] = [
	['rowBefore', 'Insert row above'],
	['rowAfter', 'Insert row below'],
	['deleteRow', 'Delete row'],
	['columnBefore', 'Insert column left'],
	['columnAfter', 'Insert column right'],
	['deleteColumn', 'Delete column'],
	['deleteTable', 'Delete table'],
];

/** Items for the current state; table items only exist while the selection is in a table. */
export function computeMenuItems(state: MenuState): MenuItem[] {
	const groups: Omit<MenuItem, 'separatorBefore'>[][] = [];
	const clip = (op: 'cut' | 'copy' | 'paste', label: LocalizationKey, disabled: boolean) => ({
		id: op,
		label,
		command: { kind: 'clipboard', op } as const,
		disabled,
		...(op === 'paste' && !state.canReadClipboard && !state.readOnly
			? { reason: 'menu.pasteUnavailable' as const }
			: {}),
	});
	groups.push([
		clip('cut', 'menu.cut', state.readOnly || !state.hasSelection),
		clip('copy', 'menu.copy', !state.hasSelection),
		clip('paste', 'menu.paste', state.readOnly || !state.canReadClipboard),
	]);
	const link = (id: string, label: LocalizationKey, command: MenuCommand) => ({
		id,
		label,
		command,
		disabled: state.readOnly,
	});
	const insertLink = { kind: 'ribbon', action: { type: 'link' } } as const;
	groups.push(
		state.onLink
			? [
					link('link-edit', 'menu.editLink', insertLink),
					link('link-remove', 'Remove link', { kind: 'removeLink' }),
				]
			: [link('link-insert', 'Insert link', insertLink)],
	);
	groups.push([
		link('comment-add', 'Add comment', {
			kind: 'ribbon',
			action: { type: 'comments', key: 'add' },
		}),
	]);
	if (state.inTable)
		groups.push(
			TABLE_ITEMS.map(([key, label]) => ({
				id: `table-${key}`,
				label,
				command: { kind: 'ribbon', action: { type: 'tableEdit', key } } as const,
				disabled: state.readOnly || !state.canTable(key),
			})),
		);
	return groups.flatMap((group, index) =>
		group.map((item, position) => ({ ...item, separatorBefore: index > 0 && position === 0 })),
	);
}

function selectionInTable(view: EditorView): boolean {
	const { $from } = view.state.selection;
	for (let depth = $from.depth; depth > 0; depth--)
		if ($from.node(depth).type === schema.nodes.table) return true;
	return false;
}

export function menuStateFromView(view: EditorView, readOnly: boolean): MenuState {
	return {
		readOnly: readOnly || !view.editable,
		hasSelection: !view.state.selection.empty,
		onLink: Boolean(linkAtSelection(view)),
		inTable: selectionInTable(view),
		canReadClipboard:
			typeof navigator !== 'undefined' &&
			(typeof navigator.clipboard?.read === 'function' ||
				typeof navigator.clipboard?.readText === 'function'),
		canTable: (command) => canExecuteTableCommand(view, command),
	};
}

/** Keeps a menu of the given size inside the viewport, preferring the requested corner. */
export function clampToViewport(
	x: number,
	y: number,
	size: { width: number; height: number },
	viewport: { width: number; height: number },
	margin = 4,
): { left: number; top: number } {
	const fit = (position: number, extent: number, limit: number) =>
		Math.max(margin, Math.min(position, limit - extent - margin));
	return {
		left: fit(x, size.width, viewport.width),
		top: fit(y, size.height, viewport.height),
	};
}
