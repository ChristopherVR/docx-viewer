/** Strings for keyboard shortcuts, the shortcut help dialog and the context menu. */
const en = {
	'shortcut.find': 'Find',
	'shortcut.replace': 'Find and replace',
	'shortcut.save': 'Save',
	'shortcut.print': 'Print',
	'shortcut.nextRegion': 'Move focus to the next region (ribbon, document, status bar)',
	'shortcut.previousRegion': 'Move focus to the previous region',
	'shortcut.focusRibbon': 'Move focus to the ribbon',
	'shortcut.escape': 'Close the open panel, dialog or menu and return to the document',
	'shortcut.help': 'Show keyboard shortcuts',
	'shortcut.contextMenu': 'Open the context menu',
	'shortcut.lineBreak': 'Insert a line break',
	'shortcut.dialogTitle': 'Keyboard shortcuts',
	'shortcut.dialogNote':
		'Shortcuts follow common word processor conventions. The editor does not claim identical behavior to Word.',
	'shortcut.columnKeys': 'Shortcut',
	'shortcut.columnAction': 'Action',
	'shortcut.close': 'Close',
	'menu.context': 'Context menu',
	'menu.cut': 'Cut',
	'menu.copy': 'Copy',
	'menu.paste': 'Paste',
	'menu.editLink': 'Edit link',
	'menu.pasteUnavailable':
		'Paste is not available: the browser does not allow reading the clipboard.',
	'menu.clipboardDenied':
		'The browser blocked access to the clipboard. Use the keyboard shortcut for this action instead.',
} as const;

const fr: Record<keyof typeof en, string> = {
	'shortcut.find': 'Rechercher',
	'shortcut.replace': 'Rechercher et remplacer',
	'shortcut.save': 'Enregistrer',
	'shortcut.print': 'Imprimer',
	'shortcut.nextRegion': 'Passer à la zone suivante (ruban, document, barre d’état)',
	'shortcut.previousRegion': 'Passer à la zone précédente',
	'shortcut.focusRibbon': 'Placer le focus sur le ruban',
	'shortcut.escape': 'Fermer le panneau, la boîte de dialogue ou le menu et revenir au document',
	'shortcut.help': 'Afficher les raccourcis clavier',
	'shortcut.contextMenu': 'Ouvrir le menu contextuel',
	'shortcut.lineBreak': 'Insérer un saut de ligne',
	'shortcut.dialogTitle': 'Raccourcis clavier',
	'shortcut.dialogNote':
		'Les raccourcis suivent les conventions courantes des traitements de texte. L’éditeur ne prétend pas se comporter comme Word.',
	'shortcut.columnKeys': 'Raccourci',
	'shortcut.columnAction': 'Action',
	'shortcut.close': 'Fermer',
	'menu.context': 'Menu contextuel',
	'menu.cut': 'Couper',
	'menu.copy': 'Copier',
	'menu.paste': 'Coller',
	'menu.editLink': 'Modifier le lien',
	'menu.pasteUnavailable':
		'Coller est indisponible : le navigateur ne permet pas de lire le presse-papiers.',
	'menu.clipboardDenied':
		'Le navigateur a bloqué l’accès au presse-papiers. Utilisez plutôt le raccourci clavier de cette action.',
};

export const keyboardStrings = { en, fr } as const;
