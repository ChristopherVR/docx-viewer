/** Display-language strings for the shared editor UI. This locale never changes DOCX language marks. */
export type EditorLocale = 'en' | 'fr';

const strings = {
	en: {
		Font: 'Font',
		Paragraph: 'Paragraph',
		'Alignment and history': 'Alignment and history',
		Editing: 'Editing',
		Tables: 'Tables',
		'Page setup': 'Page setup',
		Styles: 'Styles',
		'Document formatting': 'Document formatting',
		'Ribbon tabs': 'Ribbon tabs',
		Home: 'Home',
		Insert: 'Insert',
		Layout: 'Layout',
		Review: 'Review',
		View: 'View',
		Table: 'Table',
		'Font family': 'Font family',
		'Font size': 'Font size',
		Bold: 'Bold',
		Italic: 'Italic',
		Underline: 'Underline',
		Strikethrough: 'Strikethrough',
		Superscript: 'Superscript',
		Subscript: 'Subscript',
		'Font color': 'Font color',
		'Text highlight': 'Text highlight',
		'No highlight': 'No highlight',
		'Clear formatting': 'Clear formatting',
		Clear: 'Clear',
		Black: 'Black',
		Red: 'Red',
		Orange: 'Orange',
		Gold: 'Gold',
		Green: 'Green',
		Blue: 'Blue',
		Purple: 'Purple',
		Yellow: 'Yellow',
		Cyan: 'Cyan',
		Magenta: 'Magenta',
		White: 'White',
		'Dark blue': 'Dark blue',
		'Dark cyan': 'Dark cyan',
		'Dark green': 'Dark green',
		'Dark magenta': 'Dark magenta',
		'Dark red': 'Dark red',
		'Dark yellow': 'Dark yellow',
		'Dark gray': 'Dark gray',
		'Light gray': 'Light gray',
		Lists: 'Lists',
		'Bulleted list': 'Bulleted list',
		'Numbered list': 'Numbered list',
		'Decrease list level': 'Decrease list level',
		'Increase list level': 'Increase list level',
		'Remove list': 'Remove list',
		'Insert row above': 'Insert row above',
		'Insert row below': 'Insert row below',
		'Delete row': 'Delete row',
		'Insert column left': 'Insert column left',
		'Insert column right': 'Insert column right',
		'Delete column': 'Delete column',
		'Delete table': 'Delete table',
		'Spacing after': 'Spacing after',
		'Spacing before': 'Spacing before',
		'Line spacing': 'Line spacing',
		'Before: style default': 'Before: style default',
		'After: style default': 'After: style default',
		'Before: no paragraph space': 'Before: no paragraph space',
		'After: no paragraph space': 'After: no paragraph space',
		'Align left': 'Align left',
		'Align center': 'Align center',
		'Align right': 'Align right',
		Justify: 'Justify',
		Undo: 'Undo',
		Redo: 'Redo',
		'Find and replace': 'Find and replace',
		'Insert table': 'Insert table',
		Normal: 'Normal',
		Narrow: 'Narrow',
		Wide: 'Wide',
		Margins: 'Margins',
		Orientation: 'Orientation',
		Portrait: 'Portrait',
		Landscape: 'Landscape',
		Zoom: 'Zoom',
		Inherit: 'Inherit',
		'Direction: inherit': 'Direction: inherit',
		'Left to right': 'Left to right',
		'Right to left': 'Right to left',
		'Paragraph direction': 'Paragraph direction',
		'Run direction': 'Run direction',
		'Run direction: inherit': 'Run direction: inherit',
		'Run right to left': 'Run right to left',
		'Run explicit non-RTL': 'Run explicit non-RTL',
		'Text language': 'Text language',
		'East Asian language': 'East Asian language',
		'Complex script language': 'Complex script language',
		'Custom BCP 47 tag': 'Custom BCP 47 tag',
		'Language tags annotate text; they do not translate or spell-check it.':
			'Language tags annotate text; they do not translate or spell-check it.',
		'Find text': 'Find text',
		'Replace with': 'Replace with',
		'Match case': 'Match case',
		'Find previous': 'Find previous',
		'Find next': 'Find next',
		Replace: 'Replace',
		'Replace all': 'Replace all',
		'Close search': 'Close search',
		'Search results': 'Search results',
		'Enter text to search': 'Enter text to search',
		'No matches': '0 matches',
		'matches one': 'match',
		'matches many': 'matches',
		of: 'of',
		Mixed: 'Mixed',
		'Mixed paragraphs': 'Mixed paragraphs',
		'Line: style default': 'Line: style default',
		'Style default': 'Style default',
		'styles.label': 'Style',
		'styles.inherit': 'Style default',
		'status.words': '{count} words',
		Breaks: 'Breaks',
		'Insert page break': 'Insert page break',
		'Insert column break': 'Insert column break',
		'Page break': 'Page break',
		'Column break': 'Column break',
		Header: 'Header',
		Footer: 'Footer',
		'First page header': 'First page header',
		'First page footer': 'First page footer',
		'Even page header': 'Even page header',
		'Even page footer': 'Even page footer',
		Footnotes: 'Footnotes',
		Endnotes: 'Endnotes',
		'Headers and footers are read-only.': 'Headers and footers are read-only.',
		'Layout view': 'Layout view',
		Draft: 'Draft',
		'Print Layout': 'Print Layout',
		Print: 'Print',
		'status.page': 'Page {current} of {total}',
	},
	fr: {
		Font: 'Police',
		Paragraph: 'Paragraphe',
		'Alignment and history': 'Alignement et historique',
		Editing: 'Édition',
		Tables: 'Tableaux',
		'Page setup': 'Mise en page',
		Styles: 'Styles',
		'Document formatting': 'Mise en forme du document',
		'Ribbon tabs': 'Onglets du ruban',
		Home: 'Accueil',
		Insert: 'Insertion',
		Layout: 'Disposition',
		Review: 'Révision',
		View: 'Affichage',
		Table: 'Tableau',
		'Font family': 'Police',
		'Font size': 'Taille de police',
		Bold: 'Gras',
		Italic: 'Italique',
		Underline: 'Souligné',
		Strikethrough: 'Barré',
		Superscript: 'Exposant',
		Subscript: 'Indice',
		'Font color': 'Couleur de police',
		'Text highlight': 'Surlignage',
		'No highlight': 'Sans surlignage',
		'Clear formatting': 'Effacer la mise en forme',
		Clear: 'Effacer',
		Black: 'Noir',
		Red: 'Rouge',
		Orange: 'Orange',
		Gold: 'Or',
		Green: 'Vert',
		Blue: 'Bleu',
		Purple: 'Violet',
		Yellow: 'Jaune',
		Cyan: 'Cyan',
		Magenta: 'Magenta',
		White: 'Blanc',
		'Dark blue': 'Bleu foncé',
		'Dark cyan': 'Cyan foncé',
		'Dark green': 'Vert foncé',
		'Dark magenta': 'Magenta foncé',
		'Dark red': 'Rouge foncé',
		'Dark yellow': 'Jaune foncé',
		'Dark gray': 'Gris foncé',
		'Light gray': 'Gris clair',
		Lists: 'Listes',
		'Bulleted list': 'Liste à puces',
		'Numbered list': 'Liste numérotée',
		'Decrease list level': 'Réduire le niveau de liste',
		'Increase list level': 'Augmenter le niveau de liste',
		'Remove list': 'Supprimer la liste',
		'Insert row above': 'Insérer une ligne au-dessus',
		'Insert row below': 'Insérer une ligne au-dessous',
		'Delete row': 'Supprimer la ligne',
		'Insert column left': 'Insérer une colonne à gauche',
		'Insert column right': 'Insérer une colonne à droite',
		'Delete column': 'Supprimer la colonne',
		'Delete table': 'Supprimer le tableau',
		'Spacing after': 'Espacement après',
		'Spacing before': 'Espacement avant',
		'Line spacing': 'Interligne',
		'Before: style default': 'Avant : valeur du style',
		'After: style default': 'Après : valeur du style',
		'Before: no paragraph space': 'Avant : aucun espacement',
		'After: no paragraph space': 'Après : aucun espacement',
		'Align left': 'Aligner à gauche',
		'Align center': 'Centrer',
		'Align right': 'Aligner à droite',
		Justify: 'Justifier',
		Undo: 'Annuler',
		Redo: 'Rétablir',
		'Find and replace': 'Rechercher et remplacer',
		'Insert table': 'Insérer un tableau',
		Normal: 'Normales',
		Narrow: 'Étroites',
		Wide: 'Larges',
		Margins: 'Marges',
		Orientation: 'Orientation',
		Portrait: 'Portrait',
		Landscape: 'Paysage',
		Zoom: 'Zoom',
		Inherit: 'Hériter',
		'Direction: inherit': 'Direction : hériter',
		'Left to right': 'De gauche à droite',
		'Right to left': 'De droite à gauche',
		'Paragraph direction': 'Direction du paragraphe',
		'Run direction': 'Direction du texte',
		'Run direction: inherit': 'Direction du texte : hériter',
		'Run right to left': 'Texte de droite à gauche',
		'Run explicit non-RTL': 'Texte explicitement de gauche à droite',
		'Text language': 'Langue du texte',
		'East Asian language': 'Langue est-asiatique',
		'Complex script language': 'Langue des écritures complexes',
		'Custom BCP 47 tag': 'Balise BCP 47 personnalisée',
		'Language tags annotate text; they do not translate or spell-check it.':
			'Les balises indiquent la langue du texte ; elles ne le traduisent pas et ne vérifient pas son orthographe.',
		'Find text': 'Rechercher le texte',
		'Replace with': 'Remplacer par',
		'Match case': 'Respecter la casse',
		'Find previous': 'Résultat précédent',
		'Find next': 'Résultat suivant',
		Replace: 'Remplacer',
		'Replace all': 'Tout remplacer',
		'Close search': 'Fermer la recherche',
		'Search results': 'Résultats de recherche',
		'Enter text to search': 'Saisissez le texte à rechercher',
		'No matches': 'Aucun résultat',
		'matches one': 'résultat',
		'matches many': 'résultats',
		of: 'sur',
		Mixed: 'Multiple',
		'Mixed paragraphs': 'Paragraphes multiples',
		'Line: style default': 'Interligne : valeur du style',
		'Style default': 'Valeur du style',
		'styles.label': 'Style',
		'styles.inherit': 'Style par défaut',
		'status.words': '{count} mots',
		Breaks: 'Sauts',
		'Insert page break': 'Insérer un saut de page',
		'Insert column break': 'Insérer un saut de colonne',
		'Page break': 'Saut de page',
		'Column break': 'Saut de colonne',
		Header: 'En-tête',
		Footer: 'Pied de page',
		'First page header': 'En-tête de première page',
		'First page footer': 'Pied de page de première page',
		'Even page header': 'En-tête de page paire',
		'Even page footer': 'Pied de page de page paire',
		Footnotes: 'Notes de bas de page',
		Endnotes: 'Notes de fin',
		'Headers and footers are read-only.': 'Les en-têtes et pieds de page sont en lecture seule.',
		'Layout view': 'Type d’affichage',
		Draft: 'Brouillon',
		'Print Layout': 'Mise en page à l’impression',
		Print: 'Imprimer',
		'status.page': 'Page {current} sur {total}',
	},
} as const;

export type LocalizationKey = keyof typeof strings.en;
export function normalizeEditorLocale(value: string | null | undefined): EditorLocale {
	return value?.trim().toLowerCase().split(/[-_]/)[0] === 'fr' ? 'fr' : 'en';
}
export function translate(locale: EditorLocale, key: LocalizationKey): string {
	return strings[locale][key] ?? strings.en[key] ?? key;
}
function translateDynamic(locale: EditorLocale, text: string): string {
	if (locale === 'en') return text;
	const line = /^(\d+(?:\.\d+)?) lines$/.exec(text);
	if (line) return `${line[1]} ${Number(line[1]) === 1 ? 'ligne' : 'lignes'}`;
	const automatic = /^Automatic (.+) lines \((.+)\)$/.exec(text);
	if (automatic) return `Automatique ${automatic[1]} lignes (${automatic[2]})`;
	const ruleOnly = /^(Exact|At least) rule \(no amount\)$/.exec(text);
	if (ruleOnly)
		return `${ruleOnly[1] === 'Exact' ? 'Règle exacte' : 'Règle minimale'} (sans valeur)`;
	const points = /^(Exact|At least) (.+) pt$/.exec(text);
	if (points) return `${points[1] === 'Exact' ? 'Exactement' : 'Au moins'} ${points[2]} pt`;
	return text in strings.en ? translate(locale, text as LocalizationKey) : text;
}
export function translateUiText(root: HTMLElement, englishText: string): string {
	return translate(
		normalizeEditorLocale(root.dataset.editorLocale),
		englishText as LocalizationKey,
	);
}
export function formatWordCount(locale: EditorLocale, count: number): string {
	return translate(locale, 'status.words').replace(
		'{count}',
		new Intl.NumberFormat(locale).format(count),
	);
}
export function formatPageStatus(locale: EditorLocale, current: number, total: number): string {
	return translate(locale, 'status.page')
		.replace('{current}', String(current))
		.replace('{total}', String(total));
}

/** Translate literal UI labels in-place while retaining action data and current control state. */
export function localizeElement(root: HTMLElement, locale: EditorLocale): void {
	for (const element of [root, ...root.querySelectorAll<HTMLElement>('*')]) {
		if (element.classList.contains('ribbon-group') && element.dataset.label) {
			const label = element.dataset.label;
			const translated = translate(locale, label as LocalizationKey);
			element.dataset.caption = translated;
			element.setAttribute(
				'aria-label',
				locale === 'fr' ? `Commandes : ${translated}` : `${translated} controls`,
			);
		}
		for (const attribute of ['aria-label', 'title', 'placeholder']) {
			const property = `locale${attribute.replace(/[^a-z]/gi, '')}`;
			const value = element.dataset[property] ?? element.getAttribute(attribute);
			if (value && value in strings.en) {
				element.dataset[property] = value;
				element.setAttribute(attribute, translate(locale, value as LocalizationKey));
			}
		}
		for (const node of element.childNodes) {
			if (node.nodeType !== Node.TEXT_NODE || !node.textContent) continue;
			const text = node.textContent.trim();
			const holder = element as HTMLElement & { _localeText?: Map<Text, string> };
			holder._localeText ??= new Map();
			const original = holder._localeText.get(node as Text) ?? text;
			if (
				!(original in strings.en) &&
				!/^(?:\d+(?:\.\d+)? lines|Automatic |Exact |At least )/.test(original)
			)
				continue;
			holder._localeText.set(node as Text, original);
			node.textContent = translateDynamic(locale, original);
		}
	}
}

export function findLocalizedControl<T extends HTMLElement>(
	root: ParentNode,
	label: string,
): T | null {
	return root.querySelector<T>(`[aria-label="${label}"], [data-localearialabel="${label}"]`);
}
