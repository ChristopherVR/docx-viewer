export const enNavigation = {
	'Navigation pane': 'Navigation pane',
	Navigation: 'Navigation',
	Headings: 'Headings',
	'Close navigation pane': 'Close navigation pane',
	'Expand heading': 'Expand heading',
	'Collapse heading': 'Collapse heading',
	'Apply a heading style to see document headings here.':
		'Apply a heading style to see document headings here.',
} as const;
type Keys = keyof typeof enNavigation;
const fromValues = (values: string[]): Record<Keys, string> =>
	Object.fromEntries(
		Object.keys(enNavigation).map((key, index) => [key, values[index]!]),
	) as Record<Keys, string>;
export const frNavigation = fromValues([
	'Volet de navigation',
	'Navigation',
	'Titres',
	'Fermer le volet de navigation',
	'Développer le titre',
	'Réduire le titre',
	'Appliquez un style de titre pour afficher les titres du document ici.',
]);
export const deNavigation = fromValues([
	'Navigationsbereich',
	'Navigation',
	'Überschriften',
	'Navigationsbereich schließen',
	'Überschrift erweitern',
	'Überschrift reduzieren',
	'Wenden Sie eine Überschriftenformatvorlage an, um hier Dokumentüberschriften anzuzeigen.',
]);
export const esNavigation = fromValues([
	'Panel de navegación',
	'Navegación',
	'Títulos',
	'Cerrar el panel de navegación',
	'Expandir título',
	'Contraer título',
	'Aplique un estilo de título para ver aquí los títulos del documento.',
]);
export const zhNavigation = fromValues([
	'导航窗格',
	'导航',
	'标题',
	'关闭导航窗格',
	'展开标题',
	'折叠标题',
	'应用标题样式以在此处查看文档标题。',
]);
