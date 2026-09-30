export const enColumns = {
	Column: 'Column',
	'More Columns…': 'More Columns…',
	'Equal column width': 'Equal column width',
	'Line between': 'Line between',
	'Number of columns': 'Number of columns',
	'Column width (inches)': 'Column width (inches)',
	'Column spacing (inches)': 'Column spacing (inches)',
	'Unequal column widths are preserved. Select Equal column width to replace them with equal columns.':
		'Unequal column widths are preserved. Select Equal column width to replace them with equal columns.',
	'Enter whole columns and spacing that leaves at least 0.5 inch for each column.':
		'Enter whole columns and spacing that leaves at least 0.5 inch for each column.',
} as const;
type Keys = keyof typeof enColumns;
const fromValues = (values: string[]): Record<Keys, string> =>
	Object.fromEntries(Object.keys(enColumns).map((key, index) => [key, values[index]!])) as Record<
		Keys,
		string
	>;
export const frColumns = fromValues([
	'Colonne',
	'Autres colonnes…',
	'Largeur de colonne identique',
	'Ligne séparatrice',
	'Nombre de colonnes',
	'Largeur de colonne (pouces)',
	'Espacement des colonnes (pouces)',
	'Les largeurs inégales sont conservées. Sélectionnez Largeur de colonne identique pour les remplacer par des colonnes égales.',
	'Saisissez un nombre entier de colonnes et un espacement laissant au moins 0,5 pouce par colonne.',
]);
export const deColumns = fromValues([
	'Spalte',
	'Weitere Spalten…',
	'Gleiche Spaltenbreite',
	'Zwischenlinie',
	'Spaltenanzahl',
	'Spaltenbreite (Zoll)',
	'Spaltenabstand (Zoll)',
	'Ungleiche Spaltenbreiten bleiben erhalten. Wählen Sie Gleiche Spaltenbreite, um sie durch gleiche Spalten zu ersetzen.',
	'Geben Sie eine ganze Spaltenanzahl und einen Abstand ein, der jeder Spalte mindestens 0,5 Zoll lässt.',
]);
export const esColumns = fromValues([
	'Columna',
	'Más columnas…',
	'Columnas de igual ancho',
	'Línea entre columnas',
	'Número de columnas',
	'Ancho de columna (pulgadas)',
	'Espaciado de columnas (pulgadas)',
	'Se conservan los anchos desiguales. Seleccione Columnas de igual ancho para sustituirlas por columnas iguales.',
	'Escriba columnas enteras y un espaciado que deje al menos 0,5 pulgadas para cada columna.',
]);
export const zhColumns = fromValues([
	'栏',
	'更多分栏…',
	'栏宽相等',
	'分隔线',
	'栏数',
	'栏宽（英寸）',
	'栏间距（英寸）',
	'将保留不等栏宽。选择“栏宽相等”可将其替换为等宽分栏。',
	'请输入整数栏数和间距，确保每栏至少有 0.5 英寸宽。',
]);
