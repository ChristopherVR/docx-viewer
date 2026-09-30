export const enTableProperties = {
	'Table properties': 'Table properties',
	Properties: 'Properties',
	Row: 'Row',
	'Specify height': 'Specify height',
	'Height (inches)': 'Height (inches)',
	'Row height is': 'Row height is',
	'At least': 'At least',
	Exactly: 'Exactly',
	'Allow row to break across pages': 'Allow row to break across pages',
	'Repeat as header row at the top of each page': 'Repeat as header row at the top of each page',
	'Default cell margins (inches)': 'Default cell margins (inches)',
	'Row settings apply to selected rows. Cell margins are table defaults; individual cell overrides are kept.':
		'Row settings apply to selected rows. Cell margins are table defaults; individual cell overrides are kept.',
	'Enter a measurement within the displayed range.':
		'Enter a measurement within the displayed range.',
} as const;
type Keys = keyof typeof enTableProperties;
const fromValues = (values: string[]): Record<Keys, string> =>
	Object.fromEntries(
		Object.keys(enTableProperties).map((key, index) => [key, values[index]!]),
	) as Record<Keys, string>;
export const frTableProperties = fromValues([
	'Propriétés du tableau',
	'Propriétés',
	'Ligne',
	'Spécifier la hauteur',
	'Hauteur (pouces)',
	'Hauteur de ligne',
	'Au moins',
	'Exactement',
	'Autoriser le fractionnement des lignes sur plusieurs pages',
	'Répéter en tant que ligne d’en-tête en haut de chaque page',
	'Marges de cellule par défaut (pouces)',
	'Les paramètres de ligne s’appliquent aux lignes sélectionnées. Les marges de cellule sont les valeurs par défaut du tableau ; les réglages individuels sont conservés.',
	'Saisissez une mesure dans la plage affichée.',
]);
export const deTableProperties = fromValues([
	'Tabelleneigenschaften',
	'Eigenschaften',
	'Zeile',
	'Höhe angeben',
	'Höhe (Zoll)',
	'Zeilenhöhe',
	'Mindestens',
	'Genau',
	'Zeilenwechsel auf Seiten zulassen',
	'Als Überschrift auf jeder Seite wiederholen',
	'Standardzellränder (Zoll)',
	'Die Zeileneinstellungen gelten für ausgewählte Zeilen. Zellränder sind Tabellenstandardwerte; individuelle Zelleneinstellungen bleiben erhalten.',
	'Geben Sie einen Wert im angezeigten Bereich ein.',
]);
export const esTableProperties = fromValues([
	'Propiedades de tabla',
	'Propiedades',
	'Fila',
	'Especificar alto',
	'Alto (pulgadas)',
	'Alto de fila',
	'Mínimo',
	'Exacto',
	'Permitir dividir filas entre páginas',
	'Repetir como fila de encabezado en cada página',
	'Márgenes de celda predeterminados (pulgadas)',
	'Los ajustes de fila se aplican a las filas seleccionadas. Los márgenes son los predeterminados de la tabla; se conservan los ajustes individuales de celda.',
	'Introduzca una medida dentro del intervalo mostrado.',
]);
export const zhTableProperties = fromValues([
	'表格属性',
	'属性',
	'行',
	'指定高度',
	'高度（英寸）',
	'行高值',
	'最小值',
	'固定值',
	'允许跨页断行',
	'在各页顶端以标题行形式重复出现',
	'默认单元格边距（英寸）',
	'行设置应用于所选行。单元格边距为表格默认值；保留单个单元格的设置。',
	'请输入显示范围内的尺寸。',
]);
