export const enLineNumbers = {
	'Add line numbering': 'Add line numbering',
	'Automatic distance from text': 'Automatic distance from text',
	'Start at': 'Start at',
	'Count by': 'Count by',
	'Distance from text (inches)': 'Distance from text (inches)',
	Numbering: 'Numbering',
	'Suppress line numbers': 'Suppress line numbers',
	'Suppress for Current Paragraph': 'Suppress for Current Paragraph',
	'Line Numbering Options…': 'Line Numbering Options…',
	'Enter whole numbers from 1 to 32767 and a distance from 0 to 22 inches.':
		'Enter whole numbers from 1 to 32767 and a distance from 0 to 22 inches.',
} as const;
type Keys = keyof typeof enLineNumbers;
const fromValues = (values: string[]): Record<Keys, string> =>
	Object.fromEntries(
		Object.keys(enLineNumbers).map((key, index) => [key, values[index]!]),
	) as Record<Keys, string>;
export const frLineNumbers = fromValues([
	'Ajouter la numérotation des lignes',
	'Distance automatique du texte',
	'Commencer à',
	'Incrément',
	'Distance du texte (pouces)',
	'Numérotation',
	'Supprimer les numéros de ligne',
	'Supprimer pour le paragraphe actif',
	'Options de numérotation des lignes…',
	'Saisissez des entiers de 1 à 32767 et une distance de 0 à 22 pouces.',
]);
export const deLineNumbers = fromValues([
	'Zeilennummern hinzufügen',
	'Automatischer Abstand zum Text',
	'Beginnen mit',
	'Zählintervall',
	'Abstand zum Text (Zoll)',
	'Nummerierung',
	'Zeilennummern unterdrücken',
	'Für aktuellen Absatz unterdrücken',
	'Zeilennummerierungsoptionen…',
	'Geben Sie ganze Zahlen von 1 bis 32767 und einen Abstand von 0 bis 22 Zoll ein.',
]);
export const esLineNumbers = fromValues([
	'Agregar números de línea',
	'Distancia automática del texto',
	'Iniciar en',
	'Contar por',
	'Distancia del texto (pulgadas)',
	'Numeración',
	'Suprimir números de línea',
	'Suprimir para el párrafo actual',
	'Opciones de numeración de líneas…',
	'Introduzca enteros de 1 a 32767 y una distancia de 0 a 22 pulgadas.',
]);
export const zhLineNumbers = fromValues([
	'添加行号',
	'自动设置距正文的距离',
	'起始编号',
	'编号间隔',
	'距正文（英寸）',
	'编号',
	'取消行号',
	'取消当前段落的行号',
	'行编号选项…',
	'请输入 1 到 32767 的整数及 0 到 22 英寸的距离。',
]);
