export const enDropCap = {
	'Drop Cap Options': 'Drop Cap Options',
	'Drop Cap Options…': 'Drop Cap Options…',
	'Lines to drop': 'Lines to drop',
	'Enter a font, 1 to 10 lines and a distance from 0 to 22 inches.':
		'Enter a font, 1 to 10 lines and a distance from 0 to 22 inches.',
} as const;
type Keys = keyof typeof enDropCap;
const fromValues = (values: string[]): Record<Keys, string> =>
	Object.fromEntries(Object.keys(enDropCap).map((key, index) => [key, values[index]!])) as Record<
		Keys,
		string
	>;
export const frDropCap = fromValues([
	'Options de lettrine',
	'Options de lettrine…',
	'Hauteur en lignes',
	'Saisissez une police, 1 à 10 lignes et une distance de 0 à 22 pouces.',
]);
export const deDropCap = fromValues([
	'Initialoptionen',
	'Initialoptionen…',
	'Initialhöhe in Zeilen',
	'Geben Sie eine Schriftart, 1 bis 10 Zeilen und einen Abstand von 0 bis 22 Zoll ein.',
]);
export const esDropCap = fromValues([
	'Opciones de letra capital',
	'Opciones de letra capital…',
	'Líneas que ocupa',
	'Escriba una fuente, de 1 a 10 líneas y una distancia de 0 a 22 pulgadas.',
]);
export const zhDropCap = fromValues([
	'首字下沉选项',
	'首字下沉选项…',
	'下沉行数',
	'请输入字体、1 至 10 行以及 0 至 22 英寸的距离。',
]);
