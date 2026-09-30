const labels = [
	'Ligatures',
	'All',
	'Standard',
	'Contextual',
	'Historical',
	'Discretional',
	'Standard Contextual',
	'Standard Historical',
	'Contextual Historical',
	'Standard Discretional',
	'Contextual Discretional',
	'Historical Discretional',
	'Standard Contextual Historical',
	'Standard Contextual Discretional',
	'Standard Historical Discretional',
	'Contextual Historical Discretional',
] as const;
type Key = (typeof labels)[number];
function translated(words: [string, string, string, string, string, string]): Record<Key, string> {
	const basic = new Map(labels.slice(0, 6).map((key, index) => [key, words[index]!]));
	return Object.fromEntries(
		labels.map((label) => [
			label,
			label
				.split(' ')
				.map((word) => basic.get(word as Key)!)
				.join(' '),
		]),
	) as Record<Key, string>;
}
export const enLigatures = translated([
	'Ligatures',
	'All',
	'Standard',
	'Contextual',
	'Historical',
	'Discretional',
]);
export const frLigatures = translated([
	'Ligatures',
	'Toutes',
	'Standard',
	'Contextuelles',
	'Historiques',
	'Discrétionnaires',
]);
export const deLigatures = translated([
	'Ligaturen',
	'Alle',
	'Standard',
	'Kontextuell',
	'Historisch',
	'Optional',
]);
export const esLigatures = translated([
	'Ligaduras',
	'Todas',
	'Estándar',
	'Contextuales',
	'Históricas',
	'Discrecionales',
]);
export const zhLigatures = translated(['连字', '全部', '标准', '上下文', '历史', '任意']);
