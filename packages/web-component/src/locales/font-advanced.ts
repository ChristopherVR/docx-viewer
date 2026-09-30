export const enFontAdvanced = {
	Advanced: 'Advanced',
	'Scale (%)': 'Scale (%)',
	Raised: 'Raised',
	Lowered: 'Lowered',
	'Position by (points)': 'Position by (points)',
	'Kerning for fonts': 'Kerning for fonts',
	'Points and above': 'Points and above',
	'Review advanced formatting in Print Layout.': 'Review advanced formatting in Print Layout.',
} as const;
type Keys = keyof typeof enFontAdvanced;
const values = (items: string[]): Record<Keys, string> =>
	Object.fromEntries(
		Object.keys(enFontAdvanced).map((key, index) => [key, items[index]!]),
	) as Record<Keys, string>;
export const frFontAdvanced = values([
	'Paramètres avancés',
	'Échelle (%)',
	'Décalage vers le haut',
	'Décalage vers le bas',
	'Décalage (points)',
	'Crénage',
	'Points et plus',
	'Vérifiez la mise en forme avancée en mode Mise en page.',
]);
export const deFontAdvanced = values([
	'Erweitert',
	'Skalieren (%)',
	'Höhergestellt',
	'Tiefergestellt',
	'Versatz (Punkte)',
	'Unterschneidung für Schriftarten',
	'Punkte und mehr',
	'Prüfen Sie die erweiterte Formatierung im Drucklayout.',
]);
export const esFontAdvanced = values([
	'Avanzado',
	'Escala (%)',
	'Elevado',
	'Disminuido',
	'Posición en (puntos)',
	'Interletraje para fuentes',
	'Puntos o más',
	'Revise el formato avanzado en Diseño de impresión.',
]);
export const zhFontAdvanced = values([
	'高级',
	'缩放 (%)',
	'提升',
	'降低',
	'位置值（磅）',
	'为字体调整字距',
	'磅及以上',
	'请在打印布局中检查高级格式。',
]);
