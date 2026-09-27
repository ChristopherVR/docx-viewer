/** Display-language strings for the shared editor UI. This locale never changes DOCX language marks. */
export type EditorLocale = 'en' | 'fr';

import { strings } from './localization-strings';

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
