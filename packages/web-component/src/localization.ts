/** Display-language strings for the shared editor UI. This locale never changes DOCX language marks. */
import { EDITOR_LOCALES, strings, type LocalizationKey } from './localization-strings';

export type EditorLocale = (typeof EDITOR_LOCALES)[number];
/** A locale code a host may pass: a canonical code (autocompleted) or any BCP 47 tag such as `de-DE`. */
export type EditorLocaleInput = EditorLocale | (string & {});
export { EDITOR_LOCALES };
export type { LocalizationKey };

/**
 * Maps any BCP 47 tag to a supported display locale, falling back to English. Region variants
 * resolve to their language (`fr-CA`, `de-DE`, `es-MX`); `zh`, `zh-CN`, `zh-Hans` and `zh-SG` use
 * Simplified Chinese, while Traditional tags (`zh-TW`, `zh-HK`, `zh-Hant`) are unsupported.
 */
export function normalizeEditorLocale(value: string | null | undefined): EditorLocale {
	const parts = (value ?? '').trim().toLowerCase().split(/[-_]/);
	switch (parts[0]) {
		case 'fr':
		case 'de':
		case 'es':
			return parts[0];
		case 'zh':
			return parts.some((part) => ['hant', 'tw', 'hk', 'mo'].includes(part)) ? 'en' : 'zh-CN';
		default:
			return 'en';
	}
}
export function translate(locale: EditorLocale, key: LocalizationKey): string {
	return strings[locale][key] ?? strings.en[key] ?? key;
}
/** Translates a `{name}` template key and fills its placeholders. */
export function translateTemplate(
	locale: EditorLocale,
	key: LocalizationKey,
	values: Record<string, string | number>,
): string {
	return translate(locale, key).replace(/\{(\w+)\}/g, (match, name: string) =>
		name in values ? String(values[name]) : match,
	);
}
function translateDynamic(locale: EditorLocale, text: string): string {
	if (locale === 'en') return text;
	const line = /^(\d+(?:\.\d+)?) lines$/.exec(text);
	if (line)
		return translateTemplate(locale, Number(line[1]) === 1 ? 'dyn.line' : 'dyn.lines', {
			n: line[1],
		});
	const automatic = /^Automatic (.+) lines \((.+)\)$/.exec(text);
	if (automatic)
		return translateTemplate(locale, 'dyn.automatic', { n: automatic[1], ratio: automatic[2] });
	const ruleOnly = /^(Exact|At least) rule \(no amount\)$/.exec(text);
	if (ruleOnly)
		return translate(locale, ruleOnly[1] === 'Exact' ? 'dyn.exactRule' : 'dyn.atLeastRule');
	const points = /^(Exact|At least) (.+) pt$/.exec(text);
	if (points)
		return translateTemplate(locale, points[1] === 'Exact' ? 'dyn.exact' : 'dyn.atLeast', {
			n: points[2],
		});
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

/** A language's name in the display locale (`Intl.DisplayNames`), or the English fallback. */
function languageName(locale: EditorLocale, tag: string, english: string): string {
	if (locale === 'en') return english;
	try {
		const names = new Intl.DisplayNames([locale], {
			type: 'language',
			languageDisplay: 'standard',
		});
		return names.of(tag) ?? english;
	} catch {
		return english;
	}
}

/** The display locale of the editor chrome that contains `element`. */
export function localeOf(element: Element): EditorLocale {
	return normalizeEditorLocale(
		element.closest<HTMLElement>('[data-editor-locale]')?.dataset.editorLocale,
	);
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
				translateTemplate(locale, 'ribbon.groupControls', { group: translated }),
			);
		}
		if (element.dataset.tabKey) {
			element.textContent = translate(locale, element.dataset.tabKey as LocalizationKey);
			continue;
		}
		if (element instanceof HTMLOptionElement && element.dataset.languageTag) {
			element.dataset.englishName ??= element.textContent ?? '';
			element.textContent = languageName(
				locale,
				element.dataset.languageTag,
				element.dataset.englishName,
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
