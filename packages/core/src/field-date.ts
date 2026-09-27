// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Word date/time field results: the `\@` date-time picture switch (e.g. `\@ "MMMM d, yyyy"`).

/** The `\@` picture of a field instruction, if any (quotes optional). */
export function datePicture(instr: string): string | undefined {
	const match = /\\@\s*(?:"([^"]*)"|(\S+))/.exec(instr);
	return match ? (match[1] ?? match[2]) : undefined;
}

const TOKENS = /yyyy|yy|MMMM|MMM|MM|M|dddd|ddd|dd|d|HH|H|hh|h|mm|ss|am\/pm|AM\/PM|'[^']*'/g;

/**
 * Formats `date` with a Word date-time picture. Month and day names follow `locale`; text in
 * single quotes is literal. Unrecognized characters are copied as-is, as Word does.
 */
export function formatWordDate(date: Date, picture: string, locale = 'en-US'): string {
	const name = (options: Intl.DateTimeFormatOptions) =>
		new Intl.DateTimeFormat(locale, options).format(date);
	const pad = (value: number) => String(value).padStart(2, '0');
	const hours12 = date.getHours() % 12 || 12;
	return picture.replace(TOKENS, (token) => {
		switch (token) {
			case 'yyyy':
				return String(date.getFullYear());
			case 'yy':
				return pad(date.getFullYear() % 100);
			case 'MMMM':
				return name({ month: 'long' });
			case 'MMM':
				return name({ month: 'short' });
			case 'MM':
				return pad(date.getMonth() + 1);
			case 'M':
				return String(date.getMonth() + 1);
			case 'dddd':
				return name({ weekday: 'long' });
			case 'ddd':
				return name({ weekday: 'short' });
			case 'dd':
				return pad(date.getDate());
			case 'd':
				return String(date.getDate());
			case 'HH':
				return pad(date.getHours());
			case 'H':
				return String(date.getHours());
			case 'hh':
				return pad(hours12);
			case 'h':
				return String(hours12);
			case 'mm':
				return pad(date.getMinutes());
			case 'ss':
				return pad(date.getSeconds());
			case 'am/pm':
				return date.getHours() < 12 ? 'am' : 'pm';
			case 'AM/PM':
				return date.getHours() < 12 ? 'AM' : 'PM';
			default:
				return token.slice(1, -1);
		}
	});
}

/** DATE and TIME results as Word shows them when it updates the field (default pictures included). */
export function dateFieldResult(
	name: 'DATE' | 'TIME',
	instr: string,
	now: Date,
	locale?: string,
): string {
	const picture = datePicture(instr) ?? (name === 'DATE' ? 'M/d/yyyy' : 'h:mm am/pm');
	return formatWordDate(now, picture, locale);
}
