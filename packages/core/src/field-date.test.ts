import { describe, expect, it } from 'vitest';
import { dateFieldResult, datePicture, formatWordDate } from './index.js';

const date = new Date(2026, 8, 7, 14, 5, 9); // Monday 7 September 2026, 2:05:09 pm

describe('Word date and time fields', () => {
	it('reads the \\@ picture switch with or without quotes', () => {
		expect(datePicture(' DATE \\@ "MMMM d, yyyy" ')).toBe('MMMM d, yyyy');
		expect(datePicture(' TIME \\@ HH:mm ')).toBe('HH:mm');
		expect(datePicture(' DATE ')).toBeUndefined();
	});

	it('formats Word date-time pictures', () => {
		expect(formatWordDate(date, 'dddd, MMMM d, yyyy')).toBe('Monday, September 7, 2026');
		expect(formatWordDate(date, 'dd/MM/yy')).toBe('07/09/26');
		expect(formatWordDate(date, 'h:mm am/pm')).toBe('2:05 pm');
		expect(formatWordDate(date, 'HH:mm:ss')).toBe('14:05:09');
		expect(formatWordDate(date, "d 'of' MMM")).toBe('7 of Sep');
	});

	it('uses Word defaults for DATE and TIME without a picture', () => {
		expect(dateFieldResult('DATE', ' DATE ', date)).toBe('9/7/2026');
		expect(dateFieldResult('TIME', ' TIME ', date)).toBe('2:05 pm');
		expect(dateFieldResult('DATE', ' DATE \\@ "d MMMM yyyy" ', date, 'fr-FR')).toBe(
			'7 septembre 2026',
		);
	});
});
