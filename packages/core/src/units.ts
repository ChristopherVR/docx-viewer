// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Branded numeric unit types. Brands are compile-time only: a `Twips` is still a plain number at
// runtime and remains assignable to `number`, so existing model fields migrate incrementally.
declare const unitBrand: unique symbol;
type Branded<Name extends string> = number & { readonly [unitBrand]: Name };

/** Twentieths of a point (1/1440 inch): WordprocessingML's native length unit. */
export type Twips = Branded<'twips'>;
/** Half-points (`w:sz`): font sizes are stored as twice their point size. */
export type HalfPoints = Branded<'halfPoints'>;
/** Eighth-points (`w:sz` on borders): border widths. */
export type EighthPoints = Branded<'eighthPoints'>;
/** English Metric Units (1/914400 inch): DrawingML lengths. */
export type Emu = Branded<'emu'>;

function integer<T extends number>(unit: string, value: number): T {
	if (!Number.isSafeInteger(value)) throw new RangeError(`${unit} must be a safe integer`);
	return value as T;
}
export const twips = (value: number): Twips => integer<Twips>('twips', value);
export const halfPoints = (value: number): HalfPoints => integer<HalfPoints>('halfPoints', value);
export const eighthPoints = (value: number): EighthPoints =>
	integer<EighthPoints>('eighthPoints', value);
export const emu = (value: number): Emu => integer<Emu>('emu', value);

export const twipsToPoints = (value: Twips): number => value / 20;
export const halfPointsToPoints = (value: HalfPoints): number => value / 2;
export const eighthPointsToPoints = (value: EighthPoints): number => value / 8;
export const emuToPixels = (value: Emu): number => value / 9525;
