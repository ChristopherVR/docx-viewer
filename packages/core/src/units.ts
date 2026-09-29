// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Branded numeric unit types. Brands are compile-time only: a `Twips` is still a plain number at
// runtime and stays assignable to `number` (so arithmetic and comparisons keep working), but a
// plain `number`, or a value in another unit, is not assignable to a branded field.
// Brand values are created only in this module: parsers use the constructors, and code that
// computes values (UI, layout math) goes through the explicit rounding constructors below.
declare const unitBrand: unique symbol;
declare const nonNegativeBrand: unique symbol;
type Branded<Name extends string> = number & { readonly [unitBrand]: Name };

/** Twentieths of a point (1/1440 inch), possibly negative: `ST_SignedTwipsMeasure`. */
export type SignedTwips = Branded<'twips'>;
/**
 * Non-negative twentieths of a point (`ST_TwipsMeasure`): WordprocessingML's native length unit.
 * A `Twips` is assignable to `SignedTwips`, never the other way round.
 */
export type Twips = SignedTwips & { readonly [nonNegativeBrand]: true };
/** Half-points (`w:sz`): font sizes are stored as twice their point size. */
export type HalfPoints = Branded<'halfPoints'>;
/** Eighth-points (`w:sz` on borders): border widths. */
export type EighthPoints = Branded<'eighthPoints'>;
/** English Metric Units (1/914400 inch): DrawingML lengths. */
export type Emu = Branded<'emu'>;

function integer<T extends number>(unit: string, value: number, allowNegative = true): T {
	if (!Number.isSafeInteger(value)) throw new RangeError(`${unit} must be a safe integer`);
	if (!allowNegative && value < 0) throw new RangeError(`${unit} must not be negative`);
	return value as T;
}
function finite(unit: string, value: number): number {
	if (!Number.isFinite(value)) throw new RangeError(`${unit} must be finite`);
	return value;
}
/** Round half away from zero (so -1.5 -> -2), never returning -0. */
const roundAway = (unit: string, value: number): number =>
	Math.sign(finite(unit, value)) * Math.round(Math.abs(value)) || 0;

// Strict constructors: the value must already be a whole number in the unit (parsers, tests).
/** A non-negative whole number of twips; throws `RangeError` otherwise. */
export const twips = (value: number): Twips => integer<Twips>('twips', value, false);
/** A whole number of twips that may be negative; throws `RangeError` unless a safe integer. */
export const signedTwips = (value: number): SignedTwips => integer<SignedTwips>('twips', value);
export const halfPoints = (value: number): HalfPoints => integer<HalfPoints>('halfPoints', value);
export const eighthPoints = (value: number): EighthPoints =>
	integer<EighthPoints>('eighthPoints', value);
export const emu = (value: number): Emu => integer<Emu>('emu', value);

// Rounding constructors: for computed values (drag math, zoom, pixel conversions). They round to
// the nearest whole unit (halves away from zero) and reject only non-finite input. `roundTwips`
// clamps negative results to 0 because `Twips` is non-negative.
export const roundTwips = (value: number): Twips =>
	integer<Twips>('twips', Math.max(0, roundAway('twips', value)));
export const roundSignedTwips = (value: number): SignedTwips =>
	integer<SignedTwips>('twips', roundAway('twips', value));
export const roundHalfPoints = (value: number): HalfPoints =>
	integer<HalfPoints>('halfPoints', roundAway('halfPoints', value));
export const roundEighthPoints = (value: number): EighthPoints =>
	integer<EighthPoints>('eighthPoints', Math.max(0, roundAway('eighthPoints', value)));
export const roundEmu = (value: number): Emu => integer<Emu>('emu', roundAway('emu', value));

// Cross-unit constructors from real-world units; all round to the nearest whole target unit.
/** Points (1/72 inch) to twips; negative input clamps to 0. */
export const twipsFromPoints = (points: number): Twips => roundTwips(points * 20);
export const signedTwipsFromPoints = (points: number): SignedTwips => roundSignedTwips(points * 20);
/** CSS pixels at 96 dpi (15 twips per pixel) to twips; negative input clamps to 0. */
export const twipsFromPixels = (pixels: number): Twips => roundTwips(pixels * 15);
export const signedTwipsFromPixels = (pixels: number): SignedTwips => roundSignedTwips(pixels * 15);
export const twipsFromInches = (inches: number): Twips => roundTwips(inches * 1440);
/** A font size in points to half-points (Word stores `w:sz` as twice the point size). */
export const halfPointsFromPoints = (points: number): HalfPoints => roundHalfPoints(points * 2);
/** A border width in points to eighth-points; negative input clamps to 0. */
export const eighthPointsFromPoints = (points: number): EighthPoints =>
	roundEighthPoints(points * 8);
export const emuFromPixels = (pixels: number): Emu => roundEmu(pixels * 9525);

// Conversions out of a unit. Results are plain (possibly fractional) numbers in the named unit.
export const twipsToPoints = (value: SignedTwips): number => value / 20;
export const twipsToPixels = (value: SignedTwips): number => value / 15;
export const twipsToInches = (value: SignedTwips): number => value / 1440;
export const halfPointsToPoints = (value: HalfPoints): number => value / 2;
export const eighthPointsToPoints = (value: EighthPoints): number => value / 8;
export const emuToPixels = (value: Emu): number => value / 9525;
