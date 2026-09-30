// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { StHighlightColor } from './generated/wml-simple-types.js';
import type { HyperlinkInfo, InlineImage } from './inline-model.js';
import type {
	ThemeColorReference,
	ThemeFontRole,
	ThemeFontScript,
	WordUnderlineStyle,
} from './theme-model.js';
import type { HalfPoints, SignedTwips } from './units.js';

/** A tracked-change revision recorded on a run or paragraph mark. */
export interface Revision {
	kind: 'insert' | 'delete' | 'moveFrom' | 'moveTo' | 'formatChange' | 'paragraphChange';
	/**
	 * For `moveFrom`/`moveTo`: the move this text belongs to. Both sides share `name` (from
	 * `w:moveFromRangeStart`/`w:moveToRangeStart`); `rangeId` is that range marker's `w:id`.
	 */
	move?: { name: string; rangeId?: string };
	author: string;
	date?: string;
	/** Source `w:id`; not guaranteed unique outside the paragraph it was parsed from. */
	id: string;
}

export interface TextRun {
	text: string;
	bold?: boolean;
	italic?: boolean;
	underline?: boolean;
	strike?: boolean;
	/** Word named highlight color, such as `yellow` or `lightGray`. */
	highlight?: StHighlightColor;
	verticalAlign?: 'baseline' | 'superscript' | 'subscript';
	/** Direct Word run language tag (`w:lang/@w:val`), without automatic detection. */
	language?: string;
	/** Direct East Asian script language tag (`w:lang/@w:eastAsia`). */
	eastAsiaLanguage?: string;
	/** Direct complex-script language tag (`w:lang/@w:bidi`). */
	bidiLanguage?: string;
	/** Explicit run-level bidirectional override; `false` means direct off, undefined inherits. */
	rtl?: boolean;
	/** Font size in points (whole or half points, e.g. 10.5); `HalfPoints` is the on-disk `w:sz`. */
	fontSize?: number;
	fontFamily?: string;
	/** Direct RGB color, e.g. #28665E. May coexist with `colorTheme` as Word's stored fallback. */
	color?: string;
	/**
	 * A run that is itself a page or column break marker (`w:br` type page/column) instead of
	 * visible text. `text` is empty on these runs; editors render a distinct visible marker.
	 */
	break?: 'page' | 'column';
	/**
	 * A run that is itself a footnote/endnote reference mark (`w:footnoteReference` /
	 * `w:endnoteReference`) instead of visible text. `text` is empty on these runs; the numeric
	 * mark is derived from document order, not stored here.
	 */
	noteReference?: { kind: 'footnote' | 'endnote'; id: string };
	/** The automatic number mark (`w:footnoteRef`/`w:endnoteRef`) that starts a note's own text. */
	noteMark?: 'footnote' | 'endnote';
	/**
	 * Present on runs holding a field's displayed result (`w:fldSimple`, or text between a complex
	 * field's `separate` and `end`). Display metadata only: fields are not recalculated on save and
	 * paragraphs containing them stay protected from edits.
	 */
	field?: { instr: string; simple?: boolean };
	/** A complex field's `w:fldChar` marker run (begin, separate or end); `text` is empty. */
	fieldChar?: 'begin' | 'separate' | 'end';
	/** A complex field's instruction text run (`w:instrText`), e.g. ` TOC \o "1-3" `; `text` is empty. */
	fieldCode?: string;
	/** Tracked-change metadata for this run; absent means the run has no pending revision. */
	revision?: Revision;
	/** IDs of comments whose range covers this run. */
	commentIds?: string[];
	/** Direct `w:color/@w:themeColor` (+ themeTint/themeShade); resolution happens in a separate layer. */
	colorTheme?: ThemeColorReference;
	/** Character style reference (`w:rStyle/@w:val`); preserved and editable, not flattened. */
	style?: string;
	caps?: boolean;
	smallCaps?: boolean;
	/** `w:dstrike`; kept distinct from the single-line `strike` toggle. */
	doubleStrike?: boolean;
	/** Hidden text (`w:vanish`); the editor renders it dimmed rather than removing it. */
	vanish?: boolean;
	/** Non-single underline style, e.g. `double`/`wave`; `underline` stays the simple on/off toggle. */
	underlineStyle?: WordUnderlineStyle;
	underlineColor?: string;
	/** `w:spacing/@w:val` character spacing, in twips (positive expands, negative condenses). */
	characterSpacingTwips?: SignedTwips;
	/** `w:w` horizontal text scaling, in whole percent; 100 explicitly cancels inherited scaling. */
	textScalePercent?: number;
	/** Office 2010 OpenType ligature selection; undefined inherits. */
	ligatures?: import('./ligatures.js').Ligatures;
	/** `w:kern` minimum font size for kerning, in half-points; zero explicitly disables it. */
	kerningHalfPoints?: HalfPoints;
	/** `w:position` baseline displacement, in signed half-points; positive raises, negative lowers. */
	positionHalfPoints?: HalfPoints;
	/** Direct `w:shd/@w:fill` run shading. */
	shadingFill?: string;
	/** Direct `w:shd` theme fill; kept alongside `shadingFill` without flattening. */
	shadingThemeFill?: ThemeColorReference;
	/** Direct `w:rFonts` theme font references, per script; resolved via the document theme. */
	fontTheme?: Partial<Record<ThemeFontScript, ThemeFontRole>>;
	/** Present when this run is an inline picture instead of text; `text` is empty. */
	image?: InlineImage;
	/** Hyperlink target for this run, from `w:hyperlink` (or a simple `HYPERLINK` field). */
	link?: HyperlinkInfo;
}
