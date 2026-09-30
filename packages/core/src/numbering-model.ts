// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { StNumberFormat } from './generated/wml-simple-types.js';
import type { SignedTwips, Twips } from './units.js';
/** A single `w:lvl` definition, fully resolved from an abstractNum (and any lvlOverride). */
export interface NumberingLevelDefinition {
	/** 0-based `w:ilvl`. */
	level: number;
	start: number;
	/** `w:numFmt/@w:val` (`ST_NumberFormat`), such as `decimal`, `bullet`, `lowerRoman`. Invalid values fall back to `decimal` with a parse warning. */
	numFmt: StNumberFormat;
	/** Raw `w:lvlText/@w:val`, with `%1`..`%9` placeholders (1-based ancestor levels) or a bullet glyph. */
	lvlText: string;
	lvlJc?: 'left' | 'center' | 'right';
	indentLeftTwips?: SignedTwips;
	hangingTwips?: Twips;
	firstLineTwips?: Twips;
	/** `w:isLgl`: render every placeholder in this level's marker as Decimal Number regardless of format. */
	isLgl?: boolean;
	/** `w:lvlRestart/@w:val`: one-based higher-level restart trigger; 0 means never. Omitted defaults to the previous level. */
	lvlRestart?: number;
	/** Marker-to-text separator from `w:suff` (defaults to `tab`). */
	suffix?: 'tab' | 'space' | 'none';
}
export interface AbstractNumDefinition {
	id: string;
	levels: Record<number, NumberingLevelDefinition>;
}
export interface NumLevelOverride {
	startOverride?: number;
	lvl?: NumberingLevelDefinition;
}
export interface NumDefinition {
	id: string;
	abstractNumId: string;
	levelOverrides?: Record<number, NumLevelOverride>;
}
/** Parsed source `word/numbering.xml`. Existing entries are read-only; only additive entries can be saved. */
export interface NumberingCatalog {
	abstractNums: Record<string, AbstractNumDefinition>;
	nums: Record<string, NumDefinition>;
	warnings: string[];
}
/** A computed, render-only list marker for one paragraph. Never written back to the model. */
export interface ParagraphListLabel {
	numId: string;
	level: number;
	/** Marker text with `%n` placeholders already substituted (no trailing suffix). */
	text: string;
	suffix: 'tab' | 'space' | 'none';
	/** Alignment of the marker around the paragraph's first-line position. */
	alignment?: 'left' | 'center' | 'right';
	indentLeftTwips?: SignedTwips;
	hangingTwips?: Twips;
	firstLineTwips?: Twips;
}
