// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
/** A single `w:lvl` definition, fully resolved from an abstractNum (and any lvlOverride). */
export interface NumberingLevelDefinition {
	/** 0-based `w:ilvl`. */
	level: number;
	start: number;
	/** Raw `w:numFmt/@w:val`, such as `decimal`, `bullet`, `lowerRoman`. Unknown formats are kept literally. */
	numFmt: string;
	/** Raw `w:lvlText/@w:val`, with `%1`..`%9` placeholders (1-based ancestor levels) or a bullet glyph. */
	lvlText: string;
	lvlJc?: 'left' | 'center' | 'right';
	indentLeftTwips?: number;
	hangingTwips?: number;
	firstLineTwips?: number;
	/** `w:isLgl`: render every placeholder in this level's marker as Decimal Number regardless of format. */
	isLgl?: boolean;
	/** `w:lvlRestart/@w:val`: the shallowest level (0-based) whose increment restarts this level's counter. */
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
	indentLeftTwips?: number;
	hangingTwips?: number;
	firstLineTwips?: number;
}
