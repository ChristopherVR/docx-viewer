/** Highlight names from the WordprocessingML colour enumeration, as CSS colours. */
const NAMED_COLORS: Record<string, string> = {
	yellow: '#ffff00',
	green: '#00ff00',
	cyan: '#00ffff',
	magenta: '#ff00ff',
	blue: '#0000ff',
	red: '#ff0000',
	darkBlue: '#000080',
	darkCyan: '#008080',
	darkGreen: '#008000',
	darkMagenta: '#800080',
	darkRed: '#800000',
	darkYellow: '#808000',
	darkGray: '#808080',
	lightGray: '#c0c0c0',
	black: '#000000',
	white: '#ffffff',
};

/** CSS colour for a ribbon colour value: a hex string, a highlight name, or `none`. */
export function swatchColor(value: string): string {
	return value === 'none' ? 'transparent' : (NAMED_COLORS[value] ?? value);
}
