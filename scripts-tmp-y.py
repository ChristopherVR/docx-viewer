import os
os.chdir('packages/web-component/src')

def sub(p, old, new, count=1):
    s = open(p, encoding='utf8').read()
    assert old in s, (p, old[:70])
    open(p, 'w', encoding='utf8').write(s.replace(old, new, count))

sub('style-gallery.ts', "/** Applies the style's resolved run formatting", """const HIDDEN_CHARACTER =
	/(char$|reference$|default paragraph font|hyperlink|page number|line number|placeholder)/i;

/** The character styles Word offers in the gallery (Emphasis, Strong, ...), not linked or technical ones. */
export function recommendedCharacterStyles(all: readonly GalleryStyle[]): GalleryStyle[] {
	return all.filter((style) => !HIDDEN_CHARACTER.test(style.name));
}

/** Applies the style's resolved run formatting""")

# StylesData: quick lists for the popover, full lists for the pane
sub('styles-popover.ts', "\tparagraph: GalleryStyle[];\n\tcharacter: GalleryStyle[];", "\t/** Recommended styles, for the gallery and its expanded view. */\n\tparagraph: GalleryStyle[];\n\tcharacter: GalleryStyle[];\n\t/** Every style, for the Styles pane. */\n\tallParagraph: GalleryStyle[];\n\tallCharacter: GalleryStyle[];")

sub('styles-group.ts', "\treturn {\n\t\tparagraph: all,\n\t\tcharacter: Object.values(model.characterStyles?.styles ?? {})\n\t\t\t.filter((style) => style.type === 'character')\n\t\t\t.map((style) => ({ id: style.id, name: style.name || style.id })),", "\tconst characters = Object.values(model.characterStyles?.styles ?? {})\n\t\t.filter((style) => style.type === 'character')\n\t\t.map((style) => ({ id: style.id, name: style.name || style.id }));\n\treturn {\n\t\tparagraph: recommendedStyles(all),\n\t\tcharacter: recommendedCharacterStyles(characters),\n\t\tallParagraph: all,\n\t\tallCharacter: characters,")
sub('styles-group.ts', "import {\n\trecommendedStyles,", "import {\n\trecommendedCharacterStyles,\n\trecommendedStyles,")

s = open('styles-pane.ts', encoding='utf8').read()
s = s.replace("\t\t\t\tdata.paragraph,\n\t\t\t\tdata.character,\n\t\t\t\tnext,", "\t\t\t\tdata.allParagraph,\n\t\t\t\tdata.allCharacter,\n\t\t\t\tnext,")
s = s.replace("list(translate(locale, 'Paragraph styles'), data.paragraph, 'paragraph')", "list(translate(locale, 'Paragraph styles'), data.allParagraph, 'paragraph')")
s = s.replace("...(data.character.length", "...(data.allCharacter.length")
s = s.replace("...data.character],", "...data.allCharacter],")
open('styles-pane.ts', 'w', encoding='utf8').write(s)
