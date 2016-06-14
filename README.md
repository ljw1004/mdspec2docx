mdspec2docx
===========

This utility converts markdown to docx. There are other utilities which do the same, but this is tailored specifically to *language specifications* that are written in markdown...

* It uses "template.docx" as the template, where you can add frontmatter and configure styles
* It has excellent syntax highlighting for C# and VB. It also has syntax highlighting for Antlr.
* It automatically adds numerous cross-references to the generated docx
* It keys off a master Antlr .g4 grammar file and ensures that the snippets of ```antlr code inside the specification markdown match that whole grammar file exactly
* It generates a hyperlinked version of the grammar


# Conventions

The markdown generates a docx that uses
* Paragraph styles: `TOC1, TOC2, Heading1, Heading2, Heading3, Heading4, ListParagraph, Grammar, Code, TableCellNormal, TableLineBefore, TableLineAfter, Note
* Table styles: `TableGrid`
* Character styles: `CodeEmbedded, Hyperlink`

## Headings and TOC

```
## Subsection

## `<c>`
```

This tool accepts markdown headings up to depth `####`. It renders them with paragraph styles `Heading1 ... Heading4`.

If the template.docx contains a table of contents, then it will be replaced with one using paragraph styles `TOC1 ... TOC2` synthesized from all one- and two-level headings.

## Lists

```
1. Item1
   * Item1a
   * Item1b
2. Item2
   * Item2a
```

This tool accepts markdown lists up to depth 4, and renders them in paragraph style `ListParagraph` with the stipulated combination of bullets/numbering. Like github-markdown, it starts each numbering sequence at `1` regardless of what number is actually present in the markdown.

## Tables

```
|    |    |
|:--:|:--:|
| a  | b  |
```

This tool accepts markdown tables using github-markdown conventions. It requires them to have pipes and a non-empty title row.

It renders the table in table style `TableGrid`, and renders each cell in paragraph style `TableCellNormal`. It emits an empty paragraph before the table with style `TableLineBefore`, and an empty paragraph after the table with style `TableLineAfter`.


## Grammar and code

```
    ```csharp
	x += await T.f();
	```

	```vb
	Dim x = Function(x As Integer) x+1
	```

	```antlr
	expr
	   : expr '+' expr
	   ;
	```
```

C#, VB and ANTLR code blocks are colorized. The C# and VB colorizer is better than most other colorizes because it's powered by Roslyn, not by regex.

If there is any `*.g4` present in the working directory where this tool is run, it validates that all productions defined inside ```antlr blocks in markdown are also found in the .g4 file, and vice versa. It generates a `.html` file of the same name as the `g4` file, which contains a hyperlinked version of the grammar. This hyperlinked version contains internal hyperlinks, and also external hyperlinks to the markdown section where each grammar production is defined.

### Code idiosyncracies

To put code inside bullets, you need newlines around the code:
```
*   Bullet1
*   Bullet2

    ```csharp
    x += y;
    ```

    More discussion

*   Bullet3
```

This tool doesn't handle code inside blockquotes.

There isn't a nice way to combine formatting and code...
```
`x `*op*` = y`     <-- renders badly on github
`x op= y`          <-- renders nicely

x<sub>1</sub>      <-- don't do this
x1...xN            <-- this is a good way to do subscripts
```




## Cross-referencing and hyperlinks

### Headings

```
See operator overloading for more details ([Operator overloading](expressions.md#operator-overloading))
```

All markdown links of the form `[title](file.md#heading)` are validated to ensure that (1) the `title` matches the `heading` and (2) that file does indeed contain a heading named `heading`.

The `heading` is obtained by preserving hyphens, underscores, numerics, alphas (converting them to lowercase), spaces (converting them to hyphens) and stripping out everything else.

Heading links are rendered in the word document as numeric section references `$10.1.2` that hyperlink to the section.

__Http__

```
See the [ECMA spec](http://www.ecma.org/123) for more details.
```

All markdown links of the form `[title](http...)` require title to be only plain text or inline code. They are rendered with text style `Hyperlink`, and link to that URL.

__Terms__

```
A ***term*** is a word that's defined in one place. Subsquent uses of the term will hyperlink back to it.
```

A term is defined with bold+italic (three asterisks or three underscores). Any subsequent mention of the term will be rendered with a dotted underline in color `#4BACC6` that hyperlinks to the term definition.

__Productions__

```
   ```antlr
   expr : expr '+' expr;
   ```

   An *expr* is formed out of one and another.
```

A production is defined inside an ```antlr block. Any subsquent mention of the production name in italics `*production*` or `_production_` is underlined and rendered in color `#6A5ACD` and hyperlinks to the antlr block which defined it.





