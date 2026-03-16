# Markdown Syntax Guidelines

This document covers syntax that is **in addition to or different from** standard CommonMark Markdown. Standard features (headings, paragraphs, bold with `**`, italic with `_`, inline code, fenced code blocks, unordered/ordered lists, tables, horizontal rules) work as expected and are not repeated here.

These extensions come from the Markdig library's advanced extensions, SmartyPants, and custom processing built into this converter.

---

## Figure Container (`^^^`)

The figure container is the primary custom syntax. It wraps an image or table with an **auto-numbered caption**. The converter uses Markdig's Figures extension (`^^^` delimiters) with custom caption and width handling.

### Image with caption

```markdown
^^^
![](./path/to/image.png)
^^^ Caption text
```

This produces a centered image with a caption below it reading **"Figure 1: Caption text"**. The figure number auto-increments across the document.

### Table with caption

```markdown
^^^
| Header 1 | Header 2 |
| -------- | -------- |
| Cell 1   | Cell 2   |
^^^ Caption text
```

This produces a table with a caption above it reading **"Table 1: Caption text"**. The table number auto-increments independently from figure numbers.

### Important notes

- Images and tables **must** be wrapped in `^^^` to get captions and auto-numbering.
- A standalone table (without `^^^`) will render but will have **no caption and no numbering**.
- A standalone image paragraph (without `^^^`) is **skipped entirely** in the output.
- The converter warns when a figure container is missing a caption.

---

## Width Operator (`|>`)

The `|>` operator is appended to a caption to control the width of the contained image or table. This is a **completely custom syntax** not found in standard Markdown or Markdig.

### Image width (pixels)

```markdown
^^^
![](./image.png)
^^^ My image caption |> 300
```

The value (`300`) is in **pixels** and is converted to inches at 96 DPI. The image height is auto-calculated to preserve the aspect ratio.

### Table width (percentage)

```markdown
^^^
| Col A | Col B |
| ----- | ----- |
| 1     | 2     |
^^^ My table caption |> 50
```

The value (`50`) is a **percentage of page width** (1--100). If omitted, the table defaults to 100% width.

### Syntax rules

- Format: `Caption text |> number`
- There must be at least one space before `|>`.
- The number must be a positive integer.
- If `|>` is omitted, images use their original dimensions and tables use 100% width.

---

## Page Break (`:::pagebreak`)

Use a Markdig custom container with `pagebreak` info to insert a DOCX page break.

```markdown
:::pagebreak
:::
```

This produces a hard page break at that location in the output document.

### Syntax rules

- The marker must be its own container block (`:::pagebreak` followed by closing `:::`).
- The `pagebreak` info value is case-insensitive.
- No visible content is rendered for this block; it only inserts a page break.

---

## Empty Paragraph (`&nbsp;`)

To insert an empty paragraph (a blank line) that isn't collapsed, use the non-breaking space HTML entity:

```markdown
&nbsp;
```

This is useful for adding vertical whitespace between elements when a standard double-newline isn't enough or when you need to ensure a paragraph exists with a specific style even if it's empty.

---

## Underline (`<u>`)

Standard Markdown has no underline syntax. This converter supports underline through inline HTML tags:

```markdown
<u>underlined text</u>
```

Underline can be combined with other inline formatting:

```markdown
**<u>bold and underlined</u>**
_<u>italic and underlined</u>_
```

---

## Strikethrough (`~~`)

Strikethrough uses the GitHub Flavored Markdown extension provided by Markdig:

```markdown
~~struck-through text~~
```

Requires exactly two tildes on each side.

---

## Formatted Code Blocks (`fmt`)

By default, content inside fenced code blocks is treated as plain literal text with no inline formatting. The `fmt` language identifier enables Markdown inline formatting to be parsed and applied inside a code-styled block.

The block still uses the `Code` paragraph style, but the content is processed for the specified formatting types instead of being output verbatim.

### Enable all formatting

Using `fmt` alone enables **all** inline formatting types (bold, italic, underline, strikethrough):

````markdown
```fmt
This is **bold**, _italic_, <u>underlined</u>, and ~~struck through~~.
A second line with **nested _bold italic_**.
```
````

### Enable specific formatting types

Append a colon and a comma-separated list of formatting types to restrict which types are active:

````markdown
```fmt:bold,italic
This is **bold** and _italic_, but ~~this~~ is not struck through.
```
````

````markdown
```fmt:underline
Only <u>underline</u> works here; **bold** and _italic_ are plain text.
```
````

### Available options

| Option | Effect |
| ------ | ------ |
| `bold` | Enables bold (`**text**`) |
| `italic` | Enables italic (`_text_` or `*text*`) |
| `underline` | Enables underline (`<u>text</u>`) |
| `strikethrough` | Enables strikethrough (`~~text~~`) |

Multiple options are combined with commas: `fmt:bold,italic,underline`.

### Important notes

- Unrecognised option names emit a warning and are ignored.
- Each line in the block is parsed as an independent inline context; block-level constructs (headings, lists, etc.) inside the block are not rendered.
- Lines are separated by line breaks within a single `Code`-styled paragraph, matching the normal code block layout.
- Smart typography (curly quotes, em-dashes, etc.) is always active because it is applied at the Markdig pipeline level.
- When a formatting type is **disabled**, its delimiter characters appear literally in the output. For example, with `fmt:italic`, writing `**bold**` outputs the literal text `**bold**` (stars included), and `~~strike~~` outputs `~~strike~~` (tildes included). Likewise, `<u>text</u>` outputs the literal HTML tags when underline is disabled. This lets you write the raw syntax characters without them being silently dropped.

---

## Smart Typography (SmartyPants)

The converter automatically transforms plain-text punctuation into typographic equivalents. This happens transparently -- you write standard ASCII characters and they become proper Unicode glyphs in the DOCX output.

| You write    | Output       | Character        |
| ------------ | ------------ | ---------------- |
| `"text"`     | \u201Ctext\u201D     | Curly double quotes |
| `'text'`     | \u2018text\u2019     | Curly single quotes |
| `...`        | \u2026           | Ellipsis           |
| `--`         | \u2013           | En-dash            |
| `---`        | \u2014           | Em-dash            |

These conversions are automatic and cannot be disabled per-instance. If you need a literal `--` or `...` in the output, use a fenced code block or inline code.

---

## References and Bibliography

The converter supports Word-compatible bibliography references using a fenced `references` code block for source definitions and `[@key]` for inline citations. When the document is opened in Word, CITATION and BIBLIOGRAPHY fields are rendered using Word's built-in bibliography engine.

### Defining sources

Sources are defined in a fenced code block with `references` as the language identifier. The content is YAML -- a list of entries, each with a `key`, `type`, and type-specific fields.

````markdown
```references
- key: smith2006
  type: Book
  author: Smith, John D; Doe, Jane
  title: How to Write Bibliographies
  year: 2006
  city: Chicago
  publisher: Adventure Works Press
  edition: 2nd

- key: chen2020
  type: JournalArticle
  author: Chen, Jacky
  title: Modern Citation Systems
  journalName: Adventure Works Monthly
  year: 2020
  pages: 50-62
  volume: III
  issue: 12
```
````

The `references` block does not produce visible output in the DOCX body. The sources are stored in the document's bibliography data (CustomXmlPart), and the template's existing bibliography/references section renders them. Inline `[@key]` citations are replaced with Word CITATION fields that link to these sources.

### Inline citations

Use `[@key]` to cite a source inline. Multiple citations can be combined with semicolons.

```markdown
This was demonstrated in prior work [@smith2006].
Several studies support this [@smith2006; @chen2020].
```

### Source types

The `type` field maps to Word's built-in bibliography source types. The following types are supported:

| Type value | Word source type |
| ---------- | ---------------- |
| `Book` | Book |
| `BookSection` | Book Section |
| `JournalArticle` | Journal Article |
| `ArticleInAPeriodical` | Article in a Periodical |
| `ConferenceProceedings` | Conference Proceedings |
| `Report` | Report |
| `WebSite` | Web site |
| `DocumentFromWebSite` | Document From Web site |
| `ElectronicSource` | Electronic Source |
| `Art` | Art |
| `SoundRecording` | Sound Recording |
| `Performance` | Performance |
| `Film` | Film |
| `Interview` | Interview |
| `Patent` | Patent |
| `Case` | Case |
| `Misc` | Miscellaneous |

Unrecognized type values default to `Misc`.

### Fields per source type

Each source type supports a different set of fields. All fields are optional except `key` and `type`. Person fields (author, editor, etc.) use the format `Last, First` with multiple people separated by semicolons.

#### Book

```yaml
- key: example
  type: Book
  author: Last, First; Last2, First2
  title: Book Title
  year: 2006
  city: Chicago
  publisher: Publisher Name
  edition: 2nd
```

#### Book Section

```yaml
- key: example
  type: BookSection
  author: Last, First
  title: Chapter Title
  bookTitle: Book Title
  year: 2006
  pages: 50-62
  city: Chicago
  publisher: Publisher Name
  editor: Last, First
```

#### Journal Article

```yaml
- key: example
  type: JournalArticle
  author: Last, First
  title: Article Title
  journalName: Journal Name
  year: 2006
  pages: 50-62
  volume: III
  issue: 12
```

#### Article in a Periodical

```yaml
- key: example
  type: ArticleInAPeriodical
  author: Last, First
  title: Article Title
  periodicalTitle: Periodical Name
  year: 2006
  month: January
  day: 1
  pages: 50-62
```

#### Conference Proceedings

```yaml
- key: example
  type: ConferenceProceedings
  author: Last, First
  title: Paper Title
  year: 2006
  city: Chicago
  publisher: Publisher Name
```

#### Report

```yaml
- key: example
  type: Report
  author: Last, First
  title: Report Title
  year: 2006
  publisher: Publisher Name
  city: Chicago
```

#### Web site

```yaml
- key: example
  type: WebSite
  author: Last, First
  nameOfWebPage: Page Title
  yearAccessed: 2024
  monthAccessed: January
  dayAccessed: 15
  url: https://www.example.com
```

#### Document From Web site

```yaml
- key: example
  type: DocumentFromWebSite
  author: Last, First
  nameOfWebPage: Page Title
  year: 2006
  month: January
  day: 1
  yearAccessed: 2024
  monthAccessed: January
  dayAccessed: 15
  url: https://www.example.com
```

#### Electronic Source

```yaml
- key: example
  type: ElectronicSource
  author: Last, First
  title: Source Title
  city: Chicago
  publisher: Publisher Name
  year: 2006
```

#### Art

```yaml
- key: example
  type: Art
  artist: Last, First
  title: Artwork Title
  institution: Gallery Name
  year: 2006
```

#### Sound Recording

```yaml
- key: example
  type: SoundRecording
  composer: Last, First
  title: Recording Title
  productionCompany: Company Name
  year: 2006
```

#### Performance

```yaml
- key: example
  type: Performance
  writer: Last, First
  title: Performance Title
  productionCompany: Company Name
  city: Chicago
  year: 2006
```

#### Film

```yaml
- key: example
  type: Film
  director: Last, First
  title: Film Title
  productionCompany: Company Name
  countryRegion: United States of America
  year: 2006
```

#### Interview

```yaml
- key: example
  type: Interview
  interviewee: Last, First
  title: Interview Title
  year: 2006
  month: January
  day: 1
```

#### Patent

```yaml
- key: example
  type: Patent
  inventor: Last, First
  title: Patent Title
  year: 2006
  countryRegion: United States of America
  patentNumber: 123 456
```

#### Case

```yaml
- key: example
  type: Case
  title: Case Title
  reporter: Supreme Court Reporter
  year: 2006
```

#### Miscellaneous

```yaml
- key: example
  type: Misc
  author: Last, First
  title: Title
  year: 2006
  city: Chicago
  publisher: Publisher Name
```

### Person field names by role

Different source types use different person roles. Use the appropriate field name:

| Field name | Used by |
| ---------- | ------- |
| `author` | Book, BookSection, JournalArticle, ArticleInAPeriodical, ConferenceProceedings, Report, WebSite, DocumentFromWebSite, ElectronicSource, Misc |
| `editor` | BookSection |
| `artist` | Art |
| `composer` | SoundRecording |
| `director` | Film |
| `writer` | Performance |
| `interviewee` | Interview |
| `inventor` | Patent |

### Important notes

- The `key` must be unique across all sources and match the keys used in `[@key]` citations.
- The `year` field can be a number or string in YAML; it is treated as a string in the output.
- Word will render the bibliography and format citations according to the selected style (defaults to APA 6th Edition). You can change the style in Word after opening the document.
- The document will prompt to update fields when opened; accept this to ensure citations and bibliography render correctly.

---

## Constraints and Limits

### Heading levels: 1--4 only

Only `#` through `####` are supported. Heading levels 5 and 6 (`#####`, `######`) are clamped to level 4 in the output.

```markdown
# Heading 1
## Heading 2
### Heading 3
#### Heading 4
```

### List nesting: 4 levels (0--3)

Lists support up to 4 levels of nesting. Deeper nesting is clamped to level 3.

```markdown
- Level 0
  - Level 1
    - Level 2
      - Level 3 (maximum)
```

Bulleted and numbered lists can be mixed at different nesting levels within the same list structure.

---

## Horizontal Rules

Horizontal rules (`---`) are recognized but **produce no visible output** in the DOCX. They can be used as logical separators in the source Markdown without affecting the generated document.

---

## Quick Reference

| Feature | Syntax | Notes |
| ------- | ------ | ----- |
| Figure with caption | `^^^ ... ^^^ Caption` | Auto-numbered as "Figure N:" |
| Table with caption | `^^^ ... ^^^ Caption` | Auto-numbered as "Table N:" |
| Page break | `:::pagebreak ... :::` | Inserts a hard page break |
| Empty Paragraph | `&nbsp;` | Inserts a non-empty blank line |
| Image width | `^^^ Caption \|> 300` | Pixels, converted at 96 DPI |
| Table width | `^^^ Caption \|> 50` | Percentage of page width |
| Inline citation | `[@key]` | Word CITATION field |
| Multiple citations | `[@key1; @key2]` | Semicolon-separated |
| References block | `` ```references `` | YAML source definitions |
| Underline | `<u>text</u>` | HTML inline tags |
| Strikethrough | `~~text~~` | GFM extension |
| Formatted code block (all) | ` ```fmt ` | All inline formatting enabled |
| Formatted code block (select) | ` ```fmt:bold,italic ` | Only specified types enabled |
| Smart quotes | `"text"` / `'text'` | Automatic |
| En-dash | `--` | Automatic |
| Em-dash | `---` | Automatic |
| Ellipsis | `...` | Automatic |
| Max heading level | `####` | Levels 5--6 clamp to 4 |
| Max list depth | 4 indents | Levels beyond 3 clamp to 3 |
