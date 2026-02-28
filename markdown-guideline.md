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
| Image width | `^^^ Caption \|> 300` | Pixels, converted at 96 DPI |
| Table width | `^^^ Caption \|> 50` | Percentage of page width |
| Underline | `<u>text</u>` | HTML inline tags |
| Strikethrough | `~~text~~` | GFM extension |
| Smart quotes | `"text"` / `'text'` | Automatic |
| En-dash | `--` | Automatic |
| Em-dash | `---` | Automatic |
| Ellipsis | `...` | Automatic |
| Max heading level | `####` | Levels 5--6 clamp to 4 |
| Max list depth | 4 indents | Levels beyond 3 clamp to 3 |
