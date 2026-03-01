# Project Example: Markdown Features

This document demonstrates a variety of Markdown features for testing conversion and syncing.

## Overview

Contents within this "section explain purpose" and include 'lists and' inline 'formatting.

- Unordered item one
  - Unordered item one-one
  - Unordered item one-two
- Unordered item -- two
- Unordered item --- three
  - Unordered item ... three-one

1. First numbered item
2. Second numbered item
3. Third numbered item

Some inline styles: **bold text**, _italic text_, <u>underlined text</u>, **_bold-italic text_**, ~~strike-through~~ and `literal text`.

### Details (Level 3)

This subsection contains its own lists and a short code example.

- Sub-item A
- Sub-item B

1. Sub-number 1
2. Sub-number 2

```bash
# Example shell snippet
echo "Hello, Markdown!"
```

#### Notes (Level 4)

Short note content inside a level 4 heading.

- Note bullet one
- Note bullet two

1. Note step one
2. Note step two

---

# Usage Examples

Another top-level section to show repeated headings and additional constructs.

## Installation

Instructions and lists for install.

- Download package
- Extract files

1. Run installer
2. Restart system

### Configuration

Configuration options with a small table and inline emphasis.

^^^
| Option | Description | Default |
| ------ | ----------- | ------- |
| theme | UI theme | light |
| sync | Auto-sync | true |
| level | Verbosity | 2 |
^^^ Configuration options table |> 50

### Image

^^^
![](./placeholder-500.png)
^^^ Example placeholder 500x500 image |> 300

^^^
![](./placeholder-400-300.png)
^^^ Example placeholder 400x300 image

#### Advanced (Level 4 repeat)

Advanced configuration details.

- Toggle `sync` to enable automatic syncing.
- Use **bold** to highlight important options.

```csharp
// Example C# snippet
using System;

namespace Example
{
	class Program
	{
		static void Main()
		{
			Console.WriteLine("Hello from C# code block.");
		}
	}
}
```

---

# API Reference

A repeated H1 to satisfy multiple entries requirement.

## Endpoints

Short description plus lists and formatting.

- GET /items — Retrieve items
- POST /items — Create item

1. Authenticate
2. Call endpoint

### Request Example

Example request body and inline emphasis: **required** fields are _id_ and _name_.

```json
{
  "id": 123,
  "name": "Sample Item",
  "active": true
}
```

#### Response Notes

Explanation of typical responses.

- 200 OK — Successful
- 400 Bad Request — Invalid input

1. Parse response
2. Handle errors

---

# Images

^^^
![](./placeholder-500.png)
^^^ Sample placeholder image |> 400

---

# Appendix

Final H1 section with mixed inline styles and another small table.

## Formatting Examples

- **Bold example**
- _Italic example_
- <u>Underline example</u>
- **_Bold-italic example_**

### Small Table

^^^
| Name | Type |
| ---- | ---- |
| foo | int |
| bar | str |
^^^ Small example table

#### Final Codeblock

```python
# Simple Python example
def greet(name):
	return f"Hello, {name}!"

print(greet('World'))
```

End of example content.

---

# References

## Inline Citations

This section demonstrates inline citation syntax. The REST API design principles used here follow established guidelines [@fielding2000]. Configuration management best practices are documented in several sources [@humble2010; @fowler2018].

The markdown processing pipeline is built on the Markdig library [@markdig2024], which supports CommonMark and a range of extensions. Image handling relies on the ImageSharp library for cross-platform dimension retrieval [@sixlabors2023].

For further background on document generation workflows, see the official OpenXML specification [@microsoftOpenXml] and the related web resource [@openxmlWebSite2024].

A case study on tooling for technical writing is also referenced here [@techWritingCase2019].

## References

```references
- key: fielding2000
  type: Book
  author: Fielding, Roy Thomas
  title: Architectural Styles and the Design of Network-based Software Architectures
  year: 2000
  city: Irvine
  publisher: University of California

- key: humble2010
  type: Book
  author: Humble, Jez; Farley, David
  title: Continuous Delivery
  year: 2010
  city: Boston
  publisher: Addison-Wesley Professional
  edition: 1st

- key: fowler2018
  type: BookSection
  author: Fowler, Martin
  title: Configuration Patterns
  bookTitle: Refactoring -- Improving the Design of Existing Code
  year: 2018
  pages: 231-256
  city: Boston
  publisher: Addison-Wesley Professional
  editor: Fowler, Martin

- key: markdig2024
  type: WebSite
  author: Bertels, Alexandre
  nameOfWebPage: Markdig -- A Fast, Powerful, CommonMark Compliant Markdown Processor
  yearAccessed: 2026
  monthAccessed: March
  dayAccessed: 1
  url: https://github.com/xoofx/markdig

- key: sixlabors2023
  type: WebSite
  nameOfWebPage: ImageSharp -- A Modern, Cross-Platform, 2D Graphics Library
  yearAccessed: 2026
  monthAccessed: March
  dayAccessed: 1
  url: https://sixlabors.com/products/imagesharp

- key: microsoftOpenXml
  type: DocumentFromWebSite
  author: Microsoft Corporation
  nameOfWebPage: Open XML SDK Documentation
  year: 2023
  month: November
  day: 1
  yearAccessed: 2026
  monthAccessed: March
  dayAccessed: 1
  url: https://learn.microsoft.com/en-us/office/open-xml/open-xml-sdk

- key: openxmlWebSite2024
  type: WebSite
  nameOfWebPage: ECMA-376 Office Open XML File Formats
  yearAccessed: 2026
  monthAccessed: March
  dayAccessed: 1
  url: https://www.ecma-international.org/publications-and-standards/standards/ecma-376

- key: techWritingCase2019
  type: JournalArticle
  author: Ganier, Franck; de Jong, Menno
  title: Technological Developments in Document and Information Design
  journalName: Journal of Technical Writing and Communication
  year: 2019
  pages: 3-22
  volume: 49
  issue: 1
```
