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
