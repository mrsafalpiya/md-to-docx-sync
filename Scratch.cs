using md_to_docx_sync.Enums;
using md_to_docx_sync.Interfaces;
using md_to_docx_sync.Services;

public class Scratch
{
    public static IDocxContent Test()
    {
        /**
         * Input markdown file: sample/input.md
        */

        DocxContent docxContent = new();

        /**
         * # Project Example: Markdown Features
         *
         * This document demonstrates a variety of Markdown features for testing conversion and syncing.
         */
        docxContent.NewHeading(1);
        docxContent.AddText("Project Example: Markdown Features");
        docxContent.NewParagraph();
        docxContent.AddText("This document demonstrates a variety of Markdown features for testing conversion and syncing.");

        /**
         * ## Overview
         *
         * Contents within this section explain purpose and include lists and inline formatting.
         *
         * - Unordered item one
         *   - Unordered item one-one
         *   - Unordered item one-two
         * - Unordered item two
         * - Unordered item three
         *   - Unordered item three-one
         * 
         * Foo bar buzz
         *
         * 1. First numbered item
         * 2. Second numbered item
         * 3. Third numbered item
         *
         * Some inline styles: **bold text**, _italic text_, <u>underlined text</u>, **_bold-italic text_**, ~~strike-through~~ and `literal text`.
         */
        docxContent.NewHeading(2);
        docxContent.AddText("Overview");
        docxContent.NewParagraph();
        docxContent.AddText("Contents within this section explain purpose and include lists and inline formatting.");

        docxContent.NewList(DocxListType.Bulleted);
        docxContent.NewListItem();
        docxContent.AddText("Unordered item one");
        docxContent.NewListItem(level: 1);
        docxContent.AddText("Unordered item one-one");
        docxContent.NewListItem(level: 1);
        docxContent.AddText("Unordered item one-two");
        docxContent.NewListItem();
        docxContent.AddText("Unordered item two");
        docxContent.NewListItem();
        docxContent.AddText("Unordered item three");
        docxContent.NewListItem(level: 1);
        docxContent.AddText("Unordered item three-one");

        docxContent.NewParagraph();
        docxContent.AddText("Foo bar buzz");

        docxContent.NewList(DocxListType.Numbered);
        docxContent.NewListItem();
        docxContent.AddText("First numbered item");
        docxContent.NewListItem();
        docxContent.AddText("Second numbered item");
        docxContent.NewListItem();
        docxContent.AddText("Third numbered item");

        docxContent.NewParagraph();
        docxContent.AddText("Some inline styles: ");
        docxContent.AddText("bold text", isBold: true);
        docxContent.AddText(", ");
        docxContent.AddText("italic text", isItalic: true);
        docxContent.AddText(", ");
        docxContent.AddText("underlined text", isUnderline: true);
        docxContent.AddText(", ");
        docxContent.AddText("bold-italic text", isBold: true, isItalic: true);
        docxContent.AddText(", ");
        docxContent.AddText("strike-through", isStrikethrough: true);
        docxContent.AddText(" and ");
        docxContent.AddText("literal text", isLiteral: true);
        docxContent.AddText(".");
        /**
         * ### Details (Level 3)
         *
         * This subsection contains its own lists and a short code example.
         * 
         * - Sub-**item A**
         * - Sub-item B
         * 
         * 1. Sub-number 1
         * 2. *Sub-number* 2
         * 
         * ```bash
         * # Example shell snippet
         * echo "Hello, Markdown!"
         * ```
         */
        docxContent.NewHeading(3);
        docxContent.AddText("Details (Level 3)");
        docxContent.NewParagraph();
        docxContent.AddText("This subsection contains its own lists and a short code example.");

        docxContent.NewList(DocxListType.Bulleted);
        docxContent.NewListItem();
        docxContent.AddText("Sub-");
        docxContent.AddText("item A", isBold: true);
        docxContent.NewListItem();
        docxContent.AddText("Sub-item B");

        docxContent.NewList(DocxListType.Numbered);
        docxContent.NewListItem();
        docxContent.AddText("Sub-number 1");
        docxContent.NewListItem();
        docxContent.AddText("Sub-number", isItalic: true);
        docxContent.AddText(" 2");

        docxContent.NewParagraph(isLiteral: true);
        docxContent.AddText("# Example shell snippet\necho \"Hello, Markdown!\"");
        docxContent.AddLineBreak();

        /**
         * Image
         */
        docxContent.AddImage("sample/placeholder-500.png", caption: "Sample placeholder image", widthInches: 3.0);

        /**
         * ### Small Table
         * 
         * | Header 1 | Header 2 |
         * | -------- | -------- |
         * |  Cell 1  |  Cell 2  |
         * |  Cell 3  |  Cell 4  |
         * |  Cell 5  |  Cell 6  |
         */
        docxContent.NewHeading(3);
        docxContent.AddText("Small Table");

        docxContent.NewTable("Sample Table", widthPercent: 50);
        docxContent.NewRow();
        docxContent.NewCell();
        docxContent.AddText("Header 1");
        docxContent.NewCell();
        docxContent.AddText("Header 2");
        docxContent.NewRow();
        docxContent.NewCell();
        docxContent.AddText("Cell 1");
        docxContent.NewCell();
        docxContent.AddText("Cell 2");
        docxContent.NewRow();
        docxContent.NewCell();
        docxContent.AddText("Cell 3");
        docxContent.NewCell();
        docxContent.AddText("Cell 4");
        docxContent.NewRow();
        docxContent.NewCell();
        docxContent.AddText("Cell 5");
        docxContent.NewCell();
        docxContent.AddText("Cell 6");

        return docxContent;
    }
}