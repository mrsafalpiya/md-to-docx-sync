using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig;
using Markdig.Syntax;
using md_to_docx_sync.Enums;
using md_to_docx_sync.Models;
using md_to_docx_sync.Services;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

// Get command line arguments

string? inputMDFilePath = args.ElementAtOrDefault(0);
if (string.IsNullOrEmpty(inputMDFilePath))
{
    Console.Error.WriteLine("[ERROR] Please provide the input MD file path as the first argument.");
    return;
}

string? inputDOCXTemplateFilePath = args.ElementAtOrDefault(1);
if (string.IsNullOrEmpty(inputDOCXTemplateFilePath))
{
    Console.Error.WriteLine("[ERROR] Please provide the input DOCX template file path as the second argument.");
    return;
}

string? outputDOCXFilePath = args.ElementAtOrDefault(2);
if (string.IsNullOrEmpty(outputDOCXFilePath))
{
    Console.Error.WriteLine("[ERROR] Please provide the output DOCX file path as the third argument.");
    return;
}

// Get the directory containing the markdown file for resolving relative paths
string? inputMDFileDir = Path.GetDirectoryName(Path.GetFullPath(inputMDFilePath));

// Read and parse Markdown File
var pipeline = new MarkdownPipelineBuilder().UseAdvancedExtensions().UseSmartyPants().Build();

// Pre-process: replace [@key] / [@key1; @key2] with «key» / «key1;key2» placeholders
// so Markdig's link parser doesn't consume the brackets and break citation detection.
var rawMarkdown = File.ReadAllText(inputMDFilePath);
var citationPreProcess = new Regex(@"\[@([a-zA-Z0-9_:-]+(?:\s*;\s*@[a-zA-Z0-9_:-]+)*)\]");
rawMarkdown = citationPreProcess.Replace(rawMarkdown, match =>
{
    // Strip the leading @ from each key and join with ;
    var keys = match.Groups[1].Value
        .Split(';', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)
        .Select(k => k.TrimStart('@'));
    return $"\u00AB{string.Join(";", keys)}\u00BB";
});

var markdownContent = Markdown.Parse(rawMarkdown, pipeline);

// Create DocxContent from markdown
DocxContent docxContent = new();

// Helper to parse caption with optional width (e.g., "Caption text |> 400")
(string caption, int? width) ParseCaptionWithWidth(string? rawCaption)
{
    if (string.IsNullOrEmpty(rawCaption))
        return (string.Empty, null);

    var match = Regex.Match(rawCaption, @"^(.+?)\s*\|>\s*(\d+)\s*$");
    if (match.Success)
    {
        return (match.Groups[1].Value.Trim(), int.Parse(match.Groups[2].Value));
    }
    return (rawCaption.Trim(), null);
}

// Helper to get plain text from inline elements (for captions)
string GetPlainText(Markdig.Syntax.Inlines.ContainerInline? container)
{
    if (container == null) return string.Empty;
    var sb = new System.Text.StringBuilder();
    foreach (var inline in container)
    {
        if (inline is Markdig.Syntax.Inlines.LiteralInline literal)
            sb.Append(literal.Content);
        else if (inline is Markdig.Syntax.Inlines.ContainerInline nested)
            sb.Append(GetPlainText(nested));
    }
    return sb.ToString();
}

// Helper to parse references YAML block into ReferenceSource list
List<ReferenceSource> ParseReferencesYaml(string yamlContent)
{
    var deserializer = new DeserializerBuilder()
        .WithNamingConvention(CamelCaseNamingConvention.Instance)
        .Build();

    var yamlData = deserializer.Deserialize<List<Dictionary<string, object>>>(yamlContent);
    if (yamlData == null) return new List<ReferenceSource>();

    var references = new List<ReferenceSource>();
    foreach (var entry in yamlData)
    {
        if (!entry.TryGetValue("key", out var keyObj) || !entry.TryGetValue("type", out var typeObj))
        {
            Console.WriteLine("[WARNING] Reference entry missing 'key' or 'type' field, skipping.");
            continue;
        }

        var reference = new ReferenceSource
        {
            Key = keyObj.ToString()!,
            Type = typeObj.ToString()!
        };

        foreach (var (fieldName, fieldValue) in entry)
        {
            if (fieldName == "key" || fieldName == "type") continue;
            reference.Fields[fieldName] = fieldValue?.ToString() ?? string.Empty;
        }

        references.Add(reference);
    }

    return references;
}

// Regex for detecting «key1;key2» citation placeholders (produced by pre-processing step)
var citationPattern = new Regex(@"\u00AB([a-zA-Z0-9_;:-]+)\u00BB");

// Helper to process literal text that may contain citation placeholders «key1;key2»
void AddTextWithCitations(string text, bool isBold, bool isItalic, bool isUnderline, bool isStrikethrough, bool isLiteral)
{
    int lastIndex = 0;
    foreach (Match match in citationPattern.Matches(text))
    {
        // Add text before the citation, stripping any trailing space
        // (Word's citation field rendering already provides the necessary spacing)
        if (match.Index > lastIndex)
        {
            var beforeText = text.Substring(lastIndex, match.Index - lastIndex).TrimEnd(' ');
            if (beforeText.Length > 0)
                docxContent.AddText(beforeText, isBold, isItalic, isUnderline, isStrikethrough, isLiteral);
        }

        // Parse citation keys from placeholder (keys are ;-separated)
        var citationKeys = match.Groups[1].Value
            .Split(';', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
        docxContent.AddCitation(citationKeys);

        lastIndex = match.Index + match.Length;
    }

    // Add remaining text after last citation
    if (lastIndex < text.Length)
    {
        docxContent.AddText(text.Substring(lastIndex), isBold, isItalic, isUnderline, isStrikethrough, isLiteral);
    }
}

// Process inline content with formatting
void ProcessInlines(Markdig.Syntax.Inlines.ContainerInline? container, bool isBold = false, bool isItalic = false, bool isUnderline = false, bool isStrikethrough = false,
    bool allowBold = true, bool allowItalic = true, bool allowUnderline = true, bool allowStrikethrough = true)
{
    if (container == null) return;

    foreach (var inline in container)
    {
        if (inline is Markdig.Syntax.Inlines.LiteralInline literal)
        {
            AddTextWithCitations(literal.Content.ToString(), isBold, isItalic, isUnderline, isStrikethrough, false);
        }
        else if (inline is Markdig.Syntax.Inlines.EmphasisInline emphasis)
        {
            bool newBold = isBold;
            bool newItalic = isItalic;
            bool newStrikethrough = isStrikethrough;
            bool fmtDisabled = false;
            string delim = new string(emphasis.DelimiterChar, emphasis.DelimiterCount);

            if (emphasis.DelimiterChar == '~' && emphasis.DelimiterCount == 2)
            {
                if (allowStrikethrough) newStrikethrough = true;
                else fmtDisabled = true;
            }
            else if (emphasis.DelimiterCount == 2)
            {
                if (allowBold) newBold = true;
                else fmtDisabled = true;
            }
            else
            {
                if (allowItalic) newItalic = true;
                else fmtDisabled = true;
            }

            if (fmtDisabled) docxContent.AddText(delim, isBold, isItalic, isUnderline, isStrikethrough);
            ProcessInlines(emphasis, newBold, newItalic, isUnderline, newStrikethrough, allowBold, allowItalic, allowUnderline, allowStrikethrough);
            if (fmtDisabled) docxContent.AddText(delim, isBold, isItalic, isUnderline, isStrikethrough);
        }
        else if (inline is Markdig.Syntax.Inlines.CodeInline codeInline)
        {
            docxContent.AddText(codeInline.Content, isBold, isItalic, isUnderline, isStrikethrough, isLiteral: true);
        }
        else if (inline is Markdig.Extensions.SmartyPants.SmartyPant smartyPant)
        {
            string smartText = smartyPant.Type switch
            {
                Markdig.Extensions.SmartyPants.SmartyPantType.Quote => "\u2018",
                Markdig.Extensions.SmartyPants.SmartyPantType.LeftQuote => "\u2018",      // '
                Markdig.Extensions.SmartyPants.SmartyPantType.RightQuote => "\u2019",     // '
                Markdig.Extensions.SmartyPants.SmartyPantType.DoubleQuote => "\u201C",
                Markdig.Extensions.SmartyPants.SmartyPantType.LeftDoubleQuote => "\u201C", // "
                Markdig.Extensions.SmartyPants.SmartyPantType.RightDoubleQuote => "\u201D", // "
                Markdig.Extensions.SmartyPants.SmartyPantType.Ellipsis => "\u2026",       // …
                Markdig.Extensions.SmartyPants.SmartyPantType.Dash2 => "\u2013",          // –
                Markdig.Extensions.SmartyPants.SmartyPantType.Dash3 => "\u2014",          // —
                _ => smartyPant.ToString()
            };
            docxContent.AddText(smartText, isBold, isItalic, isUnderline, isStrikethrough);
        }
        else if (inline is Markdig.Syntax.Inlines.HtmlInline htmlInline)
        {
            // <u>/<\/u> are handled in ProcessInlinesWithUnderline; emit literally here
            docxContent.AddText(htmlInline.Tag, isBold, isItalic, isUnderline, isStrikethrough);
        }
        else if (inline is Markdig.Syntax.Inlines.LineBreakInline)
        {
            docxContent.AddLineBreak();
        }
        else if (inline is Markdig.Syntax.Inlines.ContainerInline nestedContainer)
        {
            ProcessInlines(nestedContainer, isBold, isItalic, isUnderline, isStrikethrough, allowBold, allowItalic, allowUnderline, allowStrikethrough);
        }
    }
}

// Process inline content with HTML underline tag detection
void ProcessInlinesWithUnderline(Markdig.Syntax.Inlines.ContainerInline? container, bool isBold = false, bool isItalic = false, bool isStrikethrough = false,
    bool allowBold = true, bool allowItalic = true, bool allowUnderline = true, bool allowStrikethrough = true)
{
    if (container == null) return;

    bool currentUnderline = false;
    var inlines = container.ToList();

    for (int i = 0; i < inlines.Count; i++)
    {
        var inline = inlines[i];

        if (inline is Markdig.Syntax.Inlines.HtmlInline htmlInline)
        {
            string tag = htmlInline.Tag.ToLowerInvariant();
            if (tag == "<u>")
            {
                if (allowUnderline) currentUnderline = true;
                else docxContent.AddText(htmlInline.Tag, isBold, isItalic, currentUnderline, isStrikethrough);
            }
            else if (tag == "</u>")
            {
                if (allowUnderline) currentUnderline = false;
                else docxContent.AddText(htmlInline.Tag, isBold, isItalic, currentUnderline, isStrikethrough);
            }
        }
        else if (inline is Markdig.Syntax.Inlines.LiteralInline literal)
        {
            AddTextWithCitations(literal.Content.ToString(), isBold, isItalic, currentUnderline, isStrikethrough, false);
        }
        else if (inline is Markdig.Syntax.Inlines.EmphasisInline emphasis)
        {
            bool newBold = isBold;
            bool newItalic = isItalic;
            bool newStrikethrough = isStrikethrough;
            bool fmtDisabled = false;
            string delim = new string(emphasis.DelimiterChar, emphasis.DelimiterCount);

            if (emphasis.DelimiterChar == '~' && emphasis.DelimiterCount == 2)
            {
                if (allowStrikethrough) newStrikethrough = true;
                else fmtDisabled = true;
            }
            else if (emphasis.DelimiterCount == 2)
            {
                if (allowBold) newBold = true;
                else fmtDisabled = true;
            }
            else
            {
                if (allowItalic) newItalic = true;
                else fmtDisabled = true;
            }

            if (fmtDisabled) docxContent.AddText(delim, isBold, isItalic, currentUnderline, isStrikethrough);
            ProcessInlines(emphasis, newBold, newItalic, currentUnderline, newStrikethrough, allowBold, allowItalic, allowUnderline, allowStrikethrough);
            if (fmtDisabled) docxContent.AddText(delim, isBold, isItalic, currentUnderline, isStrikethrough);
        }
        else if (inline is Markdig.Syntax.Inlines.CodeInline codeInline)
        {
            docxContent.AddText(codeInline.Content, isBold, isItalic, currentUnderline, isStrikethrough, isLiteral: true);
        }
        else if (inline is Markdig.Extensions.SmartyPants.SmartyPant smartyPant)
        {
            string smartText = smartyPant.Type switch
            {
                Markdig.Extensions.SmartyPants.SmartyPantType.Quote => "'",
                Markdig.Extensions.SmartyPants.SmartyPantType.LeftQuote => "\u2018",      // '
                Markdig.Extensions.SmartyPants.SmartyPantType.RightQuote => "\u2019",     // '
                Markdig.Extensions.SmartyPants.SmartyPantType.DoubleQuote => "\"",
                Markdig.Extensions.SmartyPants.SmartyPantType.LeftDoubleQuote => "\u201C", // "
                Markdig.Extensions.SmartyPants.SmartyPantType.RightDoubleQuote => "\u201D", // "
                Markdig.Extensions.SmartyPants.SmartyPantType.Ellipsis => "\u2026",       // …
                Markdig.Extensions.SmartyPants.SmartyPantType.Dash2 => "\u2013",          // –
                Markdig.Extensions.SmartyPants.SmartyPantType.Dash3 => "\u2014",          // —
                _ => smartyPant.ToString()
            };
            docxContent.AddText(smartText, isBold, isItalic, currentUnderline, isStrikethrough);
        }
        else if (inline is Markdig.Syntax.Inlines.LineBreakInline)
        {
            docxContent.AddLineBreak();
        }
        else if (inline is Markdig.Syntax.Inlines.ContainerInline nestedContainer)
        {
            ProcessInlines(nestedContainer, isBold, isItalic, currentUnderline, isStrikethrough, allowBold, allowItalic, allowUnderline, allowStrikethrough);
        }
    }
}

// Process list items recursively with nesting level tracking
void ProcessListBlock(ListBlock listBlock, int nestingLevel = 0)
{
    bool isOrdered = listBlock.IsOrdered;
    DocxListType listType = isOrdered ? DocxListType.Numbered : DocxListType.Bulleted;

    // Only create a new list at the top level
    if (nestingLevel == 0)
    {
        docxContent.NewList(listType);
    }

    foreach (var item in listBlock)
    {
        if (item is ListItemBlock listItemBlock)
        {
            foreach (var child in listItemBlock)
            {
                if (child is ParagraphBlock paragraphBlock)
                {
                    docxContent.NewListItem(level: nestingLevel);
                    ProcessInlinesWithUnderline(paragraphBlock.Inline);
                }
                else if (child is ListBlock nestedList)
                {
                    // Process nested list with increased nesting level
                    ProcessListBlock(nestedList, nestingLevel + 1);
                }
            }
        }
    }
}

void ProcessTableCellContent(Markdig.Extensions.Tables.TableCell tableCell)
{
    bool hasRenderedLine = false;

    void EnsureLineStart()
    {
        if (hasRenderedLine)
        {
            docxContent.AddLineBreak();
        }
        hasRenderedLine = true;
    }

    void ProcessTableCellBlock(Block block)
    {
        switch (block)
        {
            case ParagraphBlock paragraphBlock:
                EnsureLineStart();
                ProcessInlinesWithUnderline(paragraphBlock.Inline);
                break;

            case QuoteBlock quoteBlock:
                foreach (var quoteChild in quoteBlock)
                {
                    if (quoteChild is ParagraphBlock quotedParagraph)
                    {
                        EnsureLineStart();
                        docxContent.AddText("> ");
                        ProcessInlinesWithUnderline(quotedParagraph.Inline);
                    }
                    else
                    {
                        ProcessTableCellBlock(quoteChild);
                    }
                }
                break;

            case ListBlock listBlock:
                int orderedIndex = 1;
                foreach (var listItem in listBlock)
                {
                    if (listItem is not ListItemBlock listItemBlock) continue;

                    foreach (var listItemChild in listItemBlock)
                    {
                        if (listItemChild is ParagraphBlock listItemParagraph)
                        {
                            EnsureLineStart();
                            docxContent.AddText(listBlock.IsOrdered ? $"{orderedIndex}. " : "• ");
                            ProcessInlinesWithUnderline(listItemParagraph.Inline);
                        }
                        else
                        {
                            ProcessTableCellBlock(listItemChild);
                        }
                    }

                    if (listBlock.IsOrdered) orderedIndex++;
                }
                break;

            case ContainerBlock containerBlock:
                foreach (var childBlock in containerBlock)
                {
                    ProcessTableCellBlock(childBlock);
                }
                break;

            case LeafBlock leafBlock when leafBlock.Inline != null:
                EnsureLineStart();
                ProcessInlinesWithUnderline(leafBlock.Inline);
                break;
        }
    }

    foreach (var cellContent in tableCell)
    {
        ProcessTableCellBlock(cellContent);
    }
}

void ProcessTable(Markdig.Extensions.Tables.Table table, string? caption = null, int? widthPercent = null)
{
    docxContent.NewTable(caption, widthPercent);

    // Tracks pending vertical merges per column index (remaining continuation rows).
    var activeVerticalMerges = new Dictionary<int, int>();

    foreach (var row in table)
    {
        if (row is not Markdig.Extensions.Tables.TableRow tableRow)
        {
            continue;
        }

        docxContent.NewRow();
        int columnIndex = 0;

        foreach (var cell in tableRow)
        {
            if (cell is not Markdig.Extensions.Tables.TableCell tableCell)
            {
                continue;
            }

            while (activeVerticalMerges.TryGetValue(columnIndex, out int pendingContinuations) && pendingContinuations > 0)
            {
                docxContent.NewCell(verticalMergeContinue: true);
                pendingContinuations--;

                if (pendingContinuations == 0) activeVerticalMerges.Remove(columnIndex);
                else activeVerticalMerges[columnIndex] = pendingContinuations;

                columnIndex++;
            }

            int columnSpan = Math.Max(1, tableCell.ColumnSpan);
            int rowSpan = tableCell.RowSpan;

            bool isVerticalMergeStart = rowSpan > 1;
            bool isVerticalMergeContinue = rowSpan <= 0;

            docxContent.NewCell(
                columnSpan: columnSpan,
                verticalMergeStart: isVerticalMergeStart,
                verticalMergeContinue: isVerticalMergeContinue
            );

            if (isVerticalMergeStart)
            {
                int continuationRows = rowSpan - 1;
                for (int i = 0; i < columnSpan; i++)
                {
                    activeVerticalMerges[columnIndex + i] = continuationRows;
                }
            }

            if (!isVerticalMergeContinue)
            {
                ProcessTableCellContent(tableCell);
            }

            columnIndex += columnSpan;
        }

        while (activeVerticalMerges.TryGetValue(columnIndex, out int trailingContinuations) && trailingContinuations > 0)
        {
            docxContent.NewCell(verticalMergeContinue: true);
            trailingContinuations--;

            if (trailingContinuations == 0) activeVerticalMerges.Remove(columnIndex);
            else activeVerticalMerges[columnIndex] = trailingContinuations;

            columnIndex++;
        }
    }
}

// Main recursive processor for markdown nodes
void ProcessMarkdownNode(MarkdownObject node)
{
    switch (node)
    {
        case HeadingBlock headingBlock:
            docxContent.NewHeading(headingBlock.Level);
            ProcessInlinesWithUnderline(headingBlock.Inline);
            break;

        case ParagraphBlock paragraphBlock:
            // Check if this paragraph contains only an image link
            var firstInline = paragraphBlock.Inline?.FirstChild;
            if (firstInline is Markdig.Syntax.Inlines.LinkInline { IsImage: true } imageLink &&
                paragraphBlock.Inline?.LastChild == firstInline)
            {
                // This is a standalone image paragraph - it will be handled by the Figure container
                // Skip processing here as it's handled by the Figure case
                break;
            }

            docxContent.NewParagraph();
            ProcessInlinesWithUnderline(paragraphBlock.Inline);
            break;

        case ListBlock listBlock:
            ProcessListBlock(listBlock);
            break;

        case FencedCodeBlock fencedCodeBlock:
            // Check if this is a references block
            if (fencedCodeBlock.Info?.Trim() == "references")
            {

                var yamlContent = string.Join("\n", fencedCodeBlock.Lines);
                try
                {
                    var references = ParseReferencesYaml(yamlContent);
                    if (references.Count > 0)
                    {
                        // Remove the heading that preceded this references block
                        docxContent.RemoveLastHeading();
                        docxContent.AddBibliography(references);
                        Console.WriteLine($"[INFO] Parsed {references.Count} bibliography source(s).");
                    }
                    else
                    {
                        Console.WriteLine("[WARNING] References block found but contained no valid entries.");
                    }
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"[ERROR] Failed to parse references YAML: {ex.Message}");
                }
                break;
            }

            // Check for fmt: formatted code block
            var fmtInfo = fencedCodeBlock.Info?.Trim() ?? "";
            if (fmtInfo == "fmt" || fmtInfo.StartsWith("fmt:"))
            {
                bool allowBold = false, allowItalic = false, allowUnderline = false, allowStrikethrough = false;

                if (fmtInfo == "fmt")
                {
                    allowBold = allowItalic = allowUnderline = allowStrikethrough = true;
                }
                else
                {
                    var opts = fmtInfo.Substring(4).Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
                    foreach (var opt in opts)
                    {
                        switch (opt.ToLowerInvariant())
                        {
                            case "bold": allowBold = true; break;
                            case "italic": allowItalic = true; break;
                            case "underline": allowUnderline = true; break;
                            case "strikethrough": allowStrikethrough = true; break;
                            default:
                                Console.WriteLine($"[WARNING] Unknown fmt option '{opt}'. Valid options: bold, italic, underline, strikethrough.");
                                break;
                        }
                    }
                }

                docxContent.NewParagraph(isLiteral: true);
                var fmtLineGroup = fencedCodeBlock.Lines;
                int fmtLineCount = fmtLineGroup.Count;
                for (int fmtLineIdx = 0; fmtLineIdx < fmtLineCount; fmtLineIdx++)
                {
                    var lineText = fmtLineGroup.Lines[fmtLineIdx].Slice.ToString();

                    // Separate leading whitespace so Markdig doesn't interpret
                    // indented lines as indented code blocks when parsing inline content.
                    var trimmedLine = lineText.TrimStart();
                    var leadingWhitespace = lineText.Substring(0, lineText.Length - trimmedLine.Length);

                    if (leadingWhitespace.Length > 0)
                        docxContent.AddText(leadingWhitespace);

                    var lineDoc = Markdown.Parse(trimmedLine, pipeline);
                    if (lineDoc.FirstOrDefault() is ParagraphBlock fmtPara)
                    {
                        ProcessInlinesWithUnderline(fmtPara.Inline,
                            allowBold: allowBold, allowItalic: allowItalic,
                            allowUnderline: allowUnderline, allowStrikethrough: allowStrikethrough);
                    }
                    else
                    {
                        docxContent.AddText(trimmedLine);
                    }

                    if (fmtLineIdx < fmtLineCount - 1)
                        docxContent.AddLineBreak();
                }
                break;
            }

            var codeContent = string.Join("\n", fencedCodeBlock.Lines);
            docxContent.NewParagraph(isLiteral: true);
            docxContent.AddText(codeContent);
            break;

        case ThematicBreakBlock:
            // Horizontal rule - skip or could add a separator
            break;

        case Markdig.Extensions.Figures.Figure figure:
            // Figure can contain Table, ParagraphBlock (with image), and FigureCaption
            string? figureCaption = null;
            int? figureWidth = null;

            // First, find the caption to get width info
            foreach (var child in figure)
            {
                if (child is Markdig.Extensions.Figures.FigureCaption caption)
                {
                    // FigureCaption is a LeafBlock with Inline content
                    var rawCaption = GetPlainText(caption.Inline);
                    (figureCaption, figureWidth) = ParseCaptionWithWidth(rawCaption);
                }
            }

            // Now process the figure contents
            bool hasTable = false;
            bool hasImage = false;

            foreach (var child in figure)
            {
                if (child is Markdig.Extensions.Tables.Table table)
                {
                    hasTable = true;
                    ProcessTable(table, figureCaption, figureWidth);
                }
                else if (child is ParagraphBlock figParagraphBlock)
                {
                    // Check for image inside paragraph
                    foreach (var inline in figParagraphBlock.Inline ?? Enumerable.Empty<Markdig.Syntax.Inlines.Inline>())
                    {
                        if (inline is Markdig.Syntax.Inlines.LinkInline { IsImage: true } figImageLink)
                        {
                            hasImage = true;
                            string? imagePath = figImageLink.Url;
                            if (!string.IsNullOrEmpty(imagePath) && !string.IsNullOrEmpty(inputMDFileDir))
                            {
                                // Resolve relative path
                                imagePath = Path.Combine(inputMDFileDir, imagePath);
                            }

                            // Convert width from pixels to inches (assuming 96 DPI)
                            double? widthInches = figureWidth.HasValue ? figureWidth.Value / 96.0 : null;
                            docxContent.AddImage(imagePath ?? string.Empty, figureCaption, widthInches);
                        }
                    }
                }
                // FigureCaption already processed above
            }

            // Check for missing caption and print warning
            if (string.IsNullOrWhiteSpace(figureCaption))
            {
                if (hasTable)
                {
                    Console.WriteLine("[WARNING] Table is missing a caption.");
                }
                else if (hasImage)
                {
                    Console.WriteLine("[WARNING] Image is missing a caption.");
                }
            }
            break;

        case Markdig.Extensions.Tables.Table table:
            // Standalone table (not in a figure)
            ProcessTable(table);
            break;

        case LinkReferenceDefinitionGroup:
            // Skip link reference definitions
            break;

        case MarkdownDocument document:
            foreach (var child in document)
            {
                ProcessMarkdownNode(child);
            }
            break;

        case ContainerBlock containerBlock:
            foreach (var child in containerBlock)
            {
                ProcessMarkdownNode(child);
            }
            break;
    }
}

ProcessMarkdownNode(markdownContent);

// Copy DOCX Template to Output File

System.IO.File.Copy(inputDOCXTemplateFilePath, outputDOCXFilePath, true);

// Write to DOCX File

var doc = WordprocessingDocument.Open(outputDOCXFilePath, true);
if (doc.MainDocumentPart is null)
{
    Console.Error.WriteLine("[ERROR] The DOCX template does not contain a valid main document part.");
    return;
}
if (doc.MainDocumentPart.Document is null)
{
    Console.Error.WriteLine("[ERROR] The DOCX template does not contain a valid document.");
    return;
}

var body = doc.MainDocumentPart.Document.Body;
if (body is null)
{
    Console.Error.WriteLine("[ERROR] The DOCX template does not contain a valid body.");
    return;
}

var firstSectionEndPara = body
            .Descendants<Paragraph>()
            .FirstOrDefault(p => p.Descendants<SectionProperties>().Any());
if (firstSectionEndPara is null)
{
    Console.Error.WriteLine("[ERROR] The DOCX template does not contain a second section break.");
    return;
}

// var newPara = new Paragraph(new Run(new Text("Hello, World!")));
// body.InsertAfter(newPara, firstSectionEndPara);

docxContent.WriteTo(doc, firstSectionEndPara);

doc.Save();